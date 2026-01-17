import { addDays, startOfMonth, endOfMonth, eachDayOfInterval, isWeekend, format, parse } from 'date-fns';
import { Vendor, ClientVendorMapping, ProductionItem, ProductionPlan, ProcessType } from './types';
import { vendorNameMap } from '@/data/defaults';

// 이동일 문자열을 Date로 파싱 (다양한 형식 지원)
function parseTransferDate(dateStr: string | undefined, fallbackYear: number): Date | null {
  if (!dateStr) return null;
  
  const trimmed = dateStr.trim();
  if (trimmed === '미정' || trimmed === '' || trimmed === '-') return null;
  
  // "1/9", "01/09", "1월 9일" 등의 형식 처리
  const patterns = [
    /^(\d{1,2})\/(\d{1,2})$/,           // 1/9 또는 01/09
    /^(\d{1,2})월\s*(\d{1,2})일?$/,     // 1월 9일 또는 1월9
    /^(\d{4})-(\d{1,2})-(\d{1,2})$/,    // 2025-01-09
  ];
  
  for (const pattern of patterns) {
    const match = trimmed.match(pattern);
    if (match) {
      if (match.length === 4) {
        // YYYY-MM-DD 형식
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
      } else {
        // M/D 또는 M월 D일 형식
        const month = parseInt(match[1]) - 1;
        const day = parseInt(match[2]);
        return new Date(fallbackYear, month, day);
      }
    }
  }
  
  return null;
}

// 다음 근무일 계산 (주말 제외, n일 후)
function getNextWorkingDay(date: Date, daysToAdd: number = 1): Date {
  let result = new Date(date);
  let addedDays = 0;
  
  while (addedDays < daysToAdd) {
    result = addDays(result, 1);
    // 주말이 아니면 카운트
    if (!isWeekend(result)) {
      addedDays++;
    }
  }
  
  return result;
}

// 근무일만 필터링
function getWorkingDays(days: Date[]): Date[] {
  return days.filter(day => !isWeekend(day));
}

// 외주처별 라인별 일일 할당 현황
interface LineSchedule {
  vendorId: string;
  lineNumber: number;
  date: Date;
  allocatedQuantity: number;
  productionPlanId?: string;
}

// 외주처가 특수공정을 처리할 수 있는지 확인
function canHandleProcess(vendor: Vendor, processType: ProcessType): boolean {
  if (processType === 'normal') return true;
  return vendor.capabilities.includes(processType);
}

// 월간 남은 용량 계산
function getRemainingMonthlyCapacity(
  vendor: Vendor,
  schedules: LineSchedule[],
  targetMonth: Date
): number {
  const monthStart = startOfMonth(targetMonth);
  const monthEnd = endOfMonth(targetMonth);
  
  const allocated = schedules
    .filter(s => 
      s.vendorId === vendor.id &&
      s.date >= monthStart &&
      s.date <= monthEnd
    )
    .reduce((sum, s) => sum + s.allocatedQuantity, 0);
  
  return (vendor.monthlyTarget || Infinity) - allocated;
}

// 특정 날짜의 외주처 라인별 가용 용량 계산
function getAvailableCapacity(
  vendor: Vendor,
  date: Date,
  schedules: LineSchedule[]
): { lineNumber: number; available: number }[] {
  const result: { lineNumber: number; available: number }[] = [];
  
  for (let line = 1; line <= vendor.lineCount; line++) {
    const allocated = schedules
      .filter(s => 
        s.vendorId === vendor.id &&
        s.lineNumber === line &&
        s.date.getTime() === date.getTime()
      )
      .reduce((sum, s) => sum + s.allocatedQuantity, 0);
    
    const available = vendor.dailyCapacityPerLine - allocated;
    if (available > 0) {
      result.push({ lineNumber: line, available });
    }
  }
  
  return result;
}

// 외주처 선택 (우선순위 기반)
function selectVendor(
  item: ProductionItem,
  vendors: Vendor[],
  clientMappings: ClientVendorMapping[],
  schedules: LineSchedule[],
  targetMonth: Date
): Vendor | null {
  // 1. 고객사 매칭이 있으면 우선 확인
  if (item.clientCode) {
    const mapping = clientMappings.find(m => m.clientCode === item.clientCode);
    if (mapping) {
      const preferredVendor = vendors.find(v => v.id === mapping.vendorId);
      if (preferredVendor && canHandleProcess(preferredVendor, item.specialProcess || 'normal')) {
        return preferredVendor;
      }
    }
  }
  
  // 2. 특수공정 처리 가능하고, 월간 목표가 있는 외주처 우선
  const capableVendors = vendors
    .filter(v => canHandleProcess(v, item.specialProcess || 'normal'))
    .sort((a, b) => {
      // 우선순위 낮은 것(높은 우선순위) 먼저
      if (a.priority !== b.priority) return a.priority - b.priority;
      
      // 월간 목표가 있는 외주처 우선
      const aRemaining = getRemainingMonthlyCapacity(a, schedules, targetMonth);
      const bRemaining = getRemainingMonthlyCapacity(b, schedules, targetMonth);
      
      // 남은 용량 비율 기준으로 정렬 (더 채워야 하는 곳 우선)
      const aRatio = a.monthlyTarget ? aRemaining / a.monthlyTarget : 0;
      const bRatio = b.monthlyTarget ? bRemaining / b.monthlyTarget : 0;
      
      return bRatio - aRatio;
    });
  
  return capableVendors[0] || null;
}

// 생산 계획 생성
export function allocateProduction(
  items: ProductionItem[],
  vendors: Vendor[],
  clientMappings: ClientVendorMapping[],
  targetMonth: Date
): { allocatedItems: ProductionItem[]; plans: ProductionPlan[] } {
  const schedules: LineSchedule[] = [];
  const plans: ProductionPlan[] = [];
  const allocatedItems: ProductionItem[] = [];
  
  // 해당 월의 작업 가능 일자 (주말 제외)
  const monthStart = startOfMonth(targetMonth);
  const monthEnd = endOfMonth(targetMonth);
  const allDays = eachDayOfInterval({ start: monthStart, end: monthEnd });
  const workingDays = getWorkingDays(allDays);
  
  const targetYear = targetMonth.getFullYear();
  
  // 이미 외주처가 배정된 항목과 미배정 항목 분리
  const assignedItems = items.filter(item => item.assignedVendor);
  const unassignedItems = items.filter(item => !item.assignedVendor);
  
  // 먼저 이미 배정된 항목 처리
  for (const item of assignedItems) {
    allocatedItems.push({ ...item });
  }
  
  // 미배정 항목 처리
  for (const item of unassignedItems) {
    const vendor = selectVendor(item, vendors, clientMappings, schedules, targetMonth);
    
    if (!vendor) {
      // 배정 불가능한 경우 원본 유지
      allocatedItems.push({ ...item, assignedVendor: '배정불가' });
      continue;
    }
    
    // 외주처 배정
    const allocatedItem: ProductionItem = {
      ...item,
      assignedVendor: vendorNameMap[vendor.id] || vendor.name,
    };
    
    // 이동일 파싱 및 생산 시작일 계산 (이동일 + 1 근무일)
    const transferDate = parseTransferDate(item.transferDate, targetYear);
    let productionStartDate: Date;
    
    if (transferDate) {
      // 이동일이 있으면 이동일 + 1 근무일부터 시작
      productionStartDate = getNextWorkingDay(transferDate, 1);
    } else {
      // 이동일이 없으면 월 첫 근무일부터 시작
      productionStartDate = workingDays[0];
    }
    
    // 생산 시작일 이후의 근무일만 필터링
    const availableWorkingDays = workingDays.filter(day => day >= productionStartDate);
    
    if (availableWorkingDays.length === 0) {
      // 해당 월에 생산 가능한 날이 없으면 배정불가
      allocatedItems.push({ ...item, assignedVendor: '배정불가(일정초과)' });
      continue;
    }
    
    // 생산 일정 계획
    let remainingQuantity = item.quantity;
    let currentDayIndex = 0;
    let planStartDate = availableWorkingDays[0];
    let planEndDate = availableWorkingDays[0];
    let isFirstAllocation = true;
    
    while (remainingQuantity > 0 && currentDayIndex < availableWorkingDays.length) {
      const currentDate = availableWorkingDays[currentDayIndex];
      const availableLines = getAvailableCapacity(vendor, currentDate, schedules);
      
      if (availableLines.length === 0) {
        currentDayIndex++;
        continue;
      }
      
      // 가용 라인에 배분
      for (const line of availableLines) {
        if (remainingQuantity <= 0) break;
        
        const allocateQty = Math.min(remainingQuantity, line.available);
        
        schedules.push({
          vendorId: vendor.id,
          lineNumber: line.lineNumber,
          date: currentDate,
          allocatedQuantity: allocateQty,
        });
        
        if (isFirstAllocation) {
          planStartDate = currentDate;
          isFirstAllocation = false;
        }
        
        remainingQuantity -= allocateQty;
        planEndDate = currentDate;
      }
      
      currentDayIndex++;
    }
    
    // 생산 계획 생성
    const plan: ProductionPlan = {
      id: `plan-${item.id}-${Date.now()}`,
      productionItemId: item.id,
      vendorId: vendor.id,
      vendorName: vendorNameMap[vendor.id] || vendor.name,
      lineNumber: 1, // 다중 라인 사용시 대표값
      startDate: planStartDate,
      endDate: planEndDate,
      dailyQuantity: vendor.dailyCapacityPerLine,
      totalQuantity: item.quantity,
      productName: item.productName,
      productCode: item.productCode,
      status: 'planned',
    };
    
    plans.push(plan);
    allocatedItems.push(allocatedItem);
  }
  
  return { allocatedItems, plans };
}

// 외주처별 통계 계산
export function calculateVendorStats(
  plans: ProductionPlan[],
  vendors: Vendor[]
): { vendorId: string; vendorName: string; totalQuantity: number; monthlyTarget: number }[] {
  const stats: Record<string, number> = {};
  
  for (const plan of plans) {
    stats[plan.vendorId] = (stats[plan.vendorId] || 0) + plan.totalQuantity;
  }
  
  return vendors.map(vendor => ({
    vendorId: vendor.id,
    vendorName: vendorNameMap[vendor.id] || vendor.name,
    totalQuantity: stats[vendor.id] || 0,
    monthlyTarget: vendor.monthlyTarget || 0,
  }));
}

// 일별 총 생산량 계산
export function calculateDailyTotals(
  plans: ProductionPlan[],
  targetMonth: Date
): { date: string; quantity: number }[] {
  const monthStart = startOfMonth(targetMonth);
  const monthEnd = endOfMonth(targetMonth);
  const days = eachDayOfInterval({ start: monthStart, end: monthEnd });
  
  const dailyTotals: Record<string, number> = {};
  
  for (const plan of plans) {
    const planStart = new Date(plan.startDate);
    const planEnd = new Date(plan.endDate);
    const planDays = eachDayOfInterval({ start: planStart, end: planEnd });
    const dailyQty = Math.ceil(plan.totalQuantity / planDays.length);
    
    for (const day of planDays) {
      const dateStr = format(day, 'yyyy-MM-dd');
      dailyTotals[dateStr] = (dailyTotals[dateStr] || 0) + dailyQty;
    }
  }
  
  return days.map(day => ({
    date: format(day, 'MM/dd'),
    quantity: dailyTotals[format(day, 'yyyy-MM-dd')] || 0,
  }));
}

// ==================== 고급 알고리즘 ====================

// 1. LPT (Longest Processing Time) 규칙: 처리량 큰 항목 우선 배정
export function sortByLPT(items: ProductionItem[]): ProductionItem[] {
  return [...items].sort((a, b) => b.quantity - a.quantity);
}

// 2. 외주처별 부하 계산 (Min-Max 최적화용)
export interface VendorLoadMetrics {
  vendorId: string;
  vendorName: string;
  totalLoad: number;
  utilization: number; // 0-100%
  monthlyTargetAchievementRate: number; // 0-100%
}

export function calculateVendorLoads(
  plans: ProductionPlan[],
  vendors: Vendor[]
): VendorLoadMetrics[] {
  const vendorTotals: Record<string, number> = {};
  
  for (const plan of plans) {
    vendorTotals[plan.vendorId] = (vendorTotals[plan.vendorId] || 0) + plan.totalQuantity;
  }
  
  return vendors.map(vendor => {
    const totalLoad = vendorTotals[vendor.id] || 0;
    const maxDailyCapacity = vendor.lineCount * vendor.dailyCapacityPerLine * 22; // 월간 22 근무일 기준
    const utilization = (totalLoad / maxDailyCapacity) * 100;
    const monthlyTarget = vendor.monthlyTarget || maxDailyCapacity;
    const achievementRate = (totalLoad / monthlyTarget) * 100;
    
    return {
      vendorId: vendor.id,
      vendorName: vendorNameMap[vendor.id] || vendor.name,
      totalLoad,
      utilization: Math.min(utilization, 100),
      monthlyTargetAchievementRate: Math.min(achievementRate, 100),
    };
  });
}

// 3. 부하 균형도 계산 (0 = 완벽한 균형, 1 = 최악의 불균형)
export function calculateLoadBalance(metrics: VendorLoadMetrics[]): {
  balance: number; // 0-1
  maxLoad: number;
  minLoad: number;
  avgLoad: number;
  stdDeviation: number;
} {
  if (metrics.length === 0) {
    return { balance: 0, maxLoad: 0, minLoad: 0, avgLoad: 0, stdDeviation: 0 };
  }
  
  const loads = metrics.map(m => m.totalLoad);
  const maxLoad = Math.max(...loads);
  const minLoad = Math.min(...loads);
  const avgLoad = loads.reduce((a, b) => a + b, 0) / loads.length;
  
  // 표준편차 계산
  const variance = loads.reduce((sum, load) => sum + Math.pow(load - avgLoad, 2), 0) / loads.length;
  const stdDeviation = Math.sqrt(variance);
  
  // 부하 균형도: 표준편차 / 평균 (정규화된 변동 계수)
  const balance = avgLoad > 0 ? stdDeviation / avgLoad : 0;
  
  return {
    balance: Math.min(balance, 1),
    maxLoad,
    minLoad,
    avgLoad,
    stdDeviation,
  };
}

// 4. 납기일 기반 동적 우선순위 계산
export function calculateDeliveryUrgency(
  item: ProductionItem,
  referenceDate: Date = new Date()
): number {
  if (!item.deliveryDate) return 0;
  
  const deliveryDate = new Date(item.deliveryDate);
  const daysUntilDelivery = Math.ceil(
    (deliveryDate.getTime() - referenceDate.getTime()) / (1000 * 60 * 60 * 24)
  );
  
  // 기본 우선순위: 10점 만점
  // 60일 이상: 1점, 30일: 5점, 7일: 9점, 당일: 10점
  if (daysUntilDelivery >= 60) return 1;
  if (daysUntilDelivery <= 0) return 10;
  
  return Math.round(10 - (daysUntilDelivery / 60) * 9);
}

// 5. 특수공정 활용도 계산
export function calculateSpecialProcessUtilization(
  plans: ProductionPlan[],
  vendors: Vendor[]
): {
  vendorId: string;
  vendorName: string;
  specialProcessLoad: number;
  specialProcessCapability: boolean;
  utilizationRate: number;
}[] {
  // 특수공정 항목 필터링 (여기서는 예시)
  const specialProcessLoad: Record<string, number> = {};
  
  vendors.forEach(vendor => {
    specialProcessLoad[vendor.id] = 0;
  });
  
  // 실제 구현: ProductionItem에서 특수공정 판별
  // 이는 데이터 구조에 따라 달라짐
  
  return vendors.map(vendor => ({
    vendorId: vendor.id,
    vendorName: vendorNameMap[vendor.id] || vendor.name,
    specialProcessLoad: specialProcessLoad[vendor.id] || 0,
    specialProcessCapability: vendor.capabilities.length > 1,
    utilizationRate: (specialProcessLoad[vendor.id] || 0) / 1000, // 정규화
  }));
}

// 6. 최적화된 배치 함수 (LPT + Load Balancing)
export function allocateProductionOptimized(
  items: ProductionItem[],
  vendors: Vendor[],
  clientMappings: ClientVendorMapping[],
  targetMonth: Date,
  options: {
    useLPT?: boolean; // LPT 규칙 사용 (기본값: true)
    balanceLoad?: boolean; // 부하 균형화 (기본값: true)
    considerUrgency?: boolean; // 납기일 우선순위 (기본값: true)
  } = {}
): { allocatedItems: ProductionItem[]; plans: ProductionPlan[] } {
  const {
    useLPT = true,
    balanceLoad = true,
    considerUrgency = true,
  } = options;
  
  // 1단계: LPT 규칙으로 정렬 (선택사항)
  let sortedItems = items;
  if (useLPT) {
    sortedItems = sortByLPT(items);
  } else if (considerUrgency) {
    // 납기일 기반 우선순위 정렬
    sortedItems = [...items].sort((a, b) => {
      const urgencyA = calculateDeliveryUrgency(a);
      const urgencyB = calculateDeliveryUrgency(b);
      return urgencyB - urgencyA; // 내림차순
    });
  }
  
  // 2단계: 기본 배치
  let result = allocateProduction(
    sortedItems,
    vendors,
    clientMappings,
    targetMonth
  );
  
  // 3단계: 부하 균형화 (선택사항)
  if (balanceLoad) {
    result = optimizeLoadBalancing(
      result.plans,
      result.allocatedItems,
      vendors,
      clientMappings,
      targetMonth
    );
  }
  
  return result;
}

// 7. 부하 균형화 최적화 (Local Search)
function optimizeLoadBalancing(
  plans: ProductionPlan[],
  items: ProductionItem[],
  vendors: Vendor[],
  clientMappings: ClientVendorMapping[],
  targetMonth: Date
): { allocatedItems: ProductionItem[]; plans: ProductionPlan[] } {
  let currentPlans = [...plans];
  let improved = true;
  let iterations = 0;
  const maxIterations = 50; // 과도한 계산 방지
  
  while (improved && iterations < maxIterations) {
    improved = false;
    iterations++;
    
    const metrics = calculateVendorLoads(currentPlans, vendors);
    const balance = calculateLoadBalance(metrics);
    
    // 부하가 가장 큰 외주처와 작은 외주처 찾기
    const maxLoadMetric = metrics.reduce((prev, current) =>
      prev.totalLoad > current.totalLoad ? prev : current
    );
    const minLoadMetric = metrics.reduce((prev, current) =>
      prev.totalLoad < current.totalLoad ? prev : current
    );
    
    if (maxLoadMetric.totalLoad - minLoadMetric.totalLoad <= 1000) {
      // 부하 차이가 충분히 작으면 종료
      break;
    }
    
    // 최대 부하 외주처에서 최소 부하 외주처로 이동 가능한 계획 찾기
    for (let i = 0; i < currentPlans.length; i++) {
      if (currentPlans[i].vendorId !== maxLoadMetric.vendorId) continue;
      
      const plan = currentPlans[i];
      const moveQuantity = Math.min(
        plan.totalQuantity,
        Math.ceil((maxLoadMetric.totalLoad - minLoadMetric.totalLoad) / 2)
      );
      
      // 이동 가능한지 확인 (제약조건 검증)
      const minLoadVendor = vendors.find(v => v.id === minLoadMetric.vendorId);
      if (!minLoadVendor) continue;
      
      // 특수공정 능력 확인
      const item = items.find(item => item.id === plan.productionItemId);
      if (item && !canHandleProcess(minLoadVendor, item.specialProcess || 'normal')) {
        continue;
      }
      
      // 이동 실행
      currentPlans[i] = {
        ...plan,
        vendorId: minLoadMetric.vendorId,
        vendorName: minLoadMetric.vendorName,
      };
      
      improved = true;
      break;
    }
  }
  
  return { allocatedItems: items, plans: currentPlans };
}

// 8. 스케줄링 성능 지표 (Scheduling Metrics)
export interface SchedulingPerformanceMetrics {
  makespan: number; // 전체 완료 시간 (일수)
  averageUtilization: number; // 평균 자원 활용률 (0-100%)
  loadBalance: number; // 부하 균형도 (0-1, 낮을수록 좋음)
  deliveryOnTimeRate: number; // 납기 준수율 (0-100%)
  monthlyTargetAchievementRate: number; // 월간 목표 달성률 (0-100%)
  specialProcessComplianceRate: number; // 특수공정 준수율 (0-100%)
}

export function calculateSchedulingMetrics(
  plans: ProductionPlan[],
  items: ProductionItem[],
  vendors: Vendor[]
): SchedulingPerformanceMetrics {
  const metrics = calculateVendorLoads(plans, vendors);
  const balance = calculateLoadBalance(metrics);
  
  // Makespan 계산 (최종 완료일)
  const makespanDays = Math.max(
    ...plans.map(p => Math.ceil((new Date(p.endDate).getTime() - new Date(p.startDate).getTime()) / (1000 * 60 * 60 * 24))),
    0
  );
  
  // 평균 활용률
  const avgUtilization = metrics.length > 0
    ? metrics.reduce((sum, m) => sum + m.utilization, 0) / metrics.length
    : 0;
  
  // 월간 목표 달성률
  const avgTargetAchievementRate = metrics.length > 0
    ? metrics.reduce((sum, m) => sum + m.monthlyTargetAchievementRate, 0) / metrics.length
    : 0;
  
  // 납기 준수율 (실제 구현: deliveryDate 비교 필요)
  const deliveryOnTimeRate = 100; // 임시값
  
  // 특수공정 준수율
  const specialProcessComplianceRate = 100; // 임시값
  
  return {
    makespan: makespanDays,
    averageUtilization: Math.round(avgUtilization * 100) / 100,
    loadBalance: Math.round(balance.balance * 1000) / 1000,
    deliveryOnTimeRate,
    monthlyTargetAchievementRate: Math.round(avgTargetAchievementRate * 100) / 100,
    specialProcessComplianceRate,
  };
}
