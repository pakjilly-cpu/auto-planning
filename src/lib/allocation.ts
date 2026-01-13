import { addDays, startOfMonth, endOfMonth, eachDayOfInterval, isWeekend, format } from 'date-fns';
import { Vendor, ClientVendorMapping, ProductionItem, ProductionPlan, ProcessType } from './types';
import { vendorNameMap } from '@/data/defaults';

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
  
  // 해당 월의 작업 가능 일자 (주말 제외 옵션 가능)
  const monthStart = startOfMonth(targetMonth);
  const monthEnd = endOfMonth(targetMonth);
  const workingDays = eachDayOfInterval({ start: monthStart, end: monthEnd });
  
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
    
    // 생산 일정 계획
    let remainingQuantity = item.quantity;
    let currentDayIndex = 0;
    const planStartDate = workingDays[0];
    let planEndDate = workingDays[0];
    
    while (remainingQuantity > 0 && currentDayIndex < workingDays.length) {
      const currentDate = workingDays[currentDayIndex];
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
