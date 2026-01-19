'use client';

import { useMemo, useRef, useState, useCallback, useEffect } from 'react';
import { format, eachDayOfInterval, startOfMonth, endOfMonth, addMonths, subMonths, isToday, addDays, isWeekend } from 'date-fns';
import { ko } from 'date-fns/locale';
import { ChevronLeft, ChevronRight, X, ZoomIn, ZoomOut, GripVertical, Check } from 'lucide-react';
import { useAppStore, FixedPlacement } from '@/lib/store';
import { ProductionItem, Vendor } from '@/lib/types';

// 드래그 중인 아이템 정보
interface DragItem {
  item: ProductionItem;
  sourceVendor: string;
  sourceLine: number;
  sourceDate: string;
}

// 드롭 대상 정보
interface DropTarget {
  vendorName: string;
  lineNumber: number;
  dateKey: string;
}

// 보류 중인 변경사항
interface PendingChange {
  item: ProductionItem;
  from: { vendor: string; line: number; date: string };
  to: { vendor: string; line: number; date: string };
}

// 최신 트렌드 그라디언트 색상 (2024-2025)
const vendorColors: Record<string, { bg: string; text: string; border: string; light: string; gradient: string }> = {
  '위드맘': {
    bg: 'bg-gradient-to-br from-blue-600 to-blue-700',
    text: 'text-blue-600',
    border: 'border-blue-200',
    light: 'bg-blue-50',
    gradient: 'from-blue-50 via-blue-50 to-blue-100',
  },
  '리니어': {
    bg: 'bg-gradient-to-br from-emerald-600 to-emerald-700',
    text: 'text-emerald-600',
    border: 'border-emerald-200',
    light: 'bg-emerald-50',
    gradient: 'from-emerald-50 via-emerald-50 to-emerald-100',
  },
  '그램': {
    bg: 'bg-gradient-to-br from-violet-600 to-violet-700',
    text: 'text-violet-600',
    border: 'border-violet-200',
    light: 'bg-violet-50',
    gradient: 'from-violet-50 via-violet-50 to-violet-100',
  },
  '이시스': {
    bg: 'bg-gradient-to-br from-orange-600 to-orange-700',
    text: 'text-orange-600',
    border: 'border-orange-200',
    light: 'bg-orange-50',
    gradient: 'from-orange-50 via-orange-50 to-orange-100',
  },
  '엘루오': {
    bg: 'bg-gradient-to-br from-rose-600 to-rose-700',
    text: 'text-rose-600',
    border: 'border-rose-200',
    light: 'bg-rose-50',
    gradient: 'from-rose-50 via-rose-50 to-rose-100',
  },
  '케이코스텍': {
    bg: 'bg-gradient-to-br from-cyan-600 to-cyan-700',
    text: 'text-cyan-600',
    border: 'border-cyan-200',
    light: 'bg-cyan-50',
    gradient: 'from-cyan-50 via-cyan-50 to-cyan-100',
  },
  '다미': {
    bg: 'bg-gradient-to-br from-amber-600 to-amber-700',
    text: 'text-amber-600',
    border: 'border-amber-200',
    light: 'bg-amber-50',
    gradient: 'from-amber-50 via-amber-50 to-amber-100',
  },
  '배정불가': {
    bg: 'bg-gradient-to-br from-gray-500 to-gray-600',
    text: 'text-gray-600',
    border: 'border-gray-200',
    light: 'bg-gray-50',
    gradient: 'from-gray-50 via-gray-50 to-gray-100',
  },
};

// 이동일 문자열을 Date로 파싱
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
        return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
      } else {
        const month = parseInt(match[1]) - 1;
        const day = parseInt(match[2]);
        return new Date(fallbackYear, month, day);
      }
    }
  }
  
  return null;
}

// 다음 근무일 계산 (주말 제외)
function getNextWorkingDay(date: Date, daysToAdd: number = 1): Date {
  let result = new Date(date);
  let addedDays = 0;
  
  while (addedDays < daysToAdd) {
    result = addDays(result, 1);
    if (!isWeekend(result)) {
      addedDays++;
    }
  }
  
  return result;
}

// 생산 시작일 계산 (이동일 + 1 근무일)
function getProductionStartDate(transferDateStr: string | undefined, targetYear: number, fallbackDate: Date): Date {
  const transferDate = parseTransferDate(transferDateStr, targetYear);
  if (transferDate) {
    return getNextWorkingDay(transferDate, 1);
  }
  return fallbackDate;
}

// 줌 레벨 설정
const ZOOM_LEVELS = [50, 75, 100, 125, 150];

export default function GanttChart() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const { 
    productionItems, 
    selectedMonth: rawSelectedMonth, 
    setSelectedMonth, 
    vendors,
    fixedPlacements,
    setFixedPlacement,
  } = useAppStore();
  
  // selectedMonth가 문자열일 수 있으므로 Date로 변환
  const selectedMonth = rawSelectedMonth instanceof Date ? rawSelectedMonth : new Date(rawSelectedMonth);
  const [selectedItem, setSelectedItem] = useState<ProductionItem | null>(null);
  const [zoomLevel, setZoomLevel] = useState(100);
  
  // 드래그 앤 드롭 상태
  const [dragItem, setDragItem] = useState<DragItem | null>(null);
  const [dropTarget, setDropTarget] = useState<DropTarget | null>(null);
  const [pendingChange, setPendingChange] = useState<PendingChange | null>(null);
  
  // 일별 생산 품목 (CAPA 반영)
  interface DailyItem {
    item: ProductionItem;
    dailyQty: number;  // 해당 날짜에 생산할 수량
    dayNumber: number; // 몇 번째 날인지 (1, 2, 3...)
    totalDays: number; // 총 며칠 걸리는지
  }
  
  // 스케줄 상태 (useState로 관리 - 드래그 시 재계산 안함)
  const [scheduleData, setScheduleData] = useState<Record<string, Record<number, Record<string, DailyItem[]>>>>({});
  
  // 스크롤로 월 이동 (기능 비활성화 - 버튼으로만 이동)
  // 스크롤 끝에서 자동 월 전환은 UX 문제가 있어 제거
  // 스크롤로 월 이동 (비활성화 - 복잡한 상태 관리로 인해 버튼으로만 이동)
  const handleScroll = useCallback(() => {
    // 스크롤 월 이동은 비활성화
    // 월 이동은 좌/우 화살표 버튼 사용
  }, []);

  // 줌 조절
  const handleZoomIn = () => {
    const currentIndex = ZOOM_LEVELS.indexOf(zoomLevel);
    if (currentIndex < ZOOM_LEVELS.length - 1) {
      setZoomLevel(ZOOM_LEVELS[currentIndex + 1]);
    }
  };

  const handleZoomOut = () => {
    const currentIndex = ZOOM_LEVELS.indexOf(zoomLevel);
    if (currentIndex > 0) {
      setZoomLevel(ZOOM_LEVELS[currentIndex - 1]);
    }
  };
  
  // 드래그 시작
  const handleDragStart = (item: ProductionItem, vendorName: string, lineNumber: number, dateKey: string) => {
    setDragItem({
      item,
      sourceVendor: vendorName,
      sourceLine: lineNumber,
      sourceDate: dateKey,
    });
  };
  
  // 드래그 오버 (드롭 대상 설정)
  const handleDragOver = (e: React.DragEvent, vendorName: string, lineNumber: number, dateKey: string) => {
    e.preventDefault();
    setDropTarget({ vendorName, lineNumber, dateKey });
  };
  
  // 드래그 종료
  const handleDragEnd = () => {
    if (dragItem && dropTarget) {
      // 위치가 변경되었는지 확인
      if (
        dragItem.sourceVendor !== dropTarget.vendorName ||
        dragItem.sourceLine !== dropTarget.lineNumber ||
        dragItem.sourceDate !== dropTarget.dateKey
      ) {
        setPendingChange({
          item: dragItem.item,
          from: {
            vendor: dragItem.sourceVendor,
            line: dragItem.sourceLine,
            date: dragItem.sourceDate,
          },
          to: {
            vendor: dropTarget.vendorName,
            line: dropTarget.lineNumber,
            date: dropTarget.dateKey,
          },
        });
      }
    }
    setDragItem(null);
    setDropTarget(null);
  };
  
  // 변경 확정 - 스케줄을 직접 수정 (다른 품목은 그대로 유지)
  const confirmChange = () => {
    if (!pendingChange) return;
    
    const { item, from, to } = pendingChange;
    
    // 스케줄 복사
    const newSchedule = JSON.parse(JSON.stringify(scheduleData)) as typeof scheduleData;
    
    // 1. 원래 위치에서 해당 품목 모두 제거 (여러 날에 걸친 경우)
    if (newSchedule[from.vendor]?.[from.line]) {
      Object.keys(newSchedule[from.vendor][from.line]).forEach(dateKey => {
        newSchedule[from.vendor][from.line][dateKey] = 
          newSchedule[from.vendor][from.line][dateKey].filter(
            (d: DailyItem) => d.item.productCode !== item.productCode
          );
      });
    }
    
    // 2. 새 위치에 품목 추가
    const vendor = vendors.find((v: Vendor) => v.name === to.vendor);
    const dailyCapa = vendor?.dailyCapacityPerLine || 10000;
    const totalDays = Math.ceil(item.quantity / dailyCapa);
    
    // 주말 제외 근무일만 필터링
    const workingDays = days.filter(day => !isWeekend(day));
    const startDayIndex = workingDays.findIndex(d => format(d, 'yyyy-MM-dd') === to.date);
    
    if (startDayIndex !== -1 && newSchedule[to.vendor]?.[to.line]) {
      for (let dayOffset = 0; dayOffset < totalDays; dayOffset++) {
        const dayIndex = startDayIndex + dayOffset;
        if (dayIndex >= workingDays.length) break;
        
        const currentDay = workingDays[dayIndex];
        const dateKey = format(currentDay, 'yyyy-MM-dd');
        
        const remainingQty = item.quantity - (dayOffset * dailyCapa);
        const dailyQty = Math.min(remainingQty, dailyCapa);
        
        if (newSchedule[to.vendor][to.line][dateKey]) {
          newSchedule[to.vendor][to.line][dateKey].push({
            item,
            dailyQty,
            dayNumber: dayOffset + 1,
            totalDays,
          });
        }
      }
    }
    
    // 3. 스케줄 상태 업데이트
    setScheduleData(newSchedule);
    
    // 4. 고정 배치 저장 (다음 엑셀 업로드 시 유지)
    const placement: FixedPlacement = {
      productCode: item.productCode,
      vendorName: to.vendor,
      lineNumber: to.line,
      dateKey: to.date,
    };
    setFixedPlacement(placement);
    
    setPendingChange(null);
  };
  
  // 변경 취소
  const cancelChange = () => {
    setPendingChange(null);
  };

  // 줌에 따른 셀 너비 계산
  const cellWidth = Math.round(96 * (zoomLevel / 100));
  const fixedColumnWidth = Math.round(192 * (zoomLevel / 100));
  const fontSize = zoomLevel < 75 ? 'text-[10px]' : zoomLevel < 100 ? 'text-xs' : 'text-sm';

  // 품목 더블클릭 핸들러
  const handleItemDoubleClick = (item: ProductionItem) => {
    setSelectedItem(item);
  };

  // 모달 닫기
  const closeModal = () => {
    setSelectedItem(null);
  };

  // 날짜 배열 생성
  const days = useMemo(() => {
    const start = startOfMonth(selectedMonth);
    const end = endOfMonth(selectedMonth);
    return eachDayOfInterval({ start, end });
  }, [selectedMonth]);

  // 외주처별 그룹핑
  const groupedByVendor = useMemo(() => {
    const groups: Record<string, ProductionItem[]> = {};
    
    // 외주처 순서 정의 (우선순위 순)
    const vendorOrder = ['위드맘', '리니어', '그램', '이시스', '엘루오', '케이코스텍', '다미', '배정불가'];
    
    vendorOrder.forEach(vendor => {
      groups[vendor] = [];
    });
    
    productionItems.forEach((item: ProductionItem) => {
      const vendor = item.assignedVendor || '배정불가';
      if (!groups[vendor]) groups[vendor] = [];
      groups[vendor].push(item);
    });
    
    return groups;
  }, [productionItems]);

  // 초기 스케줄 계산 함수
  const calculateInitialSchedule = useCallback(() => {
    const schedules: Record<string, Record<number, Record<string, DailyItem[]>>> = {};
    
    // 주말 제외 근무일만 필터링
    const workingDays = days.filter(day => !isWeekend(day));
    
    // 모든 외주처에 대해 스케줄 초기화 (하드코딩 + 실제 데이터의 외주처 병합)
    const vendorOrder = ['위드맘', '리니어', '그램', '이시스', '엘루오', '케이코스텍', '다미', '배정불가'];
    const allVendors = new Set([...vendorOrder, ...Object.keys(groupedByVendor)]);
    allVendors.forEach(vendorName => {
      const vendor = vendors.find((v: Vendor) => v.name === vendorName);
      const lineCount = vendor?.lineCount || 1;
      
      schedules[vendorName] = {};
      for (let line = 1; line <= lineCount; line++) {
        schedules[vendorName][line] = {};
        days.forEach(day => {
          schedules[vendorName][line][format(day, 'yyyy-MM-dd')] = [];
        });
      }
    });
    
    const currentMonth = new Date(selectedMonth);
    const targetYear = currentMonth.getFullYear();
    const monthEnd = endOfMonth(currentMonth);
    
    // 고정 배치된 품목 ID 추적
    const fixedProductCodes = new Set(fixedPlacements.map((p: FixedPlacement) => p.productCode));
    
    // 1단계: 고정 배치된 품목 먼저 배치
    fixedPlacements.forEach((placement: FixedPlacement) => {
      // 해당 품목 찾기
      const item = productionItems.find((i: ProductionItem) => i.productCode === placement.productCode);
      if (!item) return;
      
      const vendor = vendors.find((v: Vendor) => v.name === placement.vendorName);
      const dailyCapa = vendor?.dailyCapacityPerLine || 10000;
      
      // 고정 배치 날짜가 현재 월에 있는지 확인
      const placementDate = new Date(placement.dateKey);
      if (placementDate > monthEnd || format(placementDate, 'yyyy-MM') !== format(currentMonth, 'yyyy-MM')) return;
      
      // 스케줄에 추가
      if (schedules[placement.vendorName]?.[placement.lineNumber]?.[placement.dateKey]) {
        const totalDays = Math.ceil(item.quantity / dailyCapa);
        
        // 고정 배치 날짜부터 시작해서 여러 날에 걸쳐 배분
        const startDayIndex = workingDays.findIndex(d => format(d, 'yyyy-MM-dd') === placement.dateKey);
        if (startDayIndex === -1) return;
        
        for (let dayOffset = 0; dayOffset < totalDays; dayOffset++) {
          const dayIndex = startDayIndex + dayOffset;
          if (dayIndex >= workingDays.length) break;
          
          const currentDay = workingDays[dayIndex];
          const dateKey = format(currentDay, 'yyyy-MM-dd');
          
          const remainingQty = item.quantity - (dayOffset * dailyCapa);
          const dailyQty = Math.min(remainingQty, dailyCapa);
          
          if (schedules[placement.vendorName][placement.lineNumber][dateKey]) {
            schedules[placement.vendorName][placement.lineNumber][dateKey].push({
              item,
              dailyQty,
              dayNumber: dayOffset + 1,
              totalDays,
            });
          }
        }
      }
    });
    
    // 2단계: 고정되지 않은 품목 자동 배치
    // 한 라인의 하루에 한 품목만 배치 (CAPA가 남더라도 라인 교체 시간 필요)
    Object.entries(groupedByVendor).forEach(([vendorName, items]) => {
      const vendor = vendors.find((v: Vendor) => v.name === vendorName);
      const lineCount = vendor?.lineCount || 1;
      const dailyCapa = vendor?.dailyCapacityPerLine || 10000;
      
      // 라인별로 점유된 날짜 추적 (fixedPlacements로 이미 배치된 날짜 포함)
      const lineOccupiedDays: Record<number, Set<string>> = {};
      for (let line = 1; line <= lineCount; line++) {
        lineOccupiedDays[line] = new Set();
        // 이미 배치된 날짜 추가
        Object.keys(schedules[vendorName]?.[line] || {}).forEach(dateKey => {
          if (schedules[vendorName][line][dateKey].length > 0) {
            lineOccupiedDays[line].add(dateKey);
          }
        });
      }
      
      // 라인별로 품목 분배
      let currentLine = 1;
      items.forEach(item => {
        // 고정 배치된 품목은 스킵
        if (fixedProductCodes.has(item.productCode)) return;
        
        // 이동일 기반 생산 시작일 계산
        const productionStartDate = getProductionStartDate(item.transferDate, targetYear, workingDays[0] || days[0]);
        
        // 해당 월 범위 체크
        if (productionStartDate > monthEnd) return;
        
        // 생산 시작일 이후의 근무일 찾기
        let startDayIndex = workingDays.findIndex(d => d >= productionStartDate);
        if (startDayIndex === -1) startDayIndex = 0;
        
        // 필요한 일수 계산
        const totalDays = Math.ceil(item.quantity / dailyCapa);
        
        // 현재 라인에서 연속된 빈 날짜 찾기
        let foundStartIndex = -1;
        for (let searchStart = startDayIndex; searchStart <= workingDays.length - totalDays; searchStart++) {
          let canPlace = true;
          for (let offset = 0; offset < totalDays; offset++) {
            const checkDate = format(workingDays[searchStart + offset], 'yyyy-MM-dd');
            if (lineOccupiedDays[currentLine].has(checkDate)) {
              canPlace = false;
              break;
            }
          }
          if (canPlace) {
            foundStartIndex = searchStart;
            break;
          }
        }
        
        // 배치 가능한 날짜를 찾지 못하면 다음 라인 시도
        if (foundStartIndex === -1) {
          // 다른 라인에서 찾기
          for (let tryLine = 1; tryLine <= lineCount; tryLine++) {
            if (tryLine === currentLine) continue;
            for (let searchStart = startDayIndex; searchStart <= workingDays.length - totalDays; searchStart++) {
              let canPlace = true;
              for (let offset = 0; offset < totalDays; offset++) {
                const checkDate = format(workingDays[searchStart + offset], 'yyyy-MM-dd');
                if (lineOccupiedDays[tryLine].has(checkDate)) {
                  canPlace = false;
                  break;
                }
              }
              if (canPlace) {
                foundStartIndex = searchStart;
                currentLine = tryLine;
                break;
              }
            }
            if (foundStartIndex !== -1) break;
          }
        }
        
        // 여전히 찾지 못하면 이 품목은 배치 불가 (월 범위 초과)
        if (foundStartIndex === -1) return;
        
        // 여러 날에 걸쳐 배분
        for (let dayOffset = 0; dayOffset < totalDays; dayOffset++) {
          const dayIndex = foundStartIndex + dayOffset;
          if (dayIndex >= workingDays.length) break;
          
          const currentDay = workingDays[dayIndex];
          const dateKey = format(currentDay, 'yyyy-MM-dd');
          
          // 해당 날짜에 생산할 수량 계산
          const remainingQty = item.quantity - (dayOffset * dailyCapa);
          const dailyQty = Math.min(remainingQty, dailyCapa);
          
          if (schedules[vendorName]?.[currentLine]?.[dateKey]) {
            schedules[vendorName][currentLine][dateKey].push({
              item,
              dailyQty,
              dayNumber: dayOffset + 1,
              totalDays,
            });
            // 점유된 날짜로 마킹
            lineOccupiedDays[currentLine]?.add(dateKey);
          }
        }
        
        // 다음 라인으로 순환
        currentLine = (currentLine % lineCount) + 1;
      });
    });
    
    return schedules;
  }, [days, vendors, selectedMonth, fixedPlacements, productionItems, groupedByVendor]);

  // productionItems 또는 selectedMonth 변경 시에만 초기 스케줄 계산
  useEffect(() => {
    if (productionItems.length > 0) {
      const initialSchedule = calculateInitialSchedule();
      setScheduleData(initialSchedule);
    }
  }, [productionItems, selectedMonth, vendors]); // fixedPlacements는 제외! (드래그 시 재계산 안함)

  // 현재 표시할 스케줄 (scheduleData 사용)
  const vendorLineSchedules = scheduleData;

  const handlePrevMonth = () => setSelectedMonth(subMonths(selectedMonth, 1));
  const handleNextMonth = () => setSelectedMonth(addMonths(selectedMonth, 1));

  if (productionItems.length === 0) {
    return (
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-8 text-center">
        <p className="text-gray-500">엑셀 파일을 업로드하면 생산계획표가 표시됩니다.</p>
      </div>
    );
  }

  return (
    <div className="bg-gradient-to-br from-white via-blue-50/20 to-white rounded-2xl shadow-lg border border-gray-100 overflow-hidden">
      {/* 헤더 - 모던 디자인 */}
      <div className="bg-gradient-to-r from-gray-900 via-gray-800 to-gray-900 px-6 py-6 border-b border-gray-200">
        <div className="flex items-center justify-between">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 bg-gradient-to-br from-blue-400 to-blue-600 rounded-xl shadow-lg flex items-center justify-center">
              <span className="text-white font-bold text-lg">📊</span>
            </div>
            <div>
              <h2 className="text-2xl font-bold text-white">생산 계획표</h2>
              <p className="text-gray-300 text-sm mt-1">실시간 라인별 배정 현황 대시보드</p>
            </div>
          </div>

          <div className="flex items-center gap-6">
            {/* 줌 컨트롤 */}
            <div className="flex items-center gap-2 bg-gray-700/50 backdrop-blur rounded-xl p-2 border border-gray-600">
              <button
                onClick={handleZoomOut}
                disabled={zoomLevel === ZOOM_LEVELS[0]}
                className="p-2 hover:bg-gray-600 rounded-lg disabled:opacity-30 disabled:cursor-not-allowed transition-all text-gray-300 hover:text-white"
                title="축소"
              >
                <ZoomOut className="w-5 h-5" />
              </button>
              <span className="text-gray-200 font-semibold min-w-[50px] text-center text-sm">{zoomLevel}%</span>
              <button
                onClick={handleZoomIn}
                disabled={zoomLevel === ZOOM_LEVELS[ZOOM_LEVELS.length - 1]}
                className="p-2 hover:bg-gray-600 rounded-lg disabled:opacity-30 disabled:cursor-not-allowed transition-all text-gray-300 hover:text-white"
                title="확대"
              >
                <ZoomIn className="w-5 h-5" />
              </button>
            </div>
            
            {/* 월 이동 */}
            <div className="flex items-center gap-3 bg-gray-700/50 backdrop-blur rounded-xl px-4 py-2 border border-gray-600">
              <button
                onClick={handlePrevMonth}
                className="p-2 hover:bg-gray-600 rounded-lg transition-all text-gray-300 hover:text-white"
              >
                <ChevronLeft className="w-5 h-5" />
              </button>
              <span className="font-bold text-white min-w-[140px] text-center text-sm">
                {format(selectedMonth, 'yyyy년 M월', { locale: ko })}
              </span>
              <button
                onClick={handleNextMonth}
                className="p-2 hover:bg-gray-600 rounded-lg transition-all text-gray-300 hover:text-white"
              >
                <ChevronRight className="w-5 h-5" />
              </button>
            </div>
          </div>
        </div>
      </div>
      
      {/* 간트 차트 - 최신 디자인 */}
      <div className="overflow-x-auto max-h-[70vh]" ref={scrollRef} onScroll={handleScroll}>
        <div className="min-w-max" style={{ fontSize: `${zoomLevel}%` }}>
          {/* 날짜 헤더 */}
          <div className="flex border-b-2 border-gray-200 sticky top-0 bg-gradient-to-r from-gray-50 to-gray-50 z-20 shadow-sm">
            <div 
              className="flex-shrink-0 p-4 font-bold text-gray-800 border-r-2 border-gray-200 bg-gradient-to-br from-gray-100 to-gray-50 sticky left-0 z-30 text-sm"
              style={{ width: `${fixedColumnWidth}px` }}
            >
              <div className="flex items-center gap-2">
                <span className="w-1 h-4 rounded-full bg-gradient-to-b from-blue-500 to-blue-600"></span>
                외주처 / 라인
              </div>
            </div>
            {days.map((day, index) => {
              const isWeekendDay = day.getDay() === 0 || day.getDay() === 6;
              const isTodayDate = isToday(day);
              return (
                <div
                  key={index}
                  className={`
                    flex-shrink-0 p-3 text-center border-r border-gray-200 ${fontSize} font-semibold
                    transition-all duration-200
                    ${isTodayDate 
                      ? 'bg-gradient-to-b from-blue-100 to-blue-50 border-blue-300 shadow-md' 
                      : isWeekendDay 
                      ? 'bg-gradient-to-b from-red-50 to-red-50/30' 
                      : 'bg-white hover:bg-gray-50'
                    }
                  `}
                  style={{ width: `${cellWidth}px` }}
                >
                  <div className={`font-bold text-sm ${isTodayDate ? 'text-blue-700' : isWeekendDay ? 'text-red-600' : 'text-gray-800'}`}>
                    {format(day, 'd')}
                  </div>
                  <div className={`text-xs font-medium mt-1 ${isTodayDate ? 'text-blue-600' : isWeekendDay ? 'text-red-500' : 'text-gray-500'}`}>
                    {format(day, 'EEE', { locale: ko })}
                  </div>
                </div>
              );
            })}
          </div>
          
          {/* 외주처별 행 */}
          {Object.entries(groupedByVendor).map(([vendorName, items]) => {
            if (items.length === 0) return null;
            
            const vendor = vendors.find((v: Vendor) => v.name === vendorName);
            const lineCount = vendor?.lineCount || 1;
            const totalQty = items.reduce((sum: number, item: ProductionItem) => sum + item.quantity, 0);
            const colorScheme = vendorColors[vendorName];
            
            return (
              <div key={vendorName} className="border-b-2 border-gray-200">
                {/* 외주처 헤더 */}
                <div className={`flex bg-gradient-to-r ${colorScheme.gradient} border-b border-gray-100`}>
                  <div 
                    className={`flex-shrink-0 p-4 border-r-2 border-gray-200 ${colorScheme.light} sticky left-0 z-10`}
                    style={{ width: `${fixedColumnWidth}px` }}
                  >
                    <div className="flex items-center gap-3">
                      <div className={`w-3 h-3 rounded-full ${colorScheme.bg} shadow-lg`}></div>
                      <div>
                        <p className={`font-bold text-sm ${colorScheme.text}`}>{vendorName}</p>
                        <p className="text-xs text-gray-600 mt-0.5 font-medium">
                          {items.length}건 · {totalQty.toLocaleString()}개
                        </p>
                      </div>
                    </div>
                  </div>
                  {days.map((_, idx) => (
                    <div 
                      key={idx} 
                      className={`flex-shrink-0 border-r border-gray-200 ${colorScheme.light} hover:opacity-80 transition-opacity`} 
                      style={{ width: `${cellWidth}px` }}
                    />
                  ))}
                </div>
                
                {/* 라인별 행 */}
                {Array.from({ length: lineCount }, (_, lineIdx) => {
                  const lineNumber = lineIdx + 1;
                  
                  return (
                    <div key={lineIdx} className="flex border-b border-gray-100 hover:bg-blue-50/30 transition-colors">
                      <div 
                        className={`flex-shrink-0 p-3 pl-6 border-r-2 border-gray-200 ${fontSize} font-semibold text-gray-700 bg-white sticky left-0 z-10 flex items-center gap-2`}
                        style={{ width: `${fixedColumnWidth}px` }}
                      >
                        <span className="w-1 h-3 rounded-full bg-gray-300"></span>
                        라인 {lineNumber}
                      </div>
                      {days.map((day, dayIdx) => {
                        const dateKey = format(day, 'yyyy-MM-dd');
                        const dayItems = vendorLineSchedules[vendorName]?.[lineNumber]?.[dateKey] || [];
                        const minHeight = zoomLevel < 75 ? 50 : zoomLevel < 100 ? 65 : 80;
                        const isDropTarget = dropTarget?.vendorName === vendorName && 
                                             dropTarget?.lineNumber === lineNumber && 
                                             dropTarget?.dateKey === dateKey;
                        const isWeekendDay = isWeekend(day);
                        
                        return (
                          <div
                            key={dayIdx}
                            className={`
                              flex-shrink-0 p-2 border-r border-gray-200 transition-all
                              ${isDropTarget ? `${colorScheme.light} ring-2 ${colorScheme.border} ring-inset shadow-inset` : ''}
                              ${isWeekendDay ? 'bg-red-50/40' : 'bg-white'}
                              ${dragItem && !isWeekendDay ? 'hover:bg-blue-100/50' : ''}
                            `}
                            style={{ width: `${cellWidth}px`, minHeight: `${minHeight}px` }}
                            onDragOver={(e) => {
                              if (!isWeekendDay) {
                                handleDragOver(e, vendorName, lineNumber, dateKey);
                              }
                            }}
                            onDragLeave={() => setDropTarget(null)}
                            onDrop={(e) => {
                              e.preventDefault();
                              if (!isWeekendDay) {
                                handleDragEnd();
                              }
                            }}
                          >
                              {dayItems.map((dailyItem, itemIdx) => {
                              const nameLength = zoomLevel < 75 ? 4 : zoomLevel < 100 ? 6 : 8;
                              const isFixed = fixedPlacements.some((p: FixedPlacement) => p.productCode === dailyItem.item.productCode);
                              return (
                                <div
                                  key={itemIdx}
                                  className="relative group"
                                >
                                  <div
                                    draggable
                                    onDragStart={(e) => {
                                      e.dataTransfer.effectAllowed = 'move';
                                      handleDragStart(dailyItem.item, vendorName, lineNumber, dateKey);
                                    }}
                                    onDragEnd={handleDragEnd}
                                    className={`
                                      p-2 rounded-lg mb-1.5 cursor-grab active:cursor-grabbing
                                      ${colorScheme.bg} text-white
                                      hover:shadow-lg hover:-translate-y-0.5 transition-all duration-200
                                      border-l-4 border-white/30
                                      select-none font-semibold backdrop-blur-sm
                                      ${zoomLevel < 75 ? 'text-[9px]' : zoomLevel < 100 ? 'text-[11px]' : 'text-xs'}
                                      ${dragItem?.item.productCode === dailyItem.item.productCode ? 'opacity-60 ring-2 ring-yellow-300 ring-offset-1' : ''}
                                      ${isFixed ? 'ring-2 ring-yellow-300 ring-offset-1' : ''}
                                    `}
                                    onDoubleClick={() => handleItemDoubleClick(dailyItem.item)}
                                  >
                                    <div className="flex items-center gap-1.5 mb-1">
                                      <GripVertical className="w-3 h-3 opacity-70 flex-shrink-0" />
                                      <span className="font-bold truncate leading-tight">
                                        {dailyItem.item.productName.slice(0, nameLength)}
                                      </span>
                                      {isFixed && <span className="text-lg">📌</span>}
                                    </div>
                                    <div className="text-white/90 font-bold text-xs">
                                      {dailyItem.dailyQty.toLocaleString()}개
                                    </div>
                                    {dailyItem.totalDays > 1 && (
                                      <div className="text-white/60 text-[8px] mt-0.5 font-medium">
                                        Day {dailyItem.dayNumber}/{dailyItem.totalDays}
                                      </div>
                                    )}
                                  </div>
                                  {/* 호버 툴팁 */}
                                  <div className="absolute left-full top-0 ml-2 z-50 hidden group-hover:block pointer-events-none">
                                    <div className="bg-gray-900 text-white text-xs rounded-lg shadow-lg p-2.5 whitespace-nowrap min-w-[180px]">
                                      <div className="font-semibold text-sm mb-1.5 text-blue-300">{dailyItem.item.productName}</div>
                                      <div className="space-y-1 text-gray-200">
                                        <div className="flex justify-between gap-4">
                                          <span className="text-gray-400">제품코드</span>
                                          <span className="font-mono">{dailyItem.item.productCode}</span>
                                        </div>
                                        <div className="flex justify-between gap-4">
                                          <span className="text-gray-400">총 수량</span>
                                          <span>{dailyItem.item.quantity.toLocaleString()}개</span>
                                        </div>
                                        <div className="flex justify-between gap-4">
                                          <span className="text-gray-400">오늘 생산</span>
                                          <span className="text-green-400">{dailyItem.dailyQty.toLocaleString()}개</span>
                                        </div>
                                        <div className="flex justify-between gap-4">
                                          <span className="text-gray-400">납기</span>
                                          <span>{dailyItem.item.deliveryDate}</span>
                                        </div>
                                        <div className="flex justify-between gap-4">
                                          <span className="text-gray-400">이동일</span>
                                          <span>{dailyItem.item.transferDate}일</span>
                                        </div>
                                        {dailyItem.totalDays > 1 && (
                                          <div className="flex justify-between gap-4 pt-1 border-t border-gray-700">
                                            <span className="text-gray-400">생산 진행</span>
                                            <span className="text-yellow-400">{dailyItem.dayNumber}일차 / {dailyItem.totalDays}일</span>
                                          </div>
                                        )}
                                      </div>
                                      {isFixed && (
                                        <div className="mt-2 pt-1.5 border-t border-gray-700 text-yellow-400 text-[10px]">
                                          📌 고정 배치된 품목
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        );
                      })}
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
      </div>

      {/* 변경 확정 다이얼로그 */}
      {pendingChange && (
        <div 
          className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50"
          onClick={cancelChange}
        >
          <div 
            className="bg-white rounded-2xl shadow-2xl p-8 max-w-lg w-full mx-4 border border-gray-100 transform transition-all"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-4 mb-6">
              <div className="w-12 h-12 bg-gradient-to-br from-blue-100 to-blue-50 rounded-full flex items-center justify-center border-2 border-blue-200">
                <Check className="w-6 h-6 text-blue-600" />
              </div>
              <h3 className="text-xl font-bold text-gray-900">배치 계획 변경</h3>
            </div>
            
            <div className="bg-gradient-to-br from-gray-50 to-blue-50 rounded-xl p-5 space-y-3 border border-gray-200 mb-6">
              <div className="flex items-center gap-3">
                <span className="text-xs font-bold text-gray-500 w-14">📦 품목</span>
                <span className="font-bold text-gray-900">{pendingChange.item.productName}</span>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs font-bold text-gray-500 w-14">📍 코드</span>
                <span className="font-mono text-sm font-semibold text-gray-700">{pendingChange.item.productCode}</span>
              </div>
              <div className="flex items-center gap-3">
                <span className="text-xs font-bold text-gray-500 w-14">📊 수량</span>
                <span className="font-bold text-lg text-gray-900">{pendingChange.item.quantity.toLocaleString()}개</span>
              </div>
            </div>
              
            <div className="grid grid-cols-2 gap-4 mb-6">
              <div className="bg-gradient-to-br from-red-50 to-red-50/50 rounded-xl p-4 border-2 border-red-200">
                <div className="text-xs font-bold text-red-600 mb-2">❌ 변경 전</div>
                <div className="space-y-1">
                  <p className="font-bold text-gray-900 text-sm">{pendingChange.from.vendor}</p>
                  <p className="text-xs text-gray-600 font-medium">라인 {pendingChange.from.line}</p>
                  <p className="text-xs text-gray-600 font-mono">{pendingChange.from.date}</p>
                </div>
              </div>
              <div className="bg-gradient-to-br from-green-50 to-green-50/50 rounded-xl p-4 border-2 border-green-200">
                <div className="text-xs font-bold text-green-600 mb-2">✅ 변경 후</div>
                <div className="space-y-1">
                  <p className="font-bold text-gray-900 text-sm">{pendingChange.to.vendor}</p>
                  <p className="text-xs text-gray-600 font-medium">라인 {pendingChange.to.line}</p>
                  <p className="text-xs text-gray-600 font-mono">{pendingChange.to.date}</p>
                </div>
              </div>
            </div>
              
            <p className="text-xs text-gray-500 bg-blue-50 rounded-lg p-3 mb-6 border border-blue-200">
              💡 <span className="font-medium">참고:</span> 이 품목은 다음 엑셀 업로드 시에도 지정한 위치에 고정 배치됩니다.
            </p>

            <div className="flex gap-3">
              <button
                onClick={cancelChange}
                className="flex-1 py-3 border-2 border-gray-300 text-gray-700 rounded-xl hover:bg-gray-50 transition-all font-bold"
              >
                취소
              </button>
              <button
                onClick={confirmChange}
                className="flex-1 py-3 bg-gradient-to-r from-blue-600 to-blue-700 text-white rounded-xl hover:shadow-lg hover:-translate-y-0.5 transition-all font-bold"
              >
                변경 확정
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 품목 상세 모달 */}
      {selectedItem && (
        <div 
          className="fixed inset-0 bg-black/60 backdrop-blur-sm flex items-center justify-center z-50"
          onClick={closeModal}
        >
          <div 
            className="bg-white rounded-2xl shadow-2xl p-8 max-w-md w-full mx-4 border border-gray-100 transform transition-all"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-6">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-gradient-to-br from-blue-100 to-blue-50 rounded-full flex items-center justify-center border border-blue-200">
                  <span className="text-lg">📦</span>
                </div>
                <h3 className="text-xl font-bold text-gray-900">상세 정보</h3>
              </div>
              <button
                onClick={closeModal}
                className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
              >
                <X className="w-5 h-5 text-gray-500" />
              </button>
            </div>
            
            <div className="space-y-4 mb-6">
              <div className="bg-gradient-to-r from-blue-50 to-blue-50/50 rounded-xl p-4 border border-blue-100">
                <label className="text-xs font-bold text-blue-600 mb-1 block">📍 제품 코드</label>
                <p className="font-mono font-bold text-gray-900 text-sm">{selectedItem.productCode}</p>
              </div>
              <div className="bg-gradient-to-r from-purple-50 to-purple-50/50 rounded-xl p-4 border border-purple-100">
                <label className="text-xs font-bold text-purple-600 mb-1 block">📦 제품명</label>
                <p className="font-bold text-gray-900">{selectedItem.productName}</p>
              </div>
              <div className="bg-gradient-to-r from-green-50 to-green-50/50 rounded-xl p-4 border border-green-100">
                <label className="text-xs font-bold text-green-600 mb-1 block">📊 수량</label>
                <p className="font-bold text-gray-900 text-xl">{selectedItem.quantity.toLocaleString()}개</p>
              </div>
              <div className="bg-gradient-to-r from-orange-50 to-orange-50/50 rounded-xl p-4 border border-orange-100">
                <label className="text-xs font-bold text-orange-600 mb-1 block">📅 납기</label>
                <p className="font-bold text-gray-900">{selectedItem.deliveryDate || '미정'}</p>
              </div>
            </div>

            <div className="grid grid-cols-2 gap-3 mb-6 p-4 bg-gray-50 rounded-xl border border-gray-200">
              <div>
                <label className="text-xs font-bold text-gray-600 block mb-1">이동일</label>
                <p className="font-medium text-gray-900 text-sm">{selectedItem.transferDate || '-'}</p>
              </div>
              <div>
                <label className="text-xs font-bold text-gray-600 block mb-1">외주처</label>
                <p className="font-medium text-gray-900 text-sm">{selectedItem.assignedVendor || '-'}</p>
              </div>
              <div>
                <label className="text-xs font-bold text-gray-600 block mb-1">공정</label>
                <p className="font-medium text-gray-900 text-sm">{selectedItem.processType || '-'}</p>
              </div>
              <div>
                <label className="text-xs font-bold text-gray-600 block mb-1">담당자</label>
                <p className="font-medium text-gray-900 text-sm">{selectedItem.manager || '-'}</p>
              </div>
            </div>

            <button
              onClick={closeModal}
              className="w-full py-3 bg-gradient-to-r from-gray-900 to-gray-800 text-white rounded-xl hover:shadow-lg hover:-translate-y-0.5 transition-all font-bold"
            >
              닫기
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
