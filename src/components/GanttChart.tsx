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

// 외주처별 색상
const vendorColors: Record<string, string> = {
  '위드맘': 'bg-blue-500',
  '리니어': 'bg-green-500',
  '그램': 'bg-purple-500',
  '이시스': 'bg-orange-500',
  '엘루오': 'bg-pink-500',
  '케이코스텍': 'bg-cyan-500',
  '다미': 'bg-amber-500',
  '배정불가': 'bg-gray-400',
};

const vendorBgColors: Record<string, string> = {
  '위드맘': 'bg-blue-100 border-blue-300',
  '리니어': 'bg-green-100 border-green-300',
  '그램': 'bg-purple-100 border-purple-300',
  '이시스': 'bg-orange-100 border-orange-300',
  '엘루오': 'bg-pink-100 border-pink-300',
  '케이코스텍': 'bg-cyan-100 border-cyan-300',
  '다미': 'bg-amber-100 border-amber-300',
  '배정불가': 'bg-gray-100 border-gray-300',
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
    
    // 모든 외주처에 대해 스케줄 초기화
    const vendorOrder = ['위드맘', '리니어', '그램', '이시스', '엘루오', '케이코스텍', '다미', '배정불가'];
    vendorOrder.forEach(vendorName => {
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
          
          if (schedules[vendorName][currentLine][dateKey]) {
            schedules[vendorName][currentLine][dateKey].push({
              item,
              dailyQty,
              dayNumber: dayOffset + 1,
              totalDays,
            });
            // 점유된 날짜로 마킹
            lineOccupiedDays[currentLine].add(dateKey);
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
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 overflow-hidden">
      {/* 헤더 */}
      <div className="flex items-center justify-between p-4 border-b border-gray-100">
        <h2 className="text-lg font-semibold">생산계획표</h2>
        <div className="flex items-center gap-4">
          {/* 줌 컨트롤 */}
          <div className="flex items-center gap-1 bg-gray-100 rounded-lg p-1">
            <button
              onClick={handleZoomOut}
              disabled={zoomLevel === ZOOM_LEVELS[0]}
              className="p-1.5 hover:bg-gray-200 rounded disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
              title="축소"
            >
              <ZoomOut className="w-4 h-4" />
            </button>
            <span className="text-sm font-medium min-w-[45px] text-center">{zoomLevel}%</span>
            <button
              onClick={handleZoomIn}
              disabled={zoomLevel === ZOOM_LEVELS[ZOOM_LEVELS.length - 1]}
              className="p-1.5 hover:bg-gray-200 rounded disabled:opacity-40 disabled:cursor-not-allowed transition-colors"
              title="확대"
            >
              <ZoomIn className="w-4 h-4" />
            </button>
          </div>
          
          {/* 월 이동 */}
          <div className="flex items-center gap-2">
            <button
              onClick={handlePrevMonth}
              className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
            >
              <ChevronLeft className="w-5 h-5" />
            </button>
            <span className="font-medium min-w-[120px] text-center">
              {format(selectedMonth, 'yyyy년 M월', { locale: ko })}
            </span>
            <button
              onClick={handleNextMonth}
              className="p-2 hover:bg-gray-100 rounded-lg transition-colors"
            >
              <ChevronRight className="w-5 h-5" />
            </button>
          </div>
        </div>
      </div>
      
      {/* 간트 차트 */}
      <div className="overflow-x-auto max-h-[70vh]" ref={scrollRef} onScroll={handleScroll}>
        <div className="min-w-max" style={{ fontSize: `${zoomLevel}%` }}>
          {/* 날짜 헤더 */}
          <div className="flex border-b border-gray-200 sticky top-0 bg-gray-50 z-20">
            <div 
              className="flex-shrink-0 p-2 font-medium text-gray-600 border-r border-gray-200 bg-gray-50 sticky left-0 z-30"
              style={{ width: `${fixedColumnWidth}px` }}
            >
              외주처 / 라인
            </div>
            {days.map((day, index) => {
              const isWeekendDay = day.getDay() === 0 || day.getDay() === 6;
              const isTodayDate = isToday(day);
              return (
                <div
                  key={index}
                  className={`
                    flex-shrink-0 p-1 text-center border-r border-gray-200 ${fontSize}
                    ${isWeekendDay ? 'bg-red-50' : 'bg-gray-50'}
                    ${isTodayDate ? 'bg-blue-100 font-bold' : ''}
                  `}
                  style={{ width: `${cellWidth}px` }}
                >
                  <div className="font-medium">{format(day, 'd')}</div>
                  <div className={`text-[10px] ${isWeekendDay ? 'text-red-500' : 'text-gray-400'}`}>
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
            
            return (
              <div key={vendorName} className="border-b border-gray-200">
                {/* 외주처 헤더 */}
                <div className="flex bg-gray-50">
                  <div 
                    className="flex-shrink-0 p-2 border-r border-gray-200 bg-gray-50 sticky left-0 z-10"
                    style={{ width: `${fixedColumnWidth}px` }}
                  >
                    <div className="flex items-center gap-2">
                      <span className={`w-2 h-2 rounded-full ${vendorColors[vendorName] || 'bg-gray-400'}`} />
                      <span className={`font-medium ${fontSize}`}>{vendorName}</span>
                    </div>
                    <div className="text-[10px] text-gray-500 mt-0.5">
                      {items.length}건 / {totalQty.toLocaleString()}개
                    </div>
                  </div>
                  {days.map((_, idx) => (
                    <div 
                      key={idx} 
                      className="flex-shrink-0 border-r border-gray-200" 
                      style={{ width: `${cellWidth}px` }}
                    />
                  ))}
                </div>
                
                {/* 라인별 행 */}
                {Array.from({ length: lineCount }, (_, lineIdx) => {
                  const lineNumber = lineIdx + 1;
                  
                  return (
                    <div key={lineIdx} className="flex hover:bg-gray-50">
                      <div 
                        className={`flex-shrink-0 p-1 pl-4 border-r border-gray-200 ${fontSize} text-gray-500 bg-white sticky left-0 z-10`}
                        style={{ width: `${fixedColumnWidth}px` }}
                      >
                        라인 {lineNumber}
                      </div>
                      {days.map((day, dayIdx) => {
                        // vendorLineSchedules에서 해당 날짜의 아이템 가져오기
                        const dateKey = format(day, 'yyyy-MM-dd');
                        const dayItems = vendorLineSchedules[vendorName]?.[lineNumber]?.[dateKey] || [];
                        const minHeight = zoomLevel < 75 ? 40 : zoomLevel < 100 ? 50 : 60;
                        const isDropTarget = dropTarget?.vendorName === vendorName && 
                                             dropTarget?.lineNumber === lineNumber && 
                                             dropTarget?.dateKey === dateKey;
                        const isWeekendDay = isWeekend(day);
                        
                        return (
                          <div
                            key={dayIdx}
                            className={`
                              flex-shrink-0 p-0.5 border-r border-gray-200 transition-colors
                              ${isDropTarget ? 'bg-blue-100 ring-2 ring-blue-400 ring-inset' : ''}
                              ${isWeekendDay ? 'bg-red-50/50' : ''}
                              ${dragItem && !isWeekendDay ? 'hover:bg-blue-50' : ''}
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
                                  draggable
                                  onDragStart={(e) => {
                                    e.dataTransfer.effectAllowed = 'move';
                                    handleDragStart(dailyItem.item, vendorName, lineNumber, dateKey);
                                  }}
                                  onDragEnd={handleDragEnd}
                                  className={`
                                    p-0.5 rounded mb-0.5 cursor-grab active:cursor-grabbing
                                    border ${vendorBgColors[vendorName] || 'bg-gray-100 border-gray-300'}
                                    hover:opacity-80 transition-opacity select-none
                                    ${zoomLevel < 75 ? 'text-[8px]' : zoomLevel < 100 ? 'text-[10px]' : 'text-xs'}
                                    ${dragItem?.item.productCode === dailyItem.item.productCode ? 'opacity-50 ring-2 ring-blue-400' : ''}
                                    ${isFixed ? 'ring-1 ring-yellow-400' : ''}
                                  `}
                                  title={isFixed ? '고정된 품목 (드래그하여 이동)' : '드래그하여 이동, 더블클릭하여 상세정보'}
                                  onDoubleClick={() => handleItemDoubleClick(dailyItem.item)}
                                >
                                  <div className="flex items-center gap-0.5">
                                    <GripVertical className="w-2 h-2 text-gray-400 flex-shrink-0" />
                                    <span className="font-medium truncate">
                                      {dailyItem.item.productName.slice(0, nameLength)}..
                                    </span>
                                  </div>
                                  <div className="text-gray-600">
                                    {dailyItem.dailyQty.toLocaleString()}
                                  </div>
                                  {dailyItem.totalDays > 1 && (
                                    <div className="text-gray-400" style={{ fontSize: '8px' }}>
                                      ({dailyItem.dayNumber}/{dailyItem.totalDays})
                                    </div>
                                  )}
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
          className="fixed inset-0 bg-black/50 flex items-center justify-center z-50"
          onClick={cancelChange}
        >
          <div 
            className="bg-white rounded-xl shadow-2xl p-6 max-w-lg w-full mx-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center gap-3 mb-4">
              <div className="w-10 h-10 bg-blue-100 rounded-full flex items-center justify-center">
                <Check className="w-5 h-5 text-blue-600" />
              </div>
              <h3 className="text-lg font-semibold">이대로 계획을 확정하시겠습니까?</h3>
            </div>
            
            <div className="bg-gray-50 rounded-lg p-4 space-y-3">
              <div className="flex items-center gap-2">
                <span className="text-sm font-medium text-gray-500 w-16">품목:</span>
                <span className="font-medium">{pendingChange.item.productName}</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-sm font-medium text-gray-500 w-16">코드:</span>
                <span className="font-mono text-sm">{pendingChange.item.productCode}</span>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-sm font-medium text-gray-500 w-16">수량:</span>
                <span>{pendingChange.item.quantity.toLocaleString()}개</span>
              </div>
              
              <div className="border-t border-gray-200 pt-3 mt-3">
                <div className="grid grid-cols-2 gap-4">
                  <div className="bg-red-50 rounded-lg p-3">
                    <div className="text-xs text-red-500 font-medium mb-1">변경 전</div>
                    <div className="text-sm font-medium">{pendingChange.from.vendor}</div>
                    <div className="text-xs text-gray-600">라인 {pendingChange.from.line}</div>
                    <div className="text-xs text-gray-600">{pendingChange.from.date}</div>
                  </div>
                  <div className="bg-green-50 rounded-lg p-3">
                    <div className="text-xs text-green-500 font-medium mb-1">변경 후</div>
                    <div className="text-sm font-medium">{pendingChange.to.vendor}</div>
                    <div className="text-xs text-gray-600">라인 {pendingChange.to.line}</div>
                    <div className="text-xs text-gray-600">{pendingChange.to.date}</div>
                  </div>
                </div>
              </div>
              
              <p className="text-xs text-gray-500 mt-2">
                * 확정하면 이 품목은 다음 엑셀 업로드 시에도 지정한 위치에 고정 배치됩니다.
              </p>
            </div>

            <div className="flex gap-3 mt-6">
              <button
                onClick={cancelChange}
                className="flex-1 py-2.5 border border-gray-300 text-gray-700 rounded-lg hover:bg-gray-50 transition-colors font-medium"
              >
                취소
              </button>
              <button
                onClick={confirmChange}
                className="flex-1 py-2.5 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors font-medium"
              >
                변경
              </button>
            </div>
          </div>
        </div>
      )}

      {/* 품목 상세 모달 */}
      {selectedItem && (
        <div 
          className="fixed inset-0 bg-black/50 flex items-center justify-center z-50"
          onClick={closeModal}
        >
          <div 
            className="bg-white rounded-xl shadow-2xl p-6 max-w-md w-full mx-4"
            onClick={(e) => e.stopPropagation()}
          >
            <div className="flex items-center justify-between mb-4">
              <h3 className="text-lg font-semibold">품목 상세정보</h3>
              <button
                onClick={closeModal}
                className="p-1 hover:bg-gray-100 rounded-lg transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="space-y-4">
              <div>
                <label className="text-sm text-gray-500">제품코드</label>
                <p className="font-mono font-medium text-gray-900">{selectedItem.productCode}</p>
              </div>
              <div>
                <label className="text-sm text-gray-500">제품명</label>
                <p className="font-medium text-gray-900">{selectedItem.productName}</p>
              </div>
              <div>
                <label className="text-sm text-gray-500">수량</label>
                <p className="font-medium text-gray-900 text-lg">{selectedItem.quantity.toLocaleString()}개</p>
              </div>
              <div>
                <label className="text-sm text-gray-500">납기</label>
                <p className="font-medium text-gray-900">{selectedItem.deliveryDate || '-'}</p>
              </div>
            </div>

            <div className="mt-6 pt-4 border-t border-gray-100">
              <div className="grid grid-cols-2 gap-4 text-sm">
                <div>
                  <label className="text-gray-500">이동일</label>
                  <p className="font-medium">{selectedItem.transferDate || '-'}</p>
                </div>
                <div>
                  <label className="text-gray-500">외주처</label>
                  <p className="font-medium">{selectedItem.assignedVendor || '-'}</p>
                </div>
                <div>
                  <label className="text-gray-500">공정</label>
                  <p className="font-medium">{selectedItem.processType || '-'}</p>
                </div>
                <div>
                  <label className="text-gray-500">담당자</label>
                  <p className="font-medium">{selectedItem.manager || '-'}</p>
                </div>
              </div>
            </div>

            <button
              onClick={closeModal}
              className="mt-6 w-full py-2 bg-gray-900 text-white rounded-lg hover:bg-gray-800 transition-colors"
            >
              닫기
            </button>
          </div>
        </div>
      )}
    </div>
  );
}
