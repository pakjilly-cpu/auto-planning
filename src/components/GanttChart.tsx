'use client';

import { useMemo, useRef, useState } from 'react';
import { format, eachDayOfInterval, startOfMonth, endOfMonth, addMonths, subMonths, isToday, addDays, isWeekend } from 'date-fns';
import { ko } from 'date-fns/locale';
import { ChevronLeft, ChevronRight, X, ZoomIn, ZoomOut } from 'lucide-react';
import { useAppStore } from '@/lib/store';
import { ProductionItem } from '@/lib/types';

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
  const { productionItems, selectedMonth, setSelectedMonth, vendors } = useAppStore();
  const [selectedItem, setSelectedItem] = useState<ProductionItem | null>(null);
  const [zoomLevel, setZoomLevel] = useState(100);

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
    const groups: Record<string, typeof productionItems> = {};
    
    // 외주처 순서 정의 (우선순위 순)
    const vendorOrder = ['위드맘', '리니어', '그램', '이시스', '엘루오', '케이코스텍', '다미', '배정불가'];
    
    vendorOrder.forEach(vendor => {
      groups[vendor] = [];
    });
    
    productionItems.forEach(item => {
      const vendor = item.assignedVendor || '배정불가';
      if (!groups[vendor]) groups[vendor] = [];
      groups[vendor].push(item);
    });
    
    return groups;
  }, [productionItems]);

  // 일별 생산 품목 (CAPA 반영)
  interface DailyItem {
    item: ProductionItem;
    dailyQty: number;  // 해당 날짜에 생산할 수량
    dayNumber: number; // 몇 번째 날인지 (1, 2, 3...)
    totalDays: number; // 총 며칠 걸리는지
  }

  // 외주처별 라인별 일정 계산 (CAPA 반영하여 여러 날에 분배)
  const vendorLineSchedules = useMemo(() => {
    const schedules: Record<string, Record<number, Record<string, DailyItem[]>>> = {};
    
    // 주말 제외 근무일만 필터링
    const workingDays = days.filter(day => !isWeekend(day));
    
    Object.entries(groupedByVendor).forEach(([vendorName, items]) => {
      const vendor = vendors.find(v => v.name === vendorName);
      const lineCount = vendor?.lineCount || 1;
      const dailyCapa = vendor?.dailyCapacityPerLine || 10000;
      
      schedules[vendorName] = {};
      
      // 라인별 다음 가용 날짜 인덱스 추적
      const lineNextDayIndex: Record<number, number> = {};
      
      for (let line = 1; line <= lineCount; line++) {
        schedules[vendorName][line] = {};
        lineNextDayIndex[line] = 0;
        days.forEach(day => {
          schedules[vendorName][line][format(day, 'yyyy-MM-dd')] = [];
        });
      }
      
      const targetYear = selectedMonth.getFullYear();
      const monthStart = startOfMonth(selectedMonth);
      const monthEnd = endOfMonth(selectedMonth);
      
      // 라인별로 품목 분배
      let currentLine = 1;
      items.forEach(item => {
        // 이동일 기반 생산 시작일 계산
        const productionStartDate = getProductionStartDate(item.transferDate, targetYear, workingDays[0] || days[0]);
        
        // 해당 월 범위 체크
        if (productionStartDate > monthEnd) return;
        
        // 생산 시작일 이후의 근무일 찾기
        let startDayIndex = workingDays.findIndex(d => d >= productionStartDate);
        if (startDayIndex === -1) startDayIndex = 0;
        
        // 라인의 현재 가용 날짜와 비교하여 더 늦은 날짜 사용
        const lineAvailableIndex = lineNextDayIndex[currentLine];
        const actualStartIndex = Math.max(startDayIndex, lineAvailableIndex);
        
        // 필요한 일수 계산
        const totalDays = Math.ceil(item.quantity / dailyCapa);
        
        // 여러 날에 걸쳐 배분
        for (let dayOffset = 0; dayOffset < totalDays; dayOffset++) {
          const dayIndex = actualStartIndex + dayOffset;
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
          }
        }
        
        // 해당 라인의 다음 가용 날짜 업데이트
        lineNextDayIndex[currentLine] = actualStartIndex + totalDays;
        
        currentLine = (currentLine % lineCount) + 1;
      });
    });
    
    return schedules;
  }, [groupedByVendor, days, vendors, selectedMonth]);

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
      <div className="overflow-x-auto max-h-[70vh]" ref={scrollRef}>
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
            
            const vendor = vendors.find(v => v.name === vendorName);
            const lineCount = vendor?.lineCount || 1;
            const totalQty = items.reduce((sum, item) => sum + item.quantity, 0);
            
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
                {Array.from({ length: Math.min(lineCount, 3) }, (_, lineIdx) => {
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
                        
                        return (
                          <div
                            key={dayIdx}
                            className="flex-shrink-0 p-0.5 border-r border-gray-200"
                            style={{ width: `${cellWidth}px`, minHeight: `${minHeight}px` }}
                          >
                            {dayItems.map((dailyItem, itemIdx) => {
                              const nameLength = zoomLevel < 75 ? 4 : zoomLevel < 100 ? 6 : 8;
                              return (
                                <div
                                  key={itemIdx}
                                  className={`
                                    p-0.5 rounded mb-0.5 cursor-pointer
                                    border ${vendorBgColors[vendorName] || 'bg-gray-100 border-gray-300'}
                                    hover:opacity-80 transition-opacity select-none
                                    ${zoomLevel < 75 ? 'text-[8px]' : zoomLevel < 100 ? 'text-[10px]' : 'text-xs'}
                                  `}
                                  title="더블클릭하여 상세정보 보기"
                                  onDoubleClick={() => handleItemDoubleClick(dailyItem.item)}
                                >
                                  <div className="font-medium truncate">
                                    {dailyItem.item.productName.slice(0, nameLength)}..
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
