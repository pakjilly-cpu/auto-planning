'use client';

import { useMemo, useRef, useState } from 'react';
import { format, eachDayOfInterval, startOfMonth, endOfMonth, addMonths, subMonths, isToday } from 'date-fns';
import { ko } from 'date-fns/locale';
import { ChevronLeft, ChevronRight, X } from 'lucide-react';
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

export default function GanttChart() {
  const scrollRef = useRef<HTMLDivElement>(null);
  const { productionItems, selectedMonth, setSelectedMonth, vendors } = useAppStore();
  const [selectedItem, setSelectedItem] = useState<ProductionItem | null>(null);

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

  // 외주처별 라인별 일정 계산
  const vendorLineSchedules = useMemo(() => {
    const schedules: Record<string, Record<number, Record<string, typeof productionItems[0][]>>> = {};
    
    Object.entries(groupedByVendor).forEach(([vendorName, items]) => {
      const vendor = vendors.find(v => v.name === vendorName);
      const lineCount = vendor?.lineCount || 1;
      
      schedules[vendorName] = {};
      
      for (let line = 1; line <= lineCount; line++) {
        schedules[vendorName][line] = {};
        days.forEach(day => {
          schedules[vendorName][line][format(day, 'yyyy-MM-dd')] = [];
        });
      }
      
      // 아이템을 라인에 배분
      let currentLine = 1;
      items.forEach(item => {
        // 간단한 배분: 라운드 로빈
        const dateKey = format(days[0], 'yyyy-MM-dd'); // 임시로 첫날에 배치
        schedules[vendorName][currentLine][dateKey].push(item);
        currentLine = (currentLine % lineCount) + 1;
      });
    });
    
    return schedules;
  }, [groupedByVendor, days, vendors]);

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
      
      {/* 간트 차트 */}
      <div className="overflow-x-auto" ref={scrollRef}>
        <div className="min-w-max">
          {/* 날짜 헤더 */}
          <div className="flex border-b border-gray-200 sticky top-0 bg-gray-50 z-10">
            <div className="w-48 flex-shrink-0 p-3 font-medium text-gray-600 border-r border-gray-200">
              외주처 / 라인
            </div>
            {days.map((day, index) => {
              const isWeekend = day.getDay() === 0 || day.getDay() === 6;
              const isTodayDate = isToday(day);
              return (
                <div
                  key={index}
                  className={`
                    w-24 flex-shrink-0 p-2 text-center border-r border-gray-200 text-sm
                    ${isWeekend ? 'bg-red-50' : ''}
                    ${isTodayDate ? 'bg-blue-100 font-bold' : ''}
                  `}
                >
                  <div className="font-medium">{format(day, 'd')}</div>
                  <div className={`text-xs ${isWeekend ? 'text-red-500' : 'text-gray-400'}`}>
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
                  <div className="w-48 flex-shrink-0 p-3 border-r border-gray-200">
                    <div className="flex items-center gap-2">
                      <span className={`w-3 h-3 rounded-full ${vendorColors[vendorName] || 'bg-gray-400'}`} />
                      <span className="font-medium">{vendorName}</span>
                    </div>
                    <div className="text-xs text-gray-500 mt-1">
                      {items.length}건 / {totalQty.toLocaleString()}개
                    </div>
                  </div>
                  {days.map((_, idx) => (
                    <div key={idx} className="w-24 flex-shrink-0 border-r border-gray-200" />
                  ))}
                </div>
                
                {/* 라인별 행 */}
                {Array.from({ length: Math.min(lineCount, 3) }, (_, lineIdx) => {
                  const lineItems = items.filter((_, i) => i % lineCount === lineIdx);
                  
                  return (
                    <div key={lineIdx} className="flex hover:bg-gray-50">
                      <div className="w-48 flex-shrink-0 p-2 pl-6 border-r border-gray-200 text-sm text-gray-500">
                        라인 {lineIdx + 1}
                      </div>
                      {days.map((day, dayIdx) => {
                        // 각 날짜에 할당된 아이템 표시 (간단화)
                        const dayItems = lineItems.slice(dayIdx * 2, dayIdx * 2 + 2);
                        
                        return (
                          <div
                            key={dayIdx}
                            className="w-24 flex-shrink-0 p-1 border-r border-gray-200 min-h-[60px]"
                          >
                            {dayItems.map((item, itemIdx) => (
                              <div
                                key={itemIdx}
                                className={`
                                  text-xs p-1 rounded mb-1 truncate cursor-pointer
                                  border ${vendorBgColors[vendorName] || 'bg-gray-100 border-gray-300'}
                                  hover:opacity-80 transition-opacity select-none
                                `}
                                title="더블클릭하여 상세정보 보기"
                                onDoubleClick={() => handleItemDoubleClick(item)}
                              >
                                <div className="font-medium truncate">
                                  {item.productName.slice(0, 10)}...
                                </div>
                                <div className="text-gray-600">
                                  {item.quantity.toLocaleString()}
                                </div>
                              </div>
                            ))}
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
