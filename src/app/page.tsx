'use client';

import { useState } from 'react';
import { Download, Settings2, RefreshCw, Trash2 } from 'lucide-react';
import ExcelUpload from '@/components/ExcelUpload';
import StatCards from '@/components/StatCards';
import GanttChart from '@/components/GanttChart';
import VendorChart from '@/components/VendorChart';
import ClientMapping from '@/components/ClientMapping';
import VendorSettings from '@/components/VendorSettings';
import { useAppStore } from '@/lib/store';
import { exportToExcel } from '@/lib/excel';
import { allocateProduction } from '@/lib/allocation';

export default function Home() {
  const [showSettings, setShowSettings] = useState(false);
  const {
    productionItems,
    setProductionItems,
    setProductionPlans,
    vendors,
    clientMappings,
    selectedMonth,
    isLoading,
  } = useAppStore();

  // 재배분 실행
  const handleReAllocate = () => {
    if (productionItems.length === 0) return;
    
    // 모든 외주처 배정 초기화
    const resetItems = productionItems.map(item => ({
      ...item,
      assignedVendor: undefined,
    }));
    
    const { allocatedItems, plans } = allocateProduction(
      resetItems,
      vendors,
      clientMappings,
      selectedMonth
    );
    
    setProductionItems(allocatedItems);
    setProductionPlans(plans);
  };

  // 엑셀 내보내기
  const handleExport = () => {
    if (productionItems.length === 0) return;
    exportToExcel(productionItems);
  };

  // 데이터 삭제 (초기화)
  const handleClear = () => {
    if (confirm('업로드한 데이터를 삭제하시겠습니까?')) {
      setProductionItems([]);
      setProductionPlans([]);
    }
  };

  return (
    <div className="min-h-screen">
      {/* 헤더 */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-50">
        <div className="max-w-screen-2xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <h1 className="text-xl font-bold text-gray-900">
              외주처 생산계획 자동화
            </h1>
            
            <div className="flex items-center gap-3">
              {productionItems.length > 0 && (
                <>
                  <button
                    onClick={handleClear}
                    className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-red-600 bg-white border border-red-300 rounded-lg hover:bg-red-50 transition-colors"
                  >
                    <Trash2 className="w-4 h-4" />
                    삭제
                  </button>
                  <button
                    onClick={handleReAllocate}
                    className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                  >
                    <RefreshCw className="w-4 h-4" />
                    재배분
                  </button>
                  <button
                    onClick={handleExport}
                    className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-green-600 rounded-lg hover:bg-green-700 transition-colors"
                  >
                    <Download className="w-4 h-4" />
                    엑셀 다운로드
                  </button>
                </>
              )}
              <button
                onClick={() => setShowSettings(!showSettings)}
                className={`
                  flex items-center gap-2 px-4 py-2 text-sm font-medium rounded-lg transition-colors
                  ${showSettings 
                    ? 'text-white bg-blue-600 hover:bg-blue-700' 
                    : 'text-gray-700 bg-white border border-gray-300 hover:bg-gray-50'
                  }
                `}
              >
                <Settings2 className="w-4 h-4" />
                설정
              </button>
            </div>
          </div>
        </div>
      </header>

      {/* 메인 컨텐츠 */}
      <main className="max-w-screen-2xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {/* 로딩 오버레이 */}
        {isLoading && (
          <div className="fixed inset-0 bg-black/20 flex items-center justify-center z-50">
            <div className="bg-white rounded-xl p-6 shadow-xl flex items-center gap-3">
              <RefreshCw className="w-6 h-6 animate-spin text-blue-600" />
              <span className="font-medium">처리 중...</span>
            </div>
          </div>
        )}

        <div className="space-y-6">
          {/* 설정 패널 */}
          {showSettings && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              <ClientMapping />
              <VendorSettings />
            </div>
          )}

          {/* 엑셀 업로드 */}
          <ExcelUpload />

          {/* 통계 카드 */}
          <StatCards />

          {/* 차트 */}
          <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
            <VendorChart />
            <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
              <h3 className="text-lg font-semibold mb-4">사용 방법</h3>
              <ol className="space-y-3 text-sm text-gray-600">
                <li className="flex gap-2">
                  <span className="flex-shrink-0 w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs font-medium">1</span>
                  <span>엑셀 파일을 드래그하거나 클릭하여 업로드합니다.</span>
                </li>
                <li className="flex gap-2">
                  <span className="flex-shrink-0 w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs font-medium">2</span>
                  <span>시스템이 자동으로 외주처별 배분을 수행합니다.</span>
                </li>
                <li className="flex gap-2">
                  <span className="flex-shrink-0 w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs font-medium">3</span>
                  <span>아래 생산계획표에서 배분 결과를 확인합니다.</span>
                </li>
                <li className="flex gap-2">
                  <span className="flex-shrink-0 w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs font-medium">4</span>
                  <span>설정에서 고객사-외주처 매칭을 조정할 수 있습니다.</span>
                </li>
                <li className="flex gap-2">
                  <span className="flex-shrink-0 w-6 h-6 bg-blue-100 text-blue-600 rounded-full flex items-center justify-center text-xs font-medium">5</span>
                  <span>엑셀 다운로드 버튼으로 결과를 내보냅니다.</span>
                </li>
              </ol>
              
              <div className="mt-6 p-4 bg-amber-50 rounded-lg">
                <h4 className="font-medium text-amber-800 mb-2">배분 우선순위</h4>
                <ul className="text-sm text-amber-700 space-y-1">
                  <li>• 고객사 코드 기반 우선 매칭 (CLO→그램 등)</li>
                  <li>• 특수공정 가능 외주처 필터링</li>
                  <li>• 월간 목표 달성률 기준 배분</li>
                  <li>• 위드맘, 리니어, 그램 우선 배정</li>
                </ul>
              </div>
            </div>
          </div>

          {/* 간트 차트 (생산계획표) */}
          <GanttChart />
        </div>
      </main>
    </div>
  );
}
