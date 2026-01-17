'use client';

import { useState } from 'react';
import { Download, Settings2, RefreshCw, Trash2 } from 'lucide-react';
import ExcelUpload from '@/components/ExcelUpload';
import PerformanceMetrics from '@/components/PerformanceMetrics';
import StatCards from '@/components/StatCards';
import GanttChart from '@/components/GanttChart';
import VendorChart from '@/components/VendorChart';
import ClientMapping from '@/components/ClientMapping';
import VendorSettings from '@/components/VendorSettings';
import { useAppStore } from '@/lib/store';
import { exportToExcel } from '@/lib/excel';
import { 
  allocateProduction,
  allocateProductionOptimized,
  calculateSchedulingMetrics 
} from '@/lib/allocation';

export default function Home() {
  const [showSettings, setShowSettings] = useState(false);
  const {
    productionItems,
    setProductionItems,
    setProductionPlans,
    schedulingMetrics,
    setSchedulingMetrics,
    vendors,
    clientMappings,
    selectedMonth,
    isLoading,
  } = useAppStore();

  // 재배분 실행
  const handleReAllocate = () => {
    if (productionItems.length === 0) return;
    
    // 모든 외주처 배정 초기화
    const resetItems = productionItems.map((item: typeof productionItems[number]) => ({
      ...item,
      assignedVendor: undefined,
    }));
    
    const { allocatedItems, plans } = allocateProduction(
      resetItems,
      vendors,
      clientMappings,
      selectedMonth
    );
    
    // 성능 지표 계산
    const metrics = calculateSchedulingMetrics(plans, allocatedItems, vendors);
    
    setProductionItems(allocatedItems);
    setProductionPlans(plans);
    setSchedulingMetrics(metrics);
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
      setSchedulingMetrics(null);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-gray-50 via-blue-50/30 to-gray-50">
      {/* 헤더 - 모던 디자인 */}
      <header className="bg-gradient-to-r from-gray-900 via-gray-800 to-gray-900 border-b-2 border-blue-500 sticky top-0 z-50 shadow-xl">
        <div className="max-w-screen-2xl mx-auto px-6 py-5">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-4">
              <div className="w-12 h-12 bg-gradient-to-br from-blue-400 to-blue-600 rounded-xl shadow-lg flex items-center justify-center">
                <span className="text-white font-bold text-xl">📊</span>
              </div>
              <div>
                <h1 className="text-2xl font-bold text-white">생산 계획 자동 배정</h1>
                <p className="text-gray-300 text-sm mt-0.5">학술 알고리즘 기반 최적화 시스템</p>
              </div>
            </div>
            
            <div className="flex items-center gap-2">
              {productionItems.length > 0 && (
                <>
                  <button
                    onClick={handleClear}
                    className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold text-red-600 bg-white/80 backdrop-blur border-2 border-red-200 rounded-xl hover:bg-red-50 hover:shadow-lg transition-all hover:-translate-y-0.5"
                  >
                    <Trash2 className="w-4 h-4" />
                    삭제
                  </button>
                  <button
                    onClick={handleReAllocate}
                    className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold text-gray-700 bg-white/80 backdrop-blur border-2 border-gray-300 rounded-xl hover:bg-gray-50 hover:shadow-lg transition-all hover:-translate-y-0.5"
                  >
                    <RefreshCw className="w-4 h-4" />
                    재배분
                  </button>
                  <button
                    onClick={handleExport}
                    className="flex items-center gap-2 px-4 py-2.5 text-sm font-bold text-white bg-gradient-to-r from-green-600 to-emerald-600 rounded-xl hover:shadow-lg transition-all hover:-translate-y-0.5"
                  >
                    <Download className="w-4 h-4" />
                    엑셀 내보내기
                  </button>
                </>
              )}
              <button
                onClick={() => setShowSettings(!showSettings)}
                className={`
                  flex items-center gap-2 px-4 py-2.5 text-sm font-bold rounded-xl transition-all hover:shadow-lg hover:-translate-y-0.5
                  ${showSettings 
                    ? 'text-white bg-gradient-to-r from-blue-600 to-blue-700 shadow-lg' 
                    : 'text-white bg-gradient-to-r from-gray-700 to-gray-800 hover:from-gray-600 hover:to-gray-700'
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
      <main className="max-w-screen-2xl mx-auto px-6 py-8">
        {/* 로딩 오버레이 */}
        {isLoading && (
          <div className="fixed inset-0 bg-black/40 backdrop-blur-sm flex items-center justify-center z-50">
            <div className="bg-white rounded-2xl p-8 shadow-2xl flex items-center gap-4 border border-gray-100">
              <RefreshCw className="w-7 h-7 animate-spin text-blue-600" />
              <span className="font-bold text-lg text-gray-900">처리 중...</span>
            </div>
          </div>
        )}

        <div className="space-y-8">
          {/* 설정 패널 */}
          {showSettings && (
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6 animate-in fade-in slide-in-from-top-4 duration-300">
              <ClientMapping />
              <VendorSettings />
            </div>
          )}

          {/* 엑셀 업로드 */}
          <ExcelUpload />

          {/* 성능 지표 */}
          {productionItems.length > 0 && (
            <div className="bg-gradient-to-br from-white to-blue-50/30 rounded-2xl shadow-lg border border-gray-100 p-8 animate-in fade-in slide-in-from-bottom-4 duration-300">
              <div className="flex items-center gap-3 mb-6">
                <div className="w-10 h-10 bg-gradient-to-br from-blue-100 to-blue-50 rounded-xl flex items-center justify-center border border-blue-200">
                  <span className="text-xl">📊</span>
                </div>
                <h2 className="text-xl font-bold text-gray-900">스케줄링 성능 지표</h2>
              </div>
              <PerformanceMetrics 
                metrics={schedulingMetrics} 
                isLoading={isLoading}
              />
            </div>
          )}

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
