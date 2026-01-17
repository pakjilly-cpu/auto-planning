'use client';

import { useCallback, useState } from 'react';
import { Upload, FileSpreadsheet, AlertCircle, Zap, Gauge } from 'lucide-react';
import { parseExcelFile } from '@/lib/excel';
import { useAppStore } from '@/lib/store';
import { 
  allocateProduction,
  allocateProductionOptimized,
  calculateSchedulingMetrics 
} from '@/lib/allocation';

export default function ExcelUpload() {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [optimizationOptions, setOptimizationOptions] = useState({
    useLPT: true,
    balanceLoad: true,
    considerUrgency: true,
  });
  const [showAdvanced, setShowAdvanced] = useState(false);
  
  const {
    setProductionItems,
    setProductionPlans,
    setSchedulingMetrics,
    vendors,
    clientMappings,
    selectedMonth,
    setIsLoading,
  } = useAppStore();

  const handleFile = useCallback(async (file: File, useOptimized: boolean = true) => {
    if (!file.name.match(/\.(xlsx|xls|xlsm)$/i)) {
      setError('엑셀 파일(.xlsx, .xls, .xlsm)만 업로드 가능합니다.');
      return;
    }

    setError(null);
    setIsLoading(true);

    try {
      // 엑셀 파싱
      const items = await parseExcelFile(file);
      
      // 배분 실행
      let allocatedItems, plans;
      
      if (useOptimized) {
        // 최적화된 배분 실행
        const result = allocateProductionOptimized(
          items,
          vendors,
          clientMappings,
          selectedMonth,
          optimizationOptions
        );
        allocatedItems = result.allocatedItems;
        plans = result.plans;
      } else {
        // 기본 배분 실행
        const result = allocateProduction(
          items,
          vendors,
          clientMappings,
          selectedMonth
        );
        allocatedItems = result.allocatedItems;
        plans = result.plans;
      }
      
      // 성능 지표 계산
      const metrics = calculateSchedulingMetrics(plans, allocatedItems, vendors);
      
      setProductionItems(allocatedItems);
      setProductionPlans(plans);
      setSchedulingMetrics(metrics);
    } catch (err) {
      setError('파일 처리 중 오류가 발생했습니다.');
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [vendors, clientMappings, selectedMonth, setProductionItems, setProductionPlans, setSchedulingMetrics, setIsLoading, optimizationOptions]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file, true);
  }, [handleFile]);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  }, []);

  const handleDragLeave = useCallback(() => {
    setIsDragging(false);
  }, []);

  const handleInputChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file, true);
  }, [handleFile]);

  const handleBasicUpload = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file, false);
  }, [handleFile]);

  return (
    <div className="w-full space-y-4">
      <div
        onDrop={handleDrop}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        className={`
          relative border-2 border-dashed rounded-xl p-8 text-center transition-all
          ${isDragging 
            ? 'border-blue-500 bg-blue-50' 
            : 'border-gray-300 hover:border-gray-400 bg-gray-50'
          }
        `}
      >
        <input
          type="file"
          accept=".xlsx,.xls,.xlsm"
          onChange={handleInputChange}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        />
        
        <div className="flex flex-col items-center gap-3">
          {isDragging ? (
            <FileSpreadsheet className="w-12 h-12 text-blue-500" />
          ) : (
            <Upload className="w-12 h-12 text-gray-400" />
          )}
          
          <div>
            <p className="text-lg font-medium text-gray-700">
              {isDragging ? '여기에 놓으세요' : '엑셀 파일을 드래그하거나 클릭하여 업로드'}
            </p>
            <p className="text-sm text-gray-500 mt-1">
              .xlsx, .xls, .xlsm 파일 지원 (최적화 알고리즘 자동 적용)
            </p>
          </div>
        </div>
      </div>

      {/* 고급 옵션 */}
      <div className="border border-gray-200 rounded-lg p-4 bg-white">
        <button
          onClick={() => setShowAdvanced(!showAdvanced)}
          className="flex items-center gap-2 text-sm font-medium text-gray-700 hover:text-gray-900"
        >
          <Gauge className="w-4 h-4" />
          고급 최적화 옵션
          <span className="text-xs text-gray-500">
            {showAdvanced ? '▼' : '▶'}
          </span>
        </button>

        {showAdvanced && (
          <div className="mt-4 space-y-3 pt-4 border-t border-gray-200">
            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={optimizationOptions.useLPT}
                onChange={(e) =>
                  setOptimizationOptions({
                    ...optimizationOptions,
                    useLPT: e.target.checked,
                  })
                }
                className="rounded border-gray-300"
              />
              <span className="text-sm font-medium text-gray-700">
                LPT (Longest Processing Time) 규칙
                <p className="text-xs text-gray-500 font-normal mt-1">
                  처리량이 큰 품목부터 우선 배정하여 완료시간 단축
                </p>
              </span>
            </label>

            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={optimizationOptions.balanceLoad}
                onChange={(e) =>
                  setOptimizationOptions({
                    ...optimizationOptions,
                    balanceLoad: e.target.checked,
                  })
                }
                className="rounded border-gray-300"
              />
              <span className="text-sm font-medium text-gray-700">
                부하 균형화 (Load Balancing)
                <p className="text-xs text-gray-500 font-normal mt-1">
                  외주처별 부하를 균등하게 분산하여 효율성 향상
                </p>
              </span>
            </label>

            <label className="flex items-center gap-3 cursor-pointer">
              <input
                type="checkbox"
                checked={optimizationOptions.considerUrgency}
                onChange={(e) =>
                  setOptimizationOptions({
                    ...optimizationOptions,
                    considerUrgency: e.target.checked,
                  })
                }
                className="rounded border-gray-300"
              />
              <span className="text-sm font-medium text-gray-700">
                납기일 우선순위
                <p className="text-xs text-gray-500 font-normal mt-1">
                  납기일이 임박한 품목을 우선으로 처리
                </p>
              </span>
            </label>

            <div className="flex gap-2 pt-2">
              <input
                type="file"
                accept=".xlsx,.xls,.xlsm"
                onChange={handleBasicUpload}
                className="hidden"
                id="basic-upload"
              />
              <label
                htmlFor="basic-upload"
                className="flex-1 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 cursor-pointer text-center"
              >
                기본 알고리즘으로 업로드
              </label>
            </div>
          </div>
        )}
      </div>
      
      {error && (
        <div className="flex items-center gap-2 text-red-600 text-sm bg-red-50 p-3 rounded-lg">
          <AlertCircle className="w-4 h-4 flex-shrink-0" />
          {error}
        </div>
      )}

      {/* 알고리즘 정보 */}
      <div className="bg-blue-50 border border-blue-200 rounded-lg p-4 text-sm">
        <div className="flex items-start gap-2">
          <Zap className="w-4 h-4 text-blue-600 mt-0.5 flex-shrink-0" />
          <div>
            <p className="font-medium text-blue-900">현재 적용된 최적화 기법</p>
            <ul className="text-blue-700 mt-2 space-y-1">
              <li>
                ✓ <strong>LPT 규칙</strong> - 처리량 기반 정렬로 완료시간 최소화
              </li>
              <li>
                ✓ <strong>Min-Max 부하분산</strong> - 외주처 간 균형잡힌 배분
              </li>
              <li>
                ✓ <strong>동적 우선순위</strong> - 납기일, 고객사, 특수공정 고려
              </li>
              <li>
                ✓ <strong>성능 지표</strong> - Makespan, 활용률, 목표달성률 등 자동 계산
              </li>
            </ul>
          </div>
        </div>
      </div>
    </div>
  );
}
