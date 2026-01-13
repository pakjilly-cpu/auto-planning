'use client';

import { useCallback, useState } from 'react';
import { Upload, FileSpreadsheet, AlertCircle } from 'lucide-react';
import { parseExcelFile } from '@/lib/excel';
import { useAppStore } from '@/lib/store';
import { allocateProduction } from '@/lib/allocation';

export default function ExcelUpload() {
  const [isDragging, setIsDragging] = useState(false);
  const [error, setError] = useState<string | null>(null);
  
  const {
    setProductionItems,
    setProductionPlans,
    vendors,
    clientMappings,
    selectedMonth,
    setIsLoading,
  } = useAppStore();

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.match(/\.(xlsx|xls|xlsm)$/i)) {
      setError('엑셀 파일(.xlsx, .xls, .xlsm)만 업로드 가능합니다.');
      return;
    }

    setError(null);
    setIsLoading(true);

    try {
      // 엑셀 파싱
      const items = await parseExcelFile(file);
      
      // 자동 배분 실행
      const { allocatedItems, plans } = allocateProduction(
        items,
        vendors,
        clientMappings,
        selectedMonth
      );
      
      setProductionItems(allocatedItems);
      setProductionPlans(plans);
    } catch (err) {
      setError('파일 처리 중 오류가 발생했습니다.');
      console.error(err);
    } finally {
      setIsLoading(false);
    }
  }, [vendors, clientMappings, selectedMonth, setProductionItems, setProductionPlans, setIsLoading]);

  const handleDrop = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
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
    if (file) handleFile(file);
  }, [handleFile]);

  return (
    <div className="w-full">
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
              .xlsx, .xls, .xlsm 파일 지원
            </p>
          </div>
        </div>
      </div>
      
      {error && (
        <div className="mt-3 flex items-center gap-2 text-red-600 text-sm">
          <AlertCircle className="w-4 h-4" />
          {error}
        </div>
      )}
    </div>
  );
}
