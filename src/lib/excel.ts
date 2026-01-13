import * as XLSX from 'xlsx';
import { ProductionItem, ProcessType, processTypeMap, EXCEL_COLUMNS } from './types';

// 제품 코드에서 고객사 코드 추출 (예: 9CLO1360310 -> CLO)
export function extractClientCode(productCode: string): string | undefined {
  if (!productCode || productCode.length < 4) return undefined;
  // 첫 번째 문자(숫자) 건너뛰고 3자리 영문 추출
  const match = productCode.match(/^[0-9]?([A-Z]{3})/);
  return match ? match[1] : undefined;
}

// 특수 공정 파싱
export function parseSpecialProcess(value: string | undefined): ProcessType {
  if (!value) return 'normal';
  const trimmed = value.trim();
  return processTypeMap[trimmed] || 'normal';
}

// 수량 파싱 (콤마 제거, 숫자만 추출)
export function parseQuantity(value: string | number | undefined): number {
  if (value === undefined || value === null || value === '') return 0;
  if (typeof value === 'number') return value;
  // 숫자만 추출
  const numStr = value.toString().replace(/[^0-9]/g, '');
  return parseInt(numStr, 10) || 0;
}

// 엑셀 파일 파싱
export async function parseExcelFile(file: File): Promise<ProductionItem[]> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        const workbook = XLSX.read(data, { type: 'array' });
        
        // "계획" 시트 우선 사용, 없으면 첫 번째 시트
        const targetSheetName = workbook.SheetNames.find(name => name === '계획') || workbook.SheetNames[0];
        const worksheet = workbook.Sheets[targetSheetName];
        
        if (!worksheet) {
          reject(new Error('시트를 찾을 수 없습니다. "계획" 시트가 있는지 확인해주세요.'));
          return;
        }
        
        // JSON으로 변환 (헤더 없이 배열로)
        const rows: (string | number)[][] = XLSX.utils.sheet_to_json(worksheet, {
          header: 1,
          raw: false,
        });
        
        // 첫 번째 행(헤더) 제외하고 파싱
        const items: ProductionItem[] = [];
        
        for (let i = 1; i < rows.length; i++) {
          const row = rows[i];
          if (!row || row.length === 0) continue;
          
          // 제품코드가 없으면 건너뛰기
          const productCode = row[EXCEL_COLUMNS.PRODUCT_CODE]?.toString().trim();
          if (!productCode) continue;
          
          // 수량이 0이면 건너뛰기
          const quantity = parseQuantity(row[EXCEL_COLUMNS.QUANTITY]);
          if (quantity === 0) continue;
          
          const item: ProductionItem = {
            id: `item-${i}-${Date.now()}`,
            transferDate: row[EXCEL_COLUMNS.TRANSFER_DATE]?.toString().trim(),
            status: row[EXCEL_COLUMNS.STATUS]?.toString().trim(),
            specialProcess: parseSpecialProcess(row[EXCEL_COLUMNS.SPECIAL_PROCESS]?.toString()),
            assignedVendor: row[EXCEL_COLUMNS.VENDOR]?.toString().trim() || undefined,
            manager: row[EXCEL_COLUMNS.MANAGER]?.toString().trim(),
            productCode,
            productName: row[EXCEL_COLUMNS.PRODUCT_NAME]?.toString().trim() || '',
            processType: row[EXCEL_COLUMNS.PROCESS_TYPE]?.toString().trim() || '',
            quantity,
            manufacturingDate: row[EXCEL_COLUMNS.MFG_DATE]?.toString().trim(),
            materialArrival: row[EXCEL_COLUMNS.MATERIAL_DATE]?.toString().trim(),
            deliveryDate: row[EXCEL_COLUMNS.DELIVERY_DATE]?.toString().trim(),
            urgency: row[EXCEL_COLUMNS.URGENCY]?.toString().trim(),
            nightShift: row[EXCEL_COLUMNS.NIGHT_SHIFT]?.toString().trim(),
            remarks: row[EXCEL_COLUMNS.REMARKS]?.toString().trim(),
            clientCode: extractClientCode(productCode),
          };
          
          items.push(item);
        }
        
        resolve(items);
      } catch (error) {
        reject(error);
      }
    };
    
    reader.onerror = () => reject(new Error('파일 읽기 실패'));
    reader.readAsArrayBuffer(file);
  });
}

// 배분 결과를 엑셀로 내보내기
export function exportToExcel(items: ProductionItem[]): void {
  const exportData = items.map((item) => ({
    '이동일': item.transferDate || '',
    '상태': item.status || '',
    '특수공정': item.specialProcess === 'normal' ? '' : 
      item.specialProcess === 'shrink' ? '수축' :
      item.specialProcess === 'mixing' ? '교반' : '고주파',
    '외주처': item.assignedVendor || '',
    '담당자': item.manager || '',
    '제품코드': item.productCode,
    '제품명': item.productName,
    '공정': item.processType,
    '수량': item.quantity,
    '제조일': item.manufacturingDate || '',
    '자재입고': item.materialArrival || '',
    '납기일': item.deliveryDate || '',
    '긴급': item.urgency || '',
    '야상': item.nightShift || '',
    '비고': item.remarks || '',
  }));
  
  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '생산계획');
  
  // 다운로드
  XLSX.writeFile(workbook, `생산계획_${new Date().toISOString().slice(0, 10)}.xlsx`);
}
