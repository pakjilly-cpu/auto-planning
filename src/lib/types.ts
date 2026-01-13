// 외주처 정보
export interface Vendor {
  id: string;
  name: string;
  lineCount: number;
  dailyCapacityPerLine: number;
  capabilities: ProcessType[];
  monthlyTarget?: number; // 월간 목표 생산량
  priority: number; // 우선순위 (낮을수록 높음)
}

// 특수 공정 타입
export type ProcessType = 'normal' | 'shrink' | 'mixing' | 'highFrequency';

// 공정 타입 한글 매핑
export const processTypeMap: Record<string, ProcessType> = {
  '수축': 'shrink',
  '교반': 'mixing',
  '고주파': 'highFrequency',
};

// 고객사-외주처 매칭
export interface ClientVendorMapping {
  clientCode: string; // 3자리 영문 코드 (CLO, ERK, DPD 등)
  vendorId: string;
  priority: number;
}

// 엑셀에서 파싱한 생산 품목
export interface ProductionItem {
  id: string;
  transferDate?: string; // A열 - 이동일 (자재를 생산처에 가져다 주는 날)
  status?: string; // B열
  specialProcess?: ProcessType; // C열 (수축, 교반, 고주파)
  assignedVendor?: string; // D열 (배정될 외주처)
  manager?: string; // E열
  productCode: string; // F열 (9CLO1360310 형식)
  productName: string; // G열
  processType: string; // H열 (충전, 포장, 충포장 등)
  quantity: number; // I열
  manufacturingDate?: string; // J열
  materialArrival?: string; // K열
  deliveryDate?: string; // L열
  urgency?: string; // M열
  nightShift?: string; // N열
  remarks?: string; // O열
  clientCode?: string; // F열에서 추출 (3자리 영문)
}

// 생산 계획 (배분 결과)
export interface ProductionPlan {
  id: string;
  productionItemId: string;
  vendorId: string;
  vendorName: string;
  lineNumber: number;
  startDate: Date;
  endDate: Date;
  dailyQuantity: number;
  totalQuantity: number;
  productName: string;
  productCode: string;
  status: 'planned' | 'in_progress' | 'completed';
}

// 외주처별 일일 생산 현황
export interface DailyProduction {
  date: Date;
  vendorId: string;
  lineNumber: number;
  productionPlanId?: string;
  quantity: number;
  isAvailable: boolean;
}

// 대시보드 통계
export interface DashboardStats {
  totalPlannedQuantity: number;
  vendorStats: VendorStat[];
  dailyTotals: { date: string; quantity: number }[];
}

export interface VendorStat {
  vendorId: string;
  vendorName: string;
  totalQuantity: number;
  monthlyTarget: number;
  achievementRate: number;
  lineUtilization: number;
}

// 엑셀 컬럼 매핑
export const EXCEL_COLUMNS = {
  TRANSFER_DATE: 0,  // A열 - 이동일
  STATUS: 1,         // B열
  SPECIAL_PROCESS: 2,// C열
  VENDOR: 3,         // D열
  MANAGER: 4,        // E열
  PRODUCT_CODE: 5,   // F열
  PRODUCT_NAME: 6,   // G열
  PROCESS_TYPE: 7,   // H열
  QUANTITY: 8,       // I열
  MFG_DATE: 9,       // J열
  MATERIAL_DATE: 10, // K열
  DELIVERY_DATE: 11, // L열
  URGENCY: 12,       // M열
  NIGHT_SHIFT: 13,   // N열
  REMARKS: 14,       // O열
} as const;
