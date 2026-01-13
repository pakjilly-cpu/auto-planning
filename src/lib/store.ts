import { create } from 'zustand';
import { persist } from 'zustand/middleware';
import { Vendor, ClientVendorMapping, ProductionItem, ProductionPlan } from './types';
import { defaultVendors, defaultClientMappings } from '@/data/defaults';

// 수기 고정된 배치 정보
export interface FixedPlacement {
  productCode: string;
  vendorName: string;
  lineNumber: number;
  dateKey: string; // yyyy-MM-dd
}

interface AppState {
  // 외주처 데이터
  vendors: Vendor[];
  setVendors: (vendors: Vendor[]) => void;
  updateVendor: (vendor: Vendor) => void;
  
  // 고객사-외주처 매칭
  clientMappings: ClientVendorMapping[];
  setClientMappings: (mappings: ClientVendorMapping[]) => void;
  addClientMapping: (mapping: ClientVendorMapping) => void;
  removeClientMapping: (clientCode: string) => void;
  
  // 생산 품목 (엑셀 업로드)
  productionItems: ProductionItem[];
  setProductionItems: (items: ProductionItem[]) => void;
  
  // 생산 계획 (배분 결과)
  productionPlans: ProductionPlan[];
  setProductionPlans: (plans: ProductionPlan[]) => void;
  
  // 선택된 월
  selectedMonth: Date;
  setSelectedMonth: (date: Date) => void;
  
  // UI 상태
  isLoading: boolean;
  setIsLoading: (loading: boolean) => void;
  
  // 수기 고정 배치
  fixedPlacements: FixedPlacement[];
  setFixedPlacement: (placement: FixedPlacement) => void;
  removeFixedPlacement: (productCode: string) => void;
  clearFixedPlacements: () => void;
  getFixedPlacement: (productCode: string) => FixedPlacement | undefined;
}

export const useAppStore = create<AppState>()(
  persist(
    (set) => ({
      // 외주처 초기값
      vendors: defaultVendors,
      setVendors: (vendors) => set({ vendors }),
      updateVendor: (vendor) =>
        set((state) => ({
          vendors: state.vendors.map((v) =>
            v.id === vendor.id ? vendor : v
          ),
        })),
      
      // 고객사-외주처 매칭 초기값
      clientMappings: defaultClientMappings,
      setClientMappings: (mappings) => set({ clientMappings: mappings }),
      addClientMapping: (mapping) =>
        set((state) => ({
          clientMappings: [
            ...state.clientMappings.filter(
              (m) => m.clientCode !== mapping.clientCode
            ),
            mapping,
          ],
        })),
      removeClientMapping: (clientCode) =>
        set((state) => ({
          clientMappings: state.clientMappings.filter(
            (m) => m.clientCode !== clientCode
          ),
        })),
      
      // 생산 품목
      productionItems: [],
      setProductionItems: (items) => set({ productionItems: items }),
      
      // 생산 계획
      productionPlans: [],
      setProductionPlans: (plans) => set({ productionPlans: plans }),
      
      // 선택된 월
      selectedMonth: new Date(),
      setSelectedMonth: (date) => set({ selectedMonth: date }),
      
      // UI 상태
      isLoading: false,
      setIsLoading: (loading) => set({ isLoading: loading }),
      
      // 수기 고정 배치
      fixedPlacements: [],
      setFixedPlacement: (placement) =>
        set((state) => ({
          fixedPlacements: [
            ...state.fixedPlacements.filter(p => p.productCode !== placement.productCode),
            placement,
          ],
        })),
      removeFixedPlacement: (productCode) =>
        set((state) => ({
          fixedPlacements: state.fixedPlacements.filter((p: FixedPlacement) => p.productCode !== productCode),
        })),
      clearFixedPlacements: () => set({ fixedPlacements: [] }),
      // getFixedPlacement는 selector로 사용: useAppStore(state => state.fixedPlacements.find(p => p.productCode === code))
      getFixedPlacement: () => undefined,
    }),
    {
      name: 'auto-planning-storage',
      partialize: (state) => ({
        vendors: state.vendors,
        clientMappings: state.clientMappings,
        fixedPlacements: state.fixedPlacements,
      }),
    }
  )
);
