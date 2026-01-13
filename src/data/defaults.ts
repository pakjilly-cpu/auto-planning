import { Vendor, ClientVendorMapping } from '@/lib/types';

// 외주처 기본 데이터
export const defaultVendors: Vendor[] = [
  {
    id: 'withmom',
    name: '위드맘',
    lineCount: 8,
    dailyCapacityPerLine: 15000,
    capabilities: ['normal', 'shrink', 'highFrequency'],
    monthlyTarget: 1600000,
    priority: 1,
  },
  {
    id: 'linear',
    name: '리니어',
    lineCount: 5,
    dailyCapacityPerLine: 13000,
    capabilities: ['normal', 'shrink', 'mixing', 'highFrequency'],
    monthlyTarget: 1600000,
    priority: 1,
  },
  {
    id: 'gram',
    name: '그램',
    lineCount: 3,
    dailyCapacityPerLine: 10000,
    capabilities: ['normal'],
    monthlyTarget: 500000,
    priority: 1,
  },
  {
    id: 'isis',
    name: '이시스',
    lineCount: 2,
    dailyCapacityPerLine: 7000,
    capabilities: ['normal', 'mixing'],
    monthlyTarget: 0,
    priority: 2,
  },
  {
    id: 'elluo',
    name: '엘루오',
    lineCount: 1,
    dailyCapacityPerLine: 7000,
    capabilities: ['normal'],
    monthlyTarget: 0,
    priority: 2,
  },
  {
    id: 'kcostech',
    name: '케이코스텍',
    lineCount: 2,
    dailyCapacityPerLine: 7000,
    capabilities: ['normal', 'mixing'],
    monthlyTarget: 0,
    priority: 2,
  },
  {
    id: 'dami',
    name: '다미',
    lineCount: 2,
    dailyCapacityPerLine: 7000,
    capabilities: ['normal', 'shrink'],
    monthlyTarget: 0,
    priority: 2,
  },
];

// 고객사-외주처 매칭 기본 데이터
export const defaultClientMappings: ClientVendorMapping[] = [
  // 그램 우선 배정
  { clientCode: 'CLO', vendorId: 'gram', priority: 1 },
  { clientCode: 'ERK', vendorId: 'gram', priority: 1 },
  
  // 위드맘 우선 배정
  { clientCode: 'DPD', vendorId: 'withmom', priority: 1 },
  { clientCode: 'GDI', vendorId: 'withmom', priority: 1 },
  
  // 리니어 우선 배정
  { clientCode: 'MDH', vendorId: 'linear', priority: 1 },
  { clientCode: 'APS', vendorId: 'linear', priority: 1 },
  
  // 케이코스텍 우선 배정
  { clientCode: 'PUR', vendorId: 'kcostech', priority: 1 },
];

// 외주처 ID -> 이름 매핑
export const vendorNameMap: Record<string, string> = {
  withmom: '위드맘',
  linear: '리니어',
  gram: '그램',
  isis: '이시스',
  elluo: '엘루오',
  kcostech: '케이코스텍',
  dami: '다미',
};

// 이름 -> 외주처 ID 매핑
export const vendorIdMap: Record<string, string> = {
  '위드맘': 'withmom',
  '리니어': 'linear',
  '그램': 'gram',
  '이시스': 'isis',
  '엘루오': 'elluo',
  '케이코스텍': 'kcostech',
  '다미': 'dami',
};
