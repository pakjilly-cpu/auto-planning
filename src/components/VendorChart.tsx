'use client';

import { useMemo } from 'react';
import {
  BarChart,
  Bar,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  ResponsiveContainer,
  Cell,
  Legend,
  ReferenceLine,
} from 'recharts';
import { useAppStore } from '@/lib/store';
import { ProductionItem, Vendor } from '@/lib/types';

const colors = [
  '#3B82F6', // blue
  '#22C55E', // green
  '#A855F7', // purple
  '#F97316', // orange
  '#EC4899', // pink
  '#06B6D4', // cyan
  '#F59E0B', // amber
];

interface ChartDataItem {
  name: string;
  quantity: number;
  target: number;
  color: string;
}

export default function VendorChart() {
  const { productionItems, vendors } = useAppStore();

  const chartData = useMemo(() => {
    const vendorQuantities: Record<string, number> = {};
    
    productionItems.forEach((item: ProductionItem) => {
      const vendor = item.assignedVendor || '미배정';
      vendorQuantities[vendor] = (vendorQuantities[vendor] || 0) + item.quantity;
    });
    
    return vendors.map((vendor: Vendor, index: number) => ({
      name: vendor.name,
      quantity: vendorQuantities[vendor.name] || 0,
      target: vendor.monthlyTarget || 0,
      color: colors[index % colors.length],
    })).filter((v: ChartDataItem) => v.quantity > 0 || v.target > 0);
  }, [productionItems, vendors]);

  const totalQuantity = chartData.reduce((sum: number, d: ChartDataItem) => sum + d.quantity, 0);

  if (productionItems.length === 0) {
    return (
      <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
        <h3 className="text-lg font-semibold mb-4">외주처별 배분 현황</h3>
        <div className="h-64 flex items-center justify-center text-gray-500">
          데이터를 업로드하면 차트가 표시됩니다.
        </div>
      </div>
    );
  }

  return (
    <div className="bg-white rounded-xl shadow-sm border border-gray-100 p-6">
      <div className="flex items-center justify-between mb-4">
        <h3 className="text-lg font-semibold">외주처별 배분 현황</h3>
        <span className="text-sm text-gray-500">
          총 {totalQuantity.toLocaleString()}개
        </span>
      </div>
      
      <div className="h-72">
        <ResponsiveContainer width="100%" height="100%">
          <BarChart
            data={chartData}
            margin={{ top: 20, right: 30, left: 20, bottom: 5 }}
          >
            <CartesianGrid strokeDasharray="3 3" vertical={false} />
            <XAxis
              dataKey="name"
              tick={{ fontSize: 12 }}
              tickLine={false}
            />
            <YAxis
              tickFormatter={(value) => `${(value / 10000).toFixed(0)}만`}
              tick={{ fontSize: 12 }}
              tickLine={false}
              axisLine={false}
            />
            <Tooltip
              formatter={(value) => [
                `${(value as number)?.toLocaleString() || 0}개`,
              ]}
              labelFormatter={(label) => label}
              contentStyle={{
                borderRadius: '8px',
                border: '1px solid #E5E7EB',
              }}
            />
            <Legend
              formatter={(value) => value === 'quantity' ? '배정량' : '월간 목표'}
            />
            <Bar
              dataKey="quantity"
              name="quantity"
              radius={[4, 4, 0, 0]}
            >
              {chartData.map((entry: ChartDataItem, index: number) => (
                <Cell key={index} fill={entry.color} />
              ))}
            </Bar>
            <Bar
              dataKey="target"
              name="target"
              fill="#E5E7EB"
              radius={[4, 4, 0, 0]}
            />
          </BarChart>
        </ResponsiveContainer>
      </div>
      
      {/* 달성률 표시 */}
      <div className="mt-4 grid grid-cols-2 sm:grid-cols-4 gap-3">
        {chartData
          .filter((d: ChartDataItem) => d.target > 0)
          .map((vendor: ChartDataItem, index: number) => {
            const rate = Math.round((vendor.quantity / vendor.target) * 100);
            return (
              <div
                key={index}
                className="p-3 rounded-lg bg-gray-50"
              >
                <div className="flex items-center gap-2 mb-1">
                  <span
                    className="w-2 h-2 rounded-full"
                    style={{ backgroundColor: vendor.color }}
                  />
                  <span className="text-sm font-medium">{vendor.name}</span>
                </div>
                <div className="text-lg font-bold" style={{ color: vendor.color }}>
                  {rate}%
                </div>
                <div className="text-xs text-gray-500">
                  {vendor.quantity.toLocaleString()} / {vendor.target.toLocaleString()}
                </div>
              </div>
            );
          })}
      </div>
    </div>
  );
}
