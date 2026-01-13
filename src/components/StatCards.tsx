'use client';

import { useMemo } from 'react';
import { Package, Factory, TrendingUp, Target } from 'lucide-react';
import { useAppStore } from '@/lib/store';
import { calculateVendorStats } from '@/lib/allocation';

export default function StatCards() {
  const { productionItems, productionPlans, vendors } = useAppStore();

  const stats = useMemo(() => {
    const vendorStats = calculateVendorStats(productionPlans, vendors);
    const totalQuantity = productionItems.reduce((sum, item) => sum + item.quantity, 0);
    const totalPlanned = productionPlans.reduce((sum, plan) => sum + plan.totalQuantity, 0);
    
    return {
      totalItems: productionItems.length,
      totalQuantity,
      totalPlanned,
      vendorStats,
    };
  }, [productionItems, productionPlans, vendors]);

  const cards = [
    {
      title: '총 품목 수',
      value: stats.totalItems.toLocaleString(),
      suffix: '건',
      icon: Package,
      color: 'bg-blue-500',
    },
    {
      title: '총 생산량',
      value: stats.totalQuantity.toLocaleString(),
      suffix: '개',
      icon: Factory,
      color: 'bg-green-500',
    },
    {
      title: '배정 완료',
      value: stats.totalPlanned.toLocaleString(),
      suffix: '개',
      icon: TrendingUp,
      color: 'bg-purple-500',
    },
    {
      title: '외주처',
      value: vendors.length.toString(),
      suffix: '곳',
      icon: Target,
      color: 'bg-orange-500',
    },
  ];

  return (
    <div className="grid grid-cols-2 lg:grid-cols-4 gap-4">
      {cards.map((card, index) => (
        <div
          key={index}
          className="bg-white rounded-xl shadow-sm border border-gray-100 p-5"
        >
          <div className="flex items-center justify-between">
            <div>
              <p className="text-sm text-gray-500 font-medium">{card.title}</p>
              <p className="text-2xl font-bold mt-1">
                {card.value}
                <span className="text-sm font-normal text-gray-400 ml-1">
                  {card.suffix}
                </span>
              </p>
            </div>
            <div className={`${card.color} p-3 rounded-lg`}>
              <card.icon className="w-6 h-6 text-white" />
            </div>
          </div>
        </div>
      ))}
    </div>
  );
}
