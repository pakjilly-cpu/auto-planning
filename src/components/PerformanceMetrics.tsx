'use client';

import { TrendingUp, Zap, BarChart3, Clock, Target, CheckCircle } from 'lucide-react';
import { SchedulingPerformanceMetrics } from '@/lib/allocation';

interface PerformanceMetricsProps {
  metrics: SchedulingPerformanceMetrics | null;
  isLoading?: boolean;
}

export default function PerformanceMetrics({ metrics, isLoading = false }: PerformanceMetricsProps) {
  if (isLoading) {
    return (
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
        {[...Array(6)].map((_, i) => (
          <div key={i} className="bg-white rounded-lg p-4 border border-gray-200 animate-pulse">
            <div className="h-4 bg-gray-200 rounded w-3/4 mb-4"></div>
            <div className="h-8 bg-gray-200 rounded w-1/2"></div>
          </div>
        ))}
      </div>
    );
  }

  if (!metrics) {
    return (
      <div className="text-center text-gray-500 py-8">
        데이터를 불러와 성능 지표를 확인하세요.
      </div>
    );
  }

  const metricCards = [
    {
      title: 'Makespan (완료시간)',
      value: `${metrics.makespan}일`,
      description: '전체 생산 완료까지 소요 일수',
      icon: Clock,
      color: 'bg-blue-50 text-blue-600',
      barColor: 'bg-blue-600',
    },
    {
      title: '평균 자원 활용률',
      value: `${metrics.averageUtilization}%`,
      description: '외주처 가용 용량 대비 실제 사용률',
      icon: Zap,
      color: 'bg-green-50 text-green-600',
      barColor: 'bg-green-600',
      progress: metrics.averageUtilization,
    },
    {
      title: '부하 균형도',
      value: `${(metrics.loadBalance * 100).toFixed(1)}%`,
      description: '0에 가까울수록 좋음 (완벽한 균형)',
      icon: BarChart3,
      color: 'bg-purple-50 text-purple-600',
      barColor: 'bg-purple-600',
      progress: Math.max(0, 100 - metrics.loadBalance * 100),
      isInverse: true,
    },
    {
      title: '납기 준수율',
      value: `${metrics.deliveryOnTimeRate}%`,
      description: '설정된 납기일 내 완료 비율',
      icon: CheckCircle,
      color: 'bg-emerald-50 text-emerald-600',
      barColor: 'bg-emerald-600',
      progress: metrics.deliveryOnTimeRate,
    },
    {
      title: '월간 목표 달성률',
      value: `${metrics.monthlyTargetAchievementRate}%`,
      description: '외주처 월간 목표 생산량 달성도',
      icon: Target,
      color: 'bg-orange-50 text-orange-600',
      barColor: 'bg-orange-600',
      progress: metrics.monthlyTargetAchievementRate,
    },
    {
      title: '특수공정 준수율',
      value: `${metrics.specialProcessComplianceRate}%`,
      description: '특수공정 능력에 맞는 배치 비율',
      icon: TrendingUp,
      color: 'bg-red-50 text-red-600',
      barColor: 'bg-red-600',
      progress: metrics.specialProcessComplianceRate,
    },
  ];

  return (
    <div className="space-y-6">
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
        {metricCards.map((card, index) => {
          const Icon = card.icon;
          const hasProgress = card.progress !== undefined;

          return (
            <div
              key={index}
              className="bg-white rounded-lg border border-gray-200 p-4 hover:shadow-md transition-shadow"
            >
              <div className="flex items-start justify-between mb-3">
                <div>
                  <p className="text-sm font-medium text-gray-600">{card.title}</p>
                  <p className="text-2xl font-bold text-gray-900 mt-1">{card.value}</p>
                </div>
                <div className={`p-2 rounded-lg ${card.color}`}>
                  <Icon className="w-5 h-5" />
                </div>
              </div>

              {hasProgress && (
                <div className="mb-2">
                  <div className="w-full bg-gray-200 rounded-full h-2">
                    <div
                      className={`h-2 rounded-full transition-all ${card.barColor}`}
                      style={{
                        width: `${Math.min(100, Math.max(0, card.progress || 0))}%`,
                      }}
                    />
                  </div>
                </div>
              )}

              <p className="text-xs text-gray-500">{card.description}</p>
            </div>
          );
        })}
      </div>

      {/* 권장사항 */}
      <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
        <h3 className="font-semibold text-blue-900 mb-2">💡 최적화 권장사항</h3>
        <ul className="text-sm text-blue-800 space-y-1">
          {metrics.loadBalance > 0.3 && (
            <li>• 외주처 간 부하 불균형이 큽니다. 부하 균형화 최적화를 시도해보세요.</li>
          )}
          {metrics.averageUtilization < 70 && (
            <li>• 자원 활용률이 낮습니다. 더 효율적인 배치를 고려해보세요.</li>
          )}
          {metrics.monthlyTargetAchievementRate < 80 && (
            <li>• 월간 목표 달성률이 낮습니다. 우선순위 외주처에 더 많은 품목을 배정하세요.</li>
          )}
          {metrics.deliveryOnTimeRate < 95 && (
            <li>• 일부 품목의 납기 준수가 어려울 수 있습니다. 납기 스케줄을 검토하세요.</li>
          )}
        </ul>
      </div>
    </div>
  );
}
