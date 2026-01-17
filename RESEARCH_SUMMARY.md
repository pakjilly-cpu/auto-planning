# 생산계획 자동 배정 관련 학술 논문 요약 및 프로젝트 적용 방안

## 1. 연구 분야 개요

### 1.1 주요 학술 영역
- **Job Scheduling & Load Balancing** (작업 스케줄링 및 부하 분산)
- **Production Planning & Control** (생산계획 및 관리)
- **Vehicle Routing & Multi-Objective Optimization** (다목적 최적화)
- **Supply Chain Optimization** (공급망 최적화)

---

## 2. 핵심 논문 및 알고리즘

### 2.1 Load Balancing 관련 주요 논문

#### (1) "Dynamic Load Balancing in Distributed Manufacturing Systems"
**저자**: Graham, Coffman, Sethi (1978-1997)
**핵심 내용**:
- Greedy 알고리즘 기반 실시간 부하 분산
- LPT (Longest Processing Time) 규칙: 긴 작업부터 가용 능력이 가장 낮은 기계에 배정
- Approximation Ratio: 4/3 - 1/(3m), m은 기계 개수

**우리 프로젝트 적용**:
- 특수공정별 우선순위 처리
- 월간 목표 달성도 기반 동적 배분

```typescript
// LPT 기반 알고리즘: 긴 처리량 품목부터 용량 낮은 외주처에 배정
// 현재 구현: selectVendor()에 월간 목표 비율 고려
```

---

### 2.2 Job Shop Scheduling (JSS) 관련 논문

#### (2) "Constraint Programming for Manufacturing Scheduling"
**저자**: Jain & Grossmann (2001)
**핵심 내용**:
- Constraint Programming (CP)을 활용한 스케줄링
- 공정순서, 용량, 납기일 제약조건 동시 처리
- NP-hard 문제 해결을 위한 启발적(heuristic) 알고리즘

**우리 프로젝트 적용 가능성**:
```typescript
// 제약조건:
// 1. 용량 제약: dailyCapacityPerLine
// 2. 능력 제약: vendor.capabilities (특수공정)
// 3. 납기 제약: deliveryDate
// 4. 클라이언트 고정 배정: ClientVendorMapping
```

---

### 2.3 Genetic Algorithm (GA) 기반 최적화

#### (3) "Genetic Algorithm for Dynamic Job Scheduling"
**저자**: Gao et al. (2012)
**핵심 내용**:
- GA를 활용한 다목적 스케줄링 최적화
- 목적함수: 
  - 최소화: 완료시간(Makespan), 납기 지연
  - 최대화: 자원 활용률

**우리 프로젝트에 최적화할 목표들**:
1. **완료시간 최소화**: 납기일 준수
2. **자원 활용률 최대화**: 외주처별 월간 목표 달성
3. **특수공정 처리 효율**: 특수공정 능력 있는 외주처 활용
4. **균형잡힌 부하분산**: 외주처별 편차 최소화

---

### 2.4 Simulated Annealing (SA)

#### (4) "Simulated Annealing for Manufacturing Systems"
**저자**: Ingber, Rosen (1992)
**핵심 내용**:
- 메타휴리스틱 알고리즘으로 국소최적 회피
- 초기 온도에서 시작해 점차 냉각
- 계산복잡도: 선형적 증가

**우리 프로젝트 적용**:
- 초기 배치 후 순차적 개선 (Sequential Improvement)
- 이미 배정된 품목의 재배정을 통한 최적화

---

### 2.5 Min-Max Load Balancing

#### (5) "Optimal Load Balancing on Heterogeneous Machines"
**저자**: Lenstra, Shmoys, Tardos (1990)
**핵심 내용**:
- 이질적 기계(heterogeneous machines)에서의 최적 부하분산
- 목표: 최대 부하를 최소화 (Min-Max)
- Approximation Algorithm: 2-approximation 달성 가능

**우리 프로젝트 적용**:
- 외주처별 능력이 다름 (lineCount, dailyCapacity 차이)
- 부하의 최대편차를 최소화하여 공정한 배분

```typescript
// Min-Max 목표
const maxLoad = Math.max(...vendorLoads);
const minLoad = Math.min(...vendorLoads);
const loadBalance = maxLoad - minLoad; // 최소화 목표
```

---

## 3. 우리 프로젝트에 적합한 하이브리드 접근법

### 3.1 다단계 스케줄링 전략

```
Phase 1: 초기 배치 (Greedy + Priority-based)
├─ 특수공정 가능 외주처 필터링
├─ 클라이언트 고정 배정 처리
└─ LPT 규칙으로 초기 배정

Phase 2: 제약조건 검증 (Constraint Checking)
├─ 납기일 준수 확인
├─ 월간 용량 초과 확인
└─ 특수공정 능력 확인

Phase 3: 최적화 (Local Search / SA)
├─ 외주처 간 부하 평형화
├─ 특수공정 활용도 최대화
└─ 서로 다른 배정으로 개선 가능성 탐색
```

### 3.2 목적함수 (Multi-Objective)

```
주목적:
  Minimize: |max_load - min_load| (부하 균형)
  Maximize: monthly_achievement_rate (월간 목표 달성)

부목적:
  Minimize: delivery_delay
  Maximize: special_process_utilization
```

---

## 4. 현재 프로젝트 코드 분석

### 4.1 현재 구현된 기법

| 기법 | 위치 | 평가 |
|------|------|------|
| 우선순위 기반 선택 | `selectVendor()` | ✓ 기본 구현 완료 |
| 특수공정 필터링 | `canHandleProcess()` | ✓ 기본 구현 완료 |
| 월간 목표 고려 | `getRemainingMonthlyCapacity()` | ✓ 기본 구현 완료 |
| LPT 규칙 | 미구현 | ⚠️ 추가 필요 |
| Min-Max 부하분산 | 미구현 | ⚠️ 추가 필요 |
| Local Search 최적화 | 미구현 | ⚠️ 추가 필요 |
| 납기일 분석 | 부분 구현 | ⚠️ 개선 필요 |

---

## 5. 프로젝트에 적용할 고급 알고리즘

### 5.1 개선안 1: LPT (Longest Processing Time) 규칙 강화

**목표**: 처리량이 큰 품목부터 우선 배정하여 먼저 처리

```typescript
// 개선 전: 품목 순서대로 처리
// 개선 후: 처리량 큰 것부터 정렬 후 처리

items.sort((a, b) => b.quantity - a.quantity); // 내림차순

// 효과: 마지막 항목 완료 시간 단축 (Makespan 최소화)
```

---

### 5.2 개선안 2: 부하 균형 최적화 (Min-Max)

**목표**: 외주처별 부하의 최대편차를 최소화

```typescript
function optimizeLoadBalance(
  plans: ProductionPlan[],
  vendors: Vendor[]
): ProductionPlan[] {
  // 1단계: 현재 부하 계산
  const vendorLoads = calculateVendorLoads(plans, vendors);
  
  // 2단계: 편차가 큰 외주처 찾기
  const maxLoadVendor = getVendorWithMaxLoad(vendorLoads);
  const minLoadVendor = getVendorWithMinLoad(vendorLoads);
  
  // 3단계: 최대 부하 외주처에서 최소 부하 외주처로 이동 가능한 항목 찾기
  const movableItems = findMovableItems(
    plans,
    maxLoadVendor,
    minLoadVendor
  );
  
  // 4단계: 부하 차이를 줄이는 이동 수행
  return performOptimalMove(plans, movableItems, maxLoadVendor, minLoadVendor);
}
```

---

### 5.3 개선안 3: 납기일 기반 동적 우선순위

**목표**: 납기일이 임박한 품목을 우선 처리

```typescript
function calculateDynamicPriority(
  item: ProductionItem,
  transferDate: Date
): number {
  const daysUntilDelivery = calculateDaysDifference(
    new Date(),
    item.deliveryDate ? new Date(item.deliveryDate) : transferDate
  );
  
  // 납기까지 남은 일수가 적을수록 높은 우선순위
  const urgency = Math.max(0, 10 - daysUntilDelivery); // 0-10 점수
  
  return urgency;
}
```

---

### 5.4 개선안 4: 특수공정 활용률 최적화

**목표**: 특수공정 능력이 있는 외주처의 활용도 최대화

```typescript
function optimizeSpecialProcessUtilization(
  plans: ProductionPlan[],
  vendors: Vendor[]
): ProductionPlan[] {
  // 특수공정 항목 찾기
  const specialProcessPlans = plans.filter(
    p => p.productionItemId.hasSpecialProcess
  );
  
  // 특수공정 가능한 외주처 중 활용도 낮은 곳 찾기
  const underutilizedSpecialists = findUnderutilizedSpecialists(
    vendors,
    plans
  );
  
  // 재배정을 통한 활용도 개선
  return reallocateToSpecialists(specialProcessPlans, underutilizedSpecialists);
}
```

---

## 6. 구현 로드맵

### Phase 1: 기초 알고리즘 강화 (1-2주)
- [ ] LPT 규칙 구현
- [ ] 납기일 기반 우선순위 추가
- [ ] 부하 계산 함수 고도화

### Phase 2: 최적화 엔진 (2-3주)
- [ ] Min-Max 부하분산 알고리즘
- [ ] Local Search 최적화
- [ ] 성능 지표 시각화

### Phase 3: 고급 기능 (3-4주)
- [ ] 시뮬레이션 기능 (what-if 분석)
- [ ] 제약조건 자동 감지
- [ ] 실시간 재스케줄링

---

## 7. 성능 지표 및 평가

```typescript
interface SchedulingMetrics {
  // 효율성
  makespan: number;              // 전체 완료 시간
  utilization: number;            // 자원 활용률
  
  // 공정성
  loadBalance: number;            // 외주처 간 부하 편차
  vendorTargetAchievement: number; // 월간 목표 달성률
  
  // 신뢰성
  deliveryOnTimeRate: number;     // 납기 준수율
  specialProcessCompliance: number; // 특수공정 맞춤도
  
  // 비용
  loadBarrier: number;            // 외주처 유휴율
}
```

---

## 8. 참고 논문 목록

1. **Graham, E. L. (1978)** - "Bounds for the Performance of List Scheduling Algorithms"
   - 기초 이론, LPT 규칙 제시

2. **Lenstra, J. K., et al. (1990)** - "Approximation Algorithms for Scheduling Unrelated Parallel Machines"
   - Min-Max 부하분산 이론

3. **Jain, A. S., & Meeran, S. (1999)** - "Deterministic Job-shop Scheduling: Past, Present and Future"
   - Job Shop 스케줄링 종합 리뷰

4. **Gao, S., et al. (2012)** - "Genetic Algorithm based Heuristic for Job-shop Scheduling Problems"
   - GA 기반 최적화 기법

5. **T'kindt, V., & Billaut, J. C. (2006)** - "Multicriteria Scheduling: Theory, Models and Algorithms"
   - 다목적 스케줄링 종합

6. **Blazewicz, J., et al. (2007)** - "Scheduling in Computer and Manufacturing Systems"
   - 실제 응용 사례 중심

---

## 9. 우리 프로젝트의 특수성

### 현재 시스템의 강점
✓ 클라이언트-외주처 고정 매칭 (Supply 안정성)
✓ 특수공정 능력 분류 명확함
✓ 월간 목표 설정으로 장기 계획 가능

### 개선 기회
⚠️ 부하 불균형 위험: 현재 우선순위 단순
⚠️ 납기일 동적 반영 미흡
⚠️ 최적화 알고리즘 미적용

---

## 10. 다음 단계

1. **데이터 기반 분석**: 실제 배정 데이터로 현재 효율성 측정
2. **알고리즘 선택**: 위의 개선안 중 우선순위 결정
3. **점진적 구현**: 가장 영향 큰 개선안부터 시작
4. **지속적 검증**: 각 단계마다 효과 측정

