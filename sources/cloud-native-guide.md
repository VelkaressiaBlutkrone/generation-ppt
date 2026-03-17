# 클라우드 네이티브 애플리케이션 구축 가이드

2026년 현대적 클라우드 네이티브 아키텍처의 핵심 원칙과 실전 구현 전략

---

<https://kubernetes.io>

# 클라우드 네이티브란?

## 정의

클라우드 네이티브는 클라우드 컴퓨팅의 이점을 최대한 활용하도록 설계된 소프트웨어 접근 방식이다.

> CNCF(Cloud Native Computing Foundation)는 클라우드 네이티브 기술을 "조직이 퍼블릭, 프라이빗, 하이브리드 클라우드에서 확장 가능한 애플리케이션을 구축하고 실행할 수 있게 해주는 기술"로 정의한다.

## 핵심 특성

- **컨테이너화**: 애플리케이션과 의존성을 컨테이너로 패키징
- **동적 오케스트레이션**: Kubernetes를 통한 자동 스케일링과 관리
- **마이크로서비스 지향**: 독립적으로 배포 가능한 서비스 단위
- **선언적 API**: 인프라를 코드로 관리 (IaC)

---

### 전통적 아키텍처와의 차이

| 항목        | 전통적 방식          | 클라우드 네이티브        |
| ----------- | -------------------- | ------------------------ |
| 배포 단위   | 모놀리식 WAR/EAR     | 컨테이너 이미지          |
| 확장 방식   | 수직 확장 (Scale-Up) | 수평 확장 (Scale-Out)    |
| 장애 대응   | 수동 복구            | 자동 복구 (Self-Healing) |
| 릴리스 주기 | 월/분기 단위         | 일/시간 단위             |
| 인프라 관리 | 수동 설정            | IaC (Terraform, Pulumi)  |

# 컨테이너와 Docker

## Docker 기본 개념

Docker는 애플리케이션을 컨테이너라는 격리된 환경에서 실행할 수 있게 해주는 플랫폼이다.

```dockerfile
FROM eclipse-temurin:21-jre-alpine
WORKDIR /app
COPY build/libs/app.jar app.jar
EXPOSE 8080
ENTRYPOINT ["java", "-jar", "app.jar"]
```

> 멀티스테이지 빌드를 사용하면 최종 이미지 크기를 70% 이상 줄일 수 있다.

### 컨테이너 vs VM

- 컨테이너는 OS 커널을 공유하므로 **수초 내 시작**
- VM은 전체 게스트 OS를 포함하므로 **수분 소요**
- 컨테이너 이미지는 보통 **50~200MB**, VM 이미지는 **수 GB**

https://www.docker.com

# Kubernetes 오케스트레이션

## 왜 Kubernetes인가?

컨테이너가 수백, 수천 개로 늘어나면 수동 관리가 불가능하다. Kubernetes는 이 문제를 자동화한다.

### 핵심 컴포넌트

1. **Pod**: 가장 작은 배포 단위 (1개 이상의 컨테이너)
2. **Service**: Pod에 대한 네트워크 추상화
3. **Deployment**: Pod의 선언적 업데이트 관리
4. **Ingress**: 외부 트래픽을 클러스터 내부로 라우팅
5. **ConfigMap / Secret**: 환경 설정과 민감정보 분리

---

## 기본 배포 매니페스트

```yaml
apiVersion: apps/v1
kind: Deployment
metadata:
  name: api-server
spec:
  replicas: 3
  selector:
    matchLabels:
      app: api-server
  template:
    spec:
      containers:
        - name: api
          image: registry.example.com/api:v2.1.0
          ports:
            - containerPort: 8080
          resources:
            requests:
              cpu: "250m"
              memory: "256Mi"
            limits:
              cpu: "500m"
              memory: "512Mi"
```

> 리소스 requests와 limits를 항상 설정하라. 미설정 시 노드 전체에 영향을 줄 수 있다.

https://kubernetes.io/docs/home/

# 마이크로서비스 아키텍처

## 서비스 분리 전략

마이크로서비스는 비즈니스 도메인 경계를 따라 분리한다.

### 도메인 주도 설계 (DDD) 기반 분리

- **Bounded Context**: 각 서비스의 명확한 경계 설정
- **Aggregate**: 데이터 일관성의 최소 단위
- **Domain Event**: 서비스 간 비동기 통신의 기반

## 서비스 간 통신 패턴

| 패턴            | 방식      | 장점                | 단점             |
| --------------- | --------- | ------------------- | ---------------- |
| REST API        | 동기 HTTP | 단순, 표준화        | 강한 결합        |
| gRPC            | 동기 RPC  | 고성능, 타입 안전   | 디버깅 어려움    |
| 메시지 큐       | 비동기    | 느슨한 결합         | 복잡성 증가      |
| 이벤트 스트리밍 | 비동기    | 실시간, 재처리 가능 | 순서 보장 어려움 |

> 동기 통신은 단순하지만 장애 전파 위험이 있다. 핵심 비즈니스 흐름에는 비동기 패턴을 우선 고려하라.

### Circuit Breaker 패턴

서비스 장애가 전파되는 것을 막기 위한 핵심 패턴이다:

1. **Closed**: 정상 상태, 요청을 그대로 전달
2. **Open**: 실패율 임계치 초과 시, 즉시 실패 응답 반환
3. **Half-Open**: 일정 시간 후 일부 요청을 시도하여 복구 확인

```java
@CircuitBreaker(name = "paymentService", fallbackMethod = "fallbackPayment")
public PaymentResponse processPayment(PaymentRequest request) {
    return paymentClient.charge(request);
}

public PaymentResponse fallbackPayment(PaymentRequest request, Throwable t) {
    return PaymentResponse.pending("결제 서비스 일시 불가, 재시도 예정");
}
```

# 관측 가능성 (Observability)

## 세 가지 축

관측 가능성은 시스템 내부 상태를 외부에서 이해할 수 있게 만드는 능력이다.

### 메트릭 (Metrics)

- 시스템의 **수치적 상태**를 시계열로 기록
- 도구: Prometheus, Grafana
- 예: 요청 수, 응답 시간, 에러율, CPU 사용량

### 로그 (Logs)

- 개별 이벤트의 **상세 기록**
- 도구: ELK Stack, Loki
- 구조화된 JSON 로그를 권장

### 트레이스 (Traces)

- 요청의 **전체 경로**를 추적
- 도구: Jaeger, Zipkin, OpenTelemetry
- 분산 시스템에서 병목 지점 파악에 필수

---

> "관측 가능성 없이 마이크로서비스를 운영하는 것은 눈을 감고 고속도로를 달리는 것과 같다."

## 핵심 지표 (Golden Signals)

| 지표       | 설명           | 알람 기준 예시          |
| ---------- | -------------- | ----------------------- |
| Latency    | 요청 처리 시간 | p99 > 500ms             |
| Traffic    | 초당 요청 수   | RPS 급변 (±50%)         |
| Errors     | 에러 응답 비율 | 5xx > 1%                |
| Saturation | 리소스 사용률  | CPU > 80%, Memory > 85% |

https://grafana.com

# CI/CD 파이프라인

## GitOps 워크플로우

GitOps는 Git 저장소를 단일 진실의 원천(Single Source of Truth)으로 사용하는 운영 모델이다.

### 파이프라인 단계

1. **코드 커밋** → GitHub/GitLab 저장소
2. **자동 빌드** → GitHub Actions / Jenkins
3. **테스트 실행** → 단위 / 통합 / E2E 테스트
4. **이미지 빌드** → Docker 이미지 생성 및 레지스트리 푸시
5. **배포 매니페스트 업데이트** → Kustomize / Helm
6. **자동 동기화** → ArgoCD가 클러스터에 반영

```yaml
# GitHub Actions 워크플로우 예시
name: CI/CD Pipeline
on:
  push:
    branches: [main]
jobs:
  build-and-deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - name: Build & Test
        run: ./gradlew build test
      - name: Build Docker Image
        run: docker build -t registry.example.com/app:${{ github.sha }} .
      - name: Push to Registry
        run: docker push registry.example.com/app:${{ github.sha }}
```

> 프로덕션 배포는 반드시 자동화하되, 승인 게이트를 추가하여 실수를 방지하라.

https://argo-cd.readthedocs.io

# 기술 스택 종합

## 추천 스택

| 영역           | 기술                        | 비고                             |
| -------------- | --------------------------- | -------------------------------- |
| 언어           | Java 21 / Go 1.22           | 팀 역량에 따라 선택              |
| 프레임워크     | Spring Boot 4.0 / Gin       | 생산성과 생태계 고려             |
| 컨테이너       | Docker + containerd         | 표준 OCI 런타임                  |
| 오케스트레이션 | Kubernetes 1.30             | 관리형 서비스 권장 (EKS/GKE/AKS) |
| 서비스 메시    | Istio / Linkerd             | 트래픽 관리, mTLS                |
| 관측           | Prometheus + Grafana + Loki | 통합 관측 스택                   |
| CI/CD          | GitHub Actions + ArgoCD     | GitOps 기반                      |
| IaC            | Terraform / Pulumi          | 인프라 버전 관리                 |

---

### 선택 기준

> 기술 선택에서 가장 중요한 것은 "최신"이 아니라 "팀이 운영할 수 있는가"이다. 운영 가능한 기술이 최고의 기술이다.

- 팀의 현재 역량과 학습 곡선
- 커뮤니티 규모와 문서화 수준
- 벤더 종속(Lock-in) 정도
- 장기 유지보수 비용

# 감사합니다

클라우드 네이티브 애플리케이션 구축 가이드

질문이 있으시면 언제든 연락 주세요
