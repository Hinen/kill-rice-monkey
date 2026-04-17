# NOL 구현 실행 계획

## 목표
- `tickets.interpark.com` NOL 예매 자동화 복구 및 완성
- 대상 버전: NEW(`26004589`) + OLD(`26003042`)
- 캡챠/구역 유무 직교 조합 전체 대응
- MELON 수준 기능 유지: 무한 대기열, 중복 좌석 회피(10회), Polly, 진행 보고, PauseGate, 세마포어, 페이지 스코어링
- 로그인 자동화는 구현하지 않음

## 필수 근거
- 실브라우저 remote-debug 탐색 결과
  - NEW `26004589`: `summaryData.isIngredientOnestop = true`, 예매 후 같은 탭 `/onestop/seat`
  - OLD `26003042`: `summaryData.isIngredientOnestop = false`, BuyButton 번들 기준 `waitingUrl === 'N'`이면 `UTIL.openPCOnestop(...)`, 아니면 `window.open(waitingUrl, 'waiting_<goodsCode>', ...)`
  - 두 상품 모두 `goodsStatus = 'Y'`, `salesTypeCode = '21001'`
- 이전 구현 히스토리 존재
  - 제거 커밋: `357d544 refactor: NOL 관련 코드 및 파일 전체 제거`
  - 복원 후보: `INolAutomationService.cs`, `NolAutomationService.cs`, `NolBookingPageClassifier.cs`, NOL 테스트 2개

## 구현 전략
1. 제거 직전(`357d544^`) NOL 코드와 테스트를 기준으로 복원 가능한 부분을 되살린다.
2. 현재 사용자 요구와 확정값에 맞지 않는 부분만 수정한다.
3. NEW 경로는 `/onestop/seat` + 캡챠 모달 + SVG 좌석 기준으로 정리한다.
4. OLD 경로는 popup/waiting/openOnestop 흐름을 유지하되, 팝업/alert/좌석완료 동작을 현재 코드 스타일에 맞춰 보정한다.

## 파일 작업 범위
- 추가
  - `src/KillRiceMonkey.Application/Abstractions/INolAutomationService.cs`
  - `src/KillRiceMonkey.Infrastructure/Services/NolAutomationService.cs`
- 수정
  - `src/KillRiceMonkey.Application/Models/TicketingTemplateType.cs`
  - `src/KillRiceMonkey.Infrastructure/DependencyInjection.cs`
  - `src/KillRiceMonkey.App/ViewModels/MainWindowViewModel.cs`
  - `src/KillRiceMonkey.App/MainWindow.xaml`
- 필요 시 복원/보정
  - `src/KillRiceMonkey.Infrastructure/Services/NolBookingPageClassifier.cs`
  - `tests/KillRiceMonkey.Tests/NolBookingPageClassifierTests.cs`
  - `tests/KillRiceMonkey.Tests/NolOnestopSeatCompleteClickTests.cs`

## 서비스 구현 체크리스트
- 상수
  - `NolRemoteDebugLaunchUrl = "https://tickets.interpark.com/"`
  - `NolRemoteDebugPort = 9225`
  - `NolCdpEndpoint = "http://127.0.0.1:9225/"`
  - `NolProfileName = "NolRemoteDebugProfile"`
  - `MaxSeatRetries = 10`
  - `MaxCaptchaAttempts = 5`
  - `NolCaptchaLength = 6`
- 공통 패턴
  - Melon/Yes24와 동일한 Polly, SemaphoreSlim, prepared browser/page 캐시
  - 진행 보고와 PauseGate 유지
  - 페이지 스코어링과 준비 상태 체크 유지
- 캡챠
  - `PlaywrightRuntime.RecognizeCaptchaTextAsync` 재사용
  - data URL → base64 → byte[] → `Cv2.ImDecode` 경로 유지
  - 대소문자 무시 입력
- NEW 경로
  - 메인 페이지에서 날짜/회차/예매 버튼 선택
  - 같은 탭 `/onestop/seat` 네비게이션 대기
  - `[class*=ModalCaptchaText_]` 대응
  - 구역 UI가 있으면 선행 선택
  - `section[class*=SeatPlan_svgWrap]` 내부 SVG 좌석 탐색
  - `button.EntButton_button.EntButton_primary.entButtonGlobal` 중 텍스트 기반으로 `입력완료`/`선택 완료` 구분 클릭
- OLD 경로
  - 클릭 후 새 팝업 또는 대기열 팝업 감지
  - alert/dialog 자동 수락 및 상태 플래그 유지
  - waiting/openOnestop 전환 감지
  - NetFunnel/대기열 무한 대기 유지
  - 구역/좌석/좌석완료는 제거 직전 구현 + 최근 fix 커밋(`401c7b2`, `d6e490d`) 기준 보정
- 중복 좌석 회피
  - MELON 방식 `HashSet` 제외 목록 + 10회 재시도 그대로 유지

## UI/VM 체크리스트
- Template 옵션에 `Nol` 추가
- `INolAutomationService` 주입
- NOL 전용 remote debug/prepare 커맨드 추가
- NOL 패널은 MELON 패널 구조 복제 기반
- NOL 선택 시 이미지 디렉터리 고정, 상태 메시지/준비 캐시/seat pause 지원 반영

## 검증 계획
1. 변경 파일 `lsp_diagnostics` 0 error
2. NOL 관련 테스트 복원 후 실행
3. `dotnet build KillRiceMonkey.sln`
4. `dotnet publish src/KillRiceMonkey.App/KillRiceMonkey.App.csproj -c Release`
5. 실브라우저 검증
   - 앱에서 `Nol` 옵션 노출 확인
   - remote debug 브라우저 실행 확인
   - NEW URL 기준 준비/실행 플로우 검증
   - 가능 범위에서 OLD URL popup/waiting 분기 검증
6. `.scratch/` 임시 산출물 정리
7. 작은 단위 한국어 커밋
