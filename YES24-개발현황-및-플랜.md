# YES24 티켓 자동화 개발 현황 및 플랜

## 1. 조사 개요

- **초기 조사일**: 2026-04-08
- **버그 재조사일**: 2026-04-09 (팝업 초기화 race condition)
- **조사 방식**: Chrome remote-debugging(CDP) + websockets로 두 URL 모두 직접 접속 → 로그인 → 시간 선택 → 예매 버튼 → 좌석 iframe까지 전 과정 DOM 구조 확인
- **무구역 사이트**: https://ticket.yes24.com/Perf/57553 (Sou LIVE TOUR 2026 「Finder」 in Seoul)
- **구역 사이트**: https://ticket.yes24.com/Perf/55989?Gcode=009_300 (tuki. 1ST ASIA TOUR 2026 IN SEOUL)
- **버그 재조사 대상**: https://ticket.yes24.com/Perf/57918?Gcode=009_210_001 (포레스텔라 정규 4집 투어 콘서트 [THE LEGACY: SYMPHONY] In Seoul, IdTime=1429899)
- **로그인**: `yes24.txt` 자격증명으로 실제 CDP 로그인 성공 확인
- **참고**: 기존 `MelonAutomationService.cs` 패턴을 그대로 따름. YES24도 동일하게 jQuery 기반이며 Melon보다 오히려 플로우가 단순(CAPTCHA·대기열 없음).

---

## 2. YES24 전체 플로우 (직접 확인)

### 2.1 공연 페이지 (`ticket.yes24.com/Perf/{IdPerf}`)

- **프레임워크**: jQuery 3.4.1 + 레거시 ASP.NET. React 아님.
- **로그인 상태**: `window.IsLogin === '1'` (hidden `#HidIsLogin`에도 동일).
- **달력**: jQuery UI datepicker (`#rncalendar`, `.ui-datepicker-*`), `data-year`/`data-month`(0-base) 속성.
- **오늘 선택된 날짜**: `#rncalendar td.ui-datepicker-current-day` 의 child `a` (`ui-state-active`).
- **회차 목록**: `.rn-04-left-calist a[idTime="{TIME_ID}"]`.
  - 텍스트 예시: `"1회 오후 5시 30분"`, `onclick="jsf_pdi_ChangePlayTime('calendar','1425353','오후5시30분');"`
  - 선택되면 부모 `.rn-04-left-calist` 내 `a.on` 이 됨.
- **예매 버튼 2종**:
  - 상단 `a.rntop-btn-reserve` (`onclick="jsf_pdi_GoPerfSale(77);"`) → **실제로 렌더링에서 width/height 0 으로 숨겨짐**.
  - 하단 `a.rn-bb03` (`onclick="jsf_pdi_GoPerfSale();"`) → **이게 실제 가시/클릭 가능한 버튼**.
- **예매 버튼 내부 로직** (`jsf_pdi_GoPerfSale`):
  - `IsLogin != '1'` 이면 `jsf_pdi_LoginMsg(...)` 호출 후 return.
  - `valopt == 77` → `.rntop-select-time a.on` 의 `idTime`
  - 그 외 → `.rn-04-left-calist a.on` 의 `idTime`
  - `jsf_base_GoPerfSaleTime(IdPerf, IdTime, IsLogin, valopt)` 호출
    → `jsf_base_ShowPerfSaleProcess(IdPerf, IdTime, HcardOpt)`
    → **`window.open('/Pages/Perf/Sale/PerfSaleProcess.aspx?IdPerf={IdPerf}&IdTime={IdTime}', ...)`** (jgBookTPopup === jcMODE_WINDOW === 2)
- **최고 속도 경로**: `jsf_base_ShowPerfSaleProcess(idPerf, idTime)` 를 직접 호출하면 날짜/시간 UI 상태 검증을 전부 스킵하고 바로 팝업을 연다 — 단, `window.open`은 trusted user gesture 가 없으면 Chrome popup blocker 에 차단되므로 **CDP `Input.dispatchMouseEvent` 기반 trusted click** 또는 Playwright 의 `Locator.ClickAsync` 필요.

### 2.2 예매 프로세스 팝업 (`Pages/Perf/Sale/PerfSaleProcess.aspx`)

새 창(별도 `window`) 이며 jQuery 기반. Melon 의 `onestop.htm` 팝업과 역할 동일.

- **URL 템플릿**: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleProcess.aspx?IdPerf={IdPerf}&IdTime={IdTime}`
- **페이지 진입 시 상태 (완전 로드 후)**:
  - `jgCalSelDate = "YYYY-MM-DD"` (IdTime 으로부터 자동 결정됨)
  - `#IdTime.val() === {IdTime}` / `#ulTime > li.on` 이 선택되어 있음
  - `#step01`, `#step01_date`, `#step01_time`, `#step01_notice` 가 `display: block`
  - `#step03`, `#step04`, `#step05` 는 `display: none`
- **⚠️ 초기화 Race Condition (중요)**: 2026-04-09 재조사에서 확인한 사실.
  - `window.open` 직후 팝업의 `Page` 객체는 Playwright에 즉시 할당되지만, HTML 파싱/인라인 JS 실행은 아직 진행 중이다.
  - CDP 폴링 타임라인 (trigger click=0ms 기준, 57918 기준 실측값):
    - 0~230ms: 구 페이지(about:blank) 컨텍스트, `readyState=complete`, `fdc_FlashSeatLoad` 미정의
    - ~230~510ms: `readyState=loading`, 문서 본문 파싱 중, `fdc_FlashSeatLoad` 여전히 미정의
    - **~520ms**: `readyState=interactive`, `fdc_FlashSeatLoad` 정의됨, `#IdTime` 엘리먼트는 존재하지만 `value=""`, `#ulTime > li` 0개
    - **~520~617ms (약 100ms 위험 윈도우)**: `fdc_FlashSeatLoad` 는 호출 가능하지만 내부 가드에서 fail → alert
    - **~618ms**: `#IdTime.value` 채워지고 `#ulTime > li.on` 1개 → 이제 안전하게 호출 가능
    - ~1050ms: `readyState=complete`
  - `fdc_FlashSeatLoad` 내부 가드(실측):
    ```js
    if ($j("#IdTime").val() == "" || $j("#IdTime").val() == "0" || $j("#ulTime > li.on").length == 0) {
        jcSTEP_SEAT_CHK = true;
        fbk_Alert("공연회차를 선택하세요.");
    } else {
        // fdc_CtrlStep(jcSTEP_SEAT) → seat iframe 생성
    }
    ```
  - 이 100ms 위험 윈도우 안에서 호출되면 `fbk_Alert("공연회차를 선택하세요.")` 가 발동한다.
- **⚠️ fbk_Alert 모드 주의**: `jgBookAlert == jcMODE_JQUERY` (=1, 팝업 실측 기본값) 이므로 `fbk_Alert` 은 **native `alert()` 가 아닌 `$j("#dialogAlert").jAlert(...)` HTML 다이얼로그** 로 렌더된다.
  - 결과: Playwright 의 `Page.Dialog` 이벤트는 **native 다이얼로그에만 반응**하므로 jQuery 다이얼로그는 감지되지 않는다.
  - 대응: 팝업 진입 직후 JS 측에서 `fbk_Alert`, `window.alert` 를 훅해 `window.__yes24AlertDetected` 플래그 세팅.
- **StepCtrlBtn01 (다음단계 버튼)**: `a` onclick=`fdc_VerifySelSeatNumber()`
  - 현재 좌석 class 가 없으면 → `fdc_FlashSeatLoad()` 호출
  - `fdc_FlashSeatLoad()` → `fdc_CtrlStep(jcSTEP_SEAT)` → `#SeatFlashArea` 내부에 **좌석 iframe 생성**
- **좌석 iframe 생성 URL** (예시):
  - 무구역: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleHtmlSeat.aspx?idTime=1425353&idHall=14062&block=0&stMax=10&pHCardAppOpt=0`
  - 구역: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleHtmlSeat.aspx?idTime=1397006&idHall=13754&block=2&stMax=10&pHCardAppOpt=0`
  - 57918: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleHtmlSeat.aspx?idTime=1429899&idHall=14129&block=101&stMax=10&pHCardAppOpt=0` (구역 사이트)
- **자동화 관점 최적화 규칙**:
  1. **팝업 초기화 대기 필수**: `fdc_FlashSeatLoad` 호출 전에 `typeof fdc_FlashSeatLoad === 'function' && #IdTime.value && #ulTime > li.on > 0` 조건을 만족할 때까지 30ms 폴링 대기.
  2. **alert 훅 설치**: 초기화 대기 전에 `window.alert` + `fbk_Alert` 를 훅해 이후 호출 중 alert 가 발동되면 `__yes24AlertDetected` 플래그로 감지.
  3. 동일 origin 이므로 좌석 iframe 접근은 `iframe.contentDocument` 또는 Playwright `Frame` API 로 직접 가능.
  4. `jsf_base_ShowPerfSaleProcess` 직접 호출은 Chrome popup blocker 로 인해 **trusted user gesture 없이는 동작하지 않는다** (CDP `Input.dispatchMouseEvent` 또는 Playwright `ClickAsync` 필요). 현재 `Yes24AutomationService.TriggerYes24BookingPopupAsync` 의 직접 호출 경로는 사실상 언제나 350ms 대기 후 폴백 버튼 클릭으로 떨어진다.

### 2.3 좌석 iframe (`PerfSaleHtmlSeat.aspx`)

- **공통 DOM**:
  - `#divContainer > #divSeatArray` 안에 좌석 `div` 들이 절대 위치로 배치됨.
  - 각 좌석: `<div class="s{N}" id="t{SeatId}" name="tk" value="{SeatId}" style="LEFT:{x}px; TOP:{y}px"></div>`
  - `s0`~`s55+` 는 **스프라이트 이미지 인덱스**(`ic_seat{N}.gif`), 가용 여부와 직결되지 않음 (사이트/블록마다 가용 클래스가 다름).
- **가용 좌석 판별 (보편 규칙)**:
  - `typeof el.onclick === 'function'` ✔
  - 또는 `el.hasAttribute('grade')` ✔
  - 또는 `el.title` 이 비어있지 않음 (예: `"[E1 GATE] 202구역H열 001번"`)
- **TicketMapping 로직** (내부 구현 참고):
  ```js
  if (ticket[2] == "0") {  // status == 0 → 가용
      objtk.style.cursor = "hand";
      objtk.onclick = ChoiceSeat;
      $("#t"+ticket[0]).attr("grade", ticket[5]);
  }
  ```
- **좌석 클릭**: `ChoiceSeat` 가 `this`(jQuery) 문맥에서 동작. 직접 호출 시 `$(this)` 가 window 가 되어 오작동 → **반드시 요소 기준 `.click()` 호출** 또는 CDP `Input.dispatchMouseEvent` trusted click.
- **선택 확인**: 선택된 좌석은 `class="son"` (원본 클래스는 `oldclass` 속성에 백업) 이 되고 jQuery 전역 핸들러가 `#liSelSeat` 에 좌석 항목을 추가.
- **주요 전역 함수** (iframe 내부):
  - `GetSeatData(blk, seq)` → AJAX: `/OSIF/Book.asmx/GetBookWhole`
  - `GetSeatData_CallBack(msg, seq)` → XML 파싱 → `TicketMapping(blockSeat)`
  - `ChoiceSeat()` → 좌석 클릭 핸들러
  - `ChangeBlock(blk)` → 구역 변경 → `GetSeatData(blk, 1)`
  - `CreateLegend`, `CreateHallBgImg`, `GetImageMap` 등
- **STCLAB 봇매니저 CAPTCHA (향후 대비)**: `GetSeatData_CallBack` 에 특수 분기가 존재.
  - XML 응답의 `IdTime` 이 `"block"` 이면 → `parent.location.href = 'https://cdn-botmanager.stclab.com/yes24/{blockid}/index.html?...'`
  - 파라미터: `blockType`, `captchaTryLimit`, `requiredCaptchaSuccessCount`, `sessionId`, `redirectUrl`
  - **현재 대상 공연엔 미발생**. 나중에 발생 시 `parent.location.href` 변경을 감지하여 일시정지 + 사용자 개입 가능한 구조를 남겨둘 것.
- **구역 선택 유무 분기**:
  - **무구역**: iframe 로드 직후 `GetSeatData(0, 0)` 가 자동 실행 → `divSeatArray` 즉시 좌석 채움. `img[usemap]` 없음.
  - **구역**: 초기 `block>0` 이지만 실제 `divSeatArray` 는 비어 있고, `<img usemap="#map_ticket" src="stkfile.yes24.com/upload2/hallimg/.../*__MAP3.jpg">` + `<area href="javascript:ChangeBlock(N)">` 다수. 사용자가 구역을 골라야 좌석이 로드됨.

### 2.4 구역 선택 상세 (두 번째 사이트 - 55989)

- **홀 이미지**: `https://stkfile.yes24.com/upload2/hallimg/13000/13754/13754__MAP3.jpg` (629×574).
- **map `#map_ticket`** 안에 `area` 다수:
  - `<area shape="rect|poly" coords="..." href="javascript:ChangeBlock(101)" onmouseover="showMapInfo(101)" onmouseout="hidMapInfo(101)">`
  - 샘플 블록: 101, 102, 103, 104, 2, 3, 4, 5, 6, 7, 8, 9, 23 ...
- **구역 변경**: `ChangeBlock(blk)` → `GetSeatData(blk, 1)` → XML → `divSeatArray` 재구성.
- **구역 변경 후 가용 좌석 재확인**: `typeof el.onclick === 'function'` 으로 재스캔.
- **검증 완료된 샘플**(`ChangeBlock(2)` 후):
  - `byClass = { s13: 458(불가), s6: 26(가능) }`
  - `withOnclick: 26`
  - 샘플 `t7300228` class `s6` grade `"지정석 VIP석"` title `"[E1 GATE] 202구역H열 001번"`

### 2.5 좌석 선택 후 단계 이동

- 좌석 선택 확인 후 `fdc_VerifySelSeatNumber()` 또는 `#StepCtrlBtn01 a` 클릭 → 검증 → `fdc_CtrlStep(jcSTEP3)` → `#step03` 등급/프로모션 화면.
- **본 과제 범위는 "좌석 선택 완료"까지** 이므로, 자동화는 "다음단계" 클릭 직후 `#step01` 이 `display:none` 으로 바뀌고 `#step03` 이 `display:block` 이 되는 순간까지만 수행한다(결제 단계 미진행).

---

## 3. 로그인 전략 (Melon UX 동일)

Melon / NOL 과 동일하게 **사용자 수동 로그인 + 프로필 영속화**:

1. `LaunchRemoteDebugBrowserAsync()` 가 Chrome 을 `--remote-debugging-port=9224 --user-data-dir=%LOCALAPPDATA%\KillRiceMonkey\Yes24RemoteDebugProfile` 로 실행.
2. 사용자가 yes24 로그인(최초 1회). 세션 쿠키는 프로필 디렉토리에 저장되므로 이후 자동 유지.
3. `PrepareAutomationAsync()` 는 공연 페이지(`ticket.yes24.com/Perf/{IdPerf}*`)를 찾고, `IsLogin === '1'` 을 검증 후 캐시.
4. 검증 실패 시 명확한 한국어 에러로 사용자에게 재로그인 요구.

> **비고**: CDP JS eval 만으로 `document.LoginSub.submit()` 기반 자동 로그인도 조사 중 성공했지만, 해시/토큰 기반 무결성 체크(`FBLoginSub_hdfLoginHash = btoa(jQuery.now()+"|"+token)`)와 다중 시도 시 토큰 소모가 존재하므로 **자동화 프로덕션 코드에는 포함하지 않는다**(Melon/NOL 동일 UX). 본 탐색 과정에서는 사용했지만 구현에는 포함하지 않음.

---

## 4. 프로젝트 통합 설계

### 4.1 변경/추가 파일 (모두 Melon 쌍대응)

| 경로 | 역할 | 비고 |
|---|---|---|
| `src/KillRiceMonkey.Application/Models/TicketingTemplateType.cs` | `Yes24` enum 추가 | 기존 Booth/Nol/Melon/Custom 뒤에 추가 |
| `src/KillRiceMonkey.Application/Abstractions/IYes24AutomationService.cs` | 신규 인터페이스 | `IMelonAutomationService` 와 시그니처 동일 |
| `src/KillRiceMonkey.Infrastructure/Services/Yes24AutomationService.cs` | 메인 오케스트레이터 | `MelonAutomationService` 구조 그대로 복사 후 셀렉터·플로우 교체 |
| `src/KillRiceMonkey.Infrastructure/DependencyInjection.cs` | `AddSingleton<IYes24AutomationService, Yes24AutomationService>()` 등록 | |
| `src/KillRiceMonkey.App/ViewModels/MainWindowViewModel.cs` | YES24 상태 필드, 커맨드, 분기 추가 | `_yes24AutomationReadyAt`, `LaunchYes24RemoteDebugCommand`, `PrepareYes24AutomationCommand`, `IsYes24Template`, `StartAutomationAsync` 분기 |
| `src/KillRiceMonkey.App/MainWindow.xaml` | `Yes24` 템플릿 옵션 + 전용 패널 | Melon 패널 구조 재사용 (날짜/시간/준비/실행) |
| `src/KillRiceMonkey.App/MainWindow.xaml.cs` | 필요 시 핫키 바인딩 보강(기존 F8/F9 재사용 가능) | |
| (선택) `button-images/yes24/` | YES24 전용 이미지 템플릿 디렉토리 (초기엔 비어 있어도 됨) | |

> **새 아키텍처 도입 금지**. 기존 Melon 의 패턴(서비스 단위 잠금, Polly 재시도, 캐시된 `IBrowser`/`IPage`, 진행 상황 `IProgress<AutomationProgress>` 보고)을 그대로 재사용.

### 4.2 포트 할당

- NOL: **9222**
- Melon: **9223**
- **YES24: 9224** (신규)

### 4.3 프로필 디렉토리

- `%LOCALAPPDATA%\KillRiceMonkey\Yes24RemoteDebugProfile`

### 4.4 실행 URL

- 초기 실행 URL: `https://ticket.yes24.com/` (사용자가 공연 페이지로 이동)

### 4.5 `TicketingJobRequest` 확장 여부

- **재사용**: `DesiredDate`, `DesiredRound`, `PauseBeforeSeatSelection`, `PauseGate` 는 그대로 사용 가능.
- **추가 검토 필드**(본 작업에서 필요):
  - `string? DesiredIdTime` (사용자가 특정 회차 ID 를 직접 고정)
  - `string? DesiredGrade` (예: `"지정석 VIP석"`, `"일반석"`) — 부분 일치로 필터
  - `string? DesiredBlock` (구역 사이트에서 우선 구역, 예: `"2"`)
  - `int DesiredSeatIndex` (`1`-base, 정렬 후 몇 번째 좌석)
- 새 필드는 optional 로 도입해 기존 Melon/NOL 호출부 영향 없음.

---

## 5. `Yes24AutomationService` 구현 설계 (단계별)

Melon 을 레퍼런스로, 캡챠/대기열이 없으므로 오히려 구조가 더 단순하다.

### 5.1 상수 및 필드

```csharp
private const string Yes24RemoteDebugLaunchUrl = "https://ticket.yes24.com/";
private const int Yes24RemoteDebugPort = 9224;
private const string Yes24CdpEndpoint = "http://127.0.0.1:9224/";
private const string Yes24ProfileName = "Yes24RemoteDebugProfile";

private static readonly string[] BrowserExecutableCandidates = /* Melon 동일 */;

private IBrowser? _preparedYes24ConnectedBrowser;
private IPage? _preparedYes24Page;
private IPage? _preparedYes24SalePopup; // 예매 팝업 재사용 캐시
private readonly SemaphoreSlim _yes24BrowserLock = new(1, 1);
```

### 5.2 Public API (IYes24AutomationService 구현)

- `IsRemoteDebugBrowserAvailableAsync` → CDP 엔드포인트 핑
- `IsAutomationPreparedAsync` → `_preparedYes24Page` 재사용 가능성 + `IsLogin==='1'` 체크
- `LaunchRemoteDebugBrowserAsync` → Chrome 실행 (Melon 동일 템플릿)
- `PrepareAutomationAsync` → 공연 페이지 탐색 + `BringToFrontAsync` + 팝업 닫기 + OCR warmup (추후 캡챠 대비 포함)
- `IsPageReadyAsync` → `ticket.yes24.com/Perf/` 포함 여부 + `IsLogin==='1'`
- `RunAsync` → Polly 재시도 래퍼 + `RunYes24AutomationAsync`

### 5.3 RunYes24AutomationAsync 플로우 (핫 패스)

```
1. EnsurePreparedYes24ConnectedPageAsync
   - ticket.yes24.com/Perf/ 페이지 탐색 + 검증
   - BringToFrontAsync
2. SelectYes24DateAsync + SelectYes24TimeAsync (옵션 — DesiredIdTime 이 비어있으면 round 텍스트 매칭)
   - a[idTime="{DesiredIdTime}"] 찾기 + ScrollIntoViewIfNeededAsync + ClickAsync(force=true)
   - 실패 시 텍스트 매칭 fallback (PlaywrightRuntime.NormalizeText + idTime 텍스트 매칭)
3. TriggerYes24BookingPopupAsync  (Melon ClickMelonBookingAsync 상응)
   - 우선순위:
     (a) page.EvaluateAsync("() => jsf_base_ShowPerfSaleProcess(perf, time)")
         → Chrome popup blocker 때문에 trusted gesture 없으면 실패 → 350ms 대기 후 (b)
     (b) a.rn-bb03 Playwright ClickAsync(force:true)
   - page.Context.Pages 에서 새 페이지 중 URL contains "/Pages/Perf/Sale/PerfSaleProcess.aspx" 인 것을 대기 (polling 30ms, timeout = request.StepTimeoutSeconds)
   - 대기열 없음 — Melon 의 queuePage 분기 전부 생략
4. TriggerYes24SeatLoadAsync
   - salePopup.BringToFrontAsync
   - **⚠️ 팝업 alert 훅 설치**: `window.alert` + `fbk_Alert` 을 감싸서 `window.__yes24AlertDetected` 플래그로 캡처 (jQuery `fbk_Alert` 는 Playwright Dialog 이벤트로는 잡히지 않음)
   - **⚠️ 팝업 초기화 대기 (신규, 2026-04-09)**: `typeof fdc_FlashSeatLoad === 'function'` AND `#IdTime.value` 가 비어있지 않고 "0"이 아님 AND `#ulTime > li.on` 가 1개 이상일 때까지 30ms 폴링. 이 조건을 건너뛰면 약 520~617ms 사이의 race window 에서 `fbk_Alert("공연회차를 선택하세요.")` 가 발동되어 자동화가 멈춘다.
   - `__yes24AlertDetected` 리셋 후 `salePopup.EvaluateAsync("() => { fdc_FlashSeatLoad(); return 'ok'; }")`
   - 함수 미정의 시 폴백: `salePopup.Locator("#StepCtrlBtn01 a").First.ClickAsync`
   - 호출 직후 `__yes24AlertDetected` 재확인 → 알림 감지되면 예외로 상위 복구 유도
5. FindYes24SeatFrameAsync
   - Frames 순회 → url contains "PerfSaleHtmlSeat.aspx" 인 Frame 캐시
   - Fast-path: `[name=tk]` 가 존재하고 `divSeatArray > div` count > 0 이면 반환
   - 30ms 폴링(PlaywrightRuntime.PollDelayMilliseconds) + StepTimeoutSeconds 데드라인
6. EnsureYes24ZoneSelectedAsync  (구역 사이트 분기)
   - seatFrame.Locator("img[usemap]").CountAsync > 0 인 경우 구역 사이트.
   - 우선순위:
     (a) DesiredBlock 이 설정되어 있으면 frame.EvaluateAsync("(b) => ChangeBlock(parseInt(b))", block)
     (b) 자동 선택: area[href^='javascript:ChangeBlock'] 순회하며 ChangeBlock(N) → 30ms 후 divSeatArray 에 가용 좌석이 나오는지 확인 → 첫 성공 블록 채택
   - 좌석 로드 완료 조건: seatFrame.Locator("[name=tk][grade]").CountAsync > 0
7. SelectYes24SeatAsync
   - scanClickResult JS:
     const all = [...document.querySelectorAll('[name=tk]')];
     const avail = all.filter(e => typeof e.onclick === 'function' || e.hasAttribute('grade'));
     // optional grade filter
     const filtered = desiredGrade ? avail.filter(e => (e.getAttribute('grade')||'').includes(desiredGrade)) : avail;
     // 정렬: title(좌석번호)→ top→ left 순 → 안정적 순위
     filtered.sort((a,b)=>{ ... });
     const excluded = excludedSeats;
     const seat = filtered.find(e => !excluded.has(e.getAttribute('value')));
     if (!seat) return {s:'empty'};
     const r = seat.getBoundingClientRect();
     seat.click();  // jQuery on() 걸린 onclick 이어서 native click() 로 작동 (검증 필요)
     return {s:'clicked', v:seat.getAttribute('value'), t:seat.title, g:seat.getAttribute('grade')};
   - 결과 검증: `seatFrame.locator('[name=tk].son').count > 0` 또는 `#liSelSeat li` 존재
   - 검증 실패 → excludedSeats 에 value 추가 후 최대 maxSeatRetries 회 재시도
   - PauseGate 지원: 좌석 선택 직전 `gate.Wait(cancellationToken)` (Melon 동일)
8. ConfirmYes24SeatAsync
   - salePopup.EvaluateAsync("() => { try { fdc_VerifySelSeatNumber(); return 'ok'; } catch (e) { return 'err:'+e.message; } }")
   - 또는 #StepCtrlBtn01 내 a Playwright ClickAsync
   - 성공 조건: step01.display === 'none' AND step03.display === 'block' (polling 30ms, StepTimeoutSeconds)
9. AutomationRunResult(true, $"YES24 좌석 선택 완료: {seatTitle}", now)
```

### 5.4 실패 복구 & 재시도

- **Polly `MaxRetryAttempts = 2`** 외부 감싸기 (Melon 동일)
- **ExcludedSeats HashSet<string>**: 실패(중복·alert) 시 다음 시도에서 제외
- **좌석 재스캔 최대 10회**: Melon 의 `maxSeatRetries` 와 동일
- **제외 목록 리셋 fallback**: `ClickYes24SeatAsync` 가 `'not_found'` (필터링된 좌석은 존재하지만 전부 `excludedSeats` 에 포함) 를 반환하면 Melon/NOL 과 동일하게 `excludedSeats.Clear()` 후 재시도한다. 그 사이 다른 고객이 결제 포기해 풀린 좌석을 재시도할 기회를 주기 위함.
- **frame detached 감지**: `PlaywrightException` 감지 시 seat iframe 재탐색
- **팝업 닫힘 감지**: `salePopup.IsClosed === true` 면 외부 재시도로 통째 다시 트리거
- **dialog/alert 감지 (이중 경로)**:
  - 경로 1 (native): `salePopup.Dialog += handler` 로 native `alert/confirm/prompt` 를 자동 accept 하고 `window.__yes24AlertDetected = true` 세팅.
  - 경로 2 (jQuery): `fbk_Alert` 와 `window.alert` 를 JS 훅으로 감싸서 HTML 기반 jQuery 다이얼로그도 동일한 플래그로 캡처.
  - 좌석 중복 시 제외 처리는 두 경로 모두 플래그 기반 → `DetectYes24SeatConflictAsync` 가 양쪽을 포괄.

### 5.7 NetFunnel 대기열 처리 (2026-04-09 추가)

YES24 의 `jsf_base_ShowPerfSaleProcess` 함수는 perf ID 에 따라 두 가지 경로 중 하나를 탄다:

1. **Direct**: 일반 perf — 바로 `window.open(PerfSaleProcess.aspx)` 호출
2. **NetFunnel**: 하드코딩된 perf ID 리스트에 포함되거나 `#HidIsNetfunnel === '1'` 일 때 — `NetFunnel_Action({action_id, ...}, {success: ..., continued: ..., ...})` 호출. 대기열이 있는 동안 `netfunnel-skin.js` 의 `yesticket` 스킨이 `#NetFunnel_Skin_Top` div 를 메인 페이지에 inline 으로 삽입해 사용자에게 대기 UI 를 표시하고, success 콜백에서 비로소 `window.open + form.submit(POST netfunnel_key)` 를 실행한다.

**감지 방법**: `#NetFunnel_Skin_Top` 의 존재 + 가시성을 체크. `netfunnel-skin.js` 소스의 `yesticket` 스킨 템플릿에서 확인한 루트 셀렉터이다.

```csharp
// IsYes24QueueActiveAsync
return await page.EvaluateAsync<bool>(@"() => {
    const el = document.querySelector('#NetFunnel_Skin_Top');
    if (!el) return false;
    const style = window.getComputedStyle(el);
    if (style.display === 'none' || style.visibility === 'hidden') return false;
    const rect = el.getBoundingClientRect();
    return rect.width > 0 && rect.height > 0;
}");
```

**대기 플로우** (`WaitForYes24SalePopupWithQueueAsync`):

1. Phase 1 — `timeout` 초 제한 내에서 30ms 간격으로 폴링:
   - `PerfSaleProcess.aspx` 팝업 발견 시 즉시 반환
   - NetFunnel `#NetFunnel_Skin_Top` 감지 시 Phase 2 진입
2. Phase 1.5 — 타임아웃 직후 한 번 더 NetFunnel 재확인 (늦게 뜨는 경우 안전망)
3. Phase 2 — **무한 대기 루프** (Melon `ClickMelonBookingAsync` / NOL `ClickNolBookingAsync` 와 동일 패턴):
   - `while (!cancellationToken.IsCancellationRequested)` 로 100ms 폴링
   - `PerfSaleProcess.aspx` 팝업 감지 시 `"대기열 통과 ({초}초)"` 진행 보고 후 반환
   - 10초마다 `"대기열 대기 중... ({초}초)"` 진행 보고 (사용자 UX)
   - 메인 페이지가 닫히면 `InvalidOperationException`
4. 대기열 없이 timeout 만료 → `TimeoutException("YES24 예매 버튼 클릭 후 팝업 전환을 확인하지 못했습니다.")`

### 5.5 성능 최적화 포인트 (Melon 대비 추가 개선)

- **팝업 기다림 폴링 간격**: Melon 의 100ms 대신 **30ms** (대기열이 없어 팝업이 매우 빨리 뜸).
- **`jsf_base_ShowPerfSaleProcess` 직접 호출** 우선 → 예매하기 버튼을 DOM 탐색 없이 스킵.
- **`fdc_FlashSeatLoad` 직접 호출** 우선 → `StepCtrlBtn01` 클릭 생략.
- **iframe 캐싱**: 동일 iframe 재사용 시 재탐색 생략(Melon fast-path 와 동일 패턴).
- **병렬화 금지**(YES24 는 순차 상태 머신이라 병렬 호출 시 race). 단, `BringToFrontAsync` 와 JS 직접 호출은 한 turn 에 묶어 한 eval 로 전송.
- **사전 준비 단계**에서 연결된 `IBrowser`/`IPage` 를 캐시 → `RunAsync` 의 첫 CDP 연결 비용 제거.
- **OCR 웜업 선행**: 현재 캡챠 없음이지만 Melon 과 동일하게 `PlaywrightRuntime.EnsureDdddOcrWarmedUp()` 호출 → 후속 캡챠 확장 시 지연 없음.

### 5.6 캡챠 (향후 확장 대비)

- `GetSeatData_CallBack` 의 `cdn-botmanager.stclab.com` 리다이렉트 감지 후크를 `Yes24AutomationService` 내 `DetectYes24StclabCaptchaAsync` 로 분리.
- 현재는 **감지 + 사용자 알림(진행 일시정지)** 만 구현, 자동 해결은 Phase 2(추후 과제).
- 남겨두는 시그니처 예: `private Task<bool> TryHandleYes24StclabCaptchaAsync(IPage salePopup, CancellationToken ct) { /* TODO: future */ return Task.FromResult(false); }`

---

## 6. ViewModel / View 변경

### 6.1 MainWindowViewModel

- 필드 추가:
  ```csharp
  private readonly IYes24AutomationService _yes24AutomationService;
  private DateTimeOffset? _yes24AutomationReadyAt;
  private string? _yes24AutomationReadyMessage;
  ```
- 생성자 DI 에 `IYes24AutomationService` 추가.
- 커맨드 2개 추가:
  - `LaunchYes24RemoteDebugCommand`
  - `PrepareYes24AutomationCommand`
- 프로퍼티:
  - `IsYes24Template` (`SelectedTemplate == TicketingTemplateType.Yes24`)
  - `Yes24AutomationReadyUntil` (캐시 10분 유지 - Melon 규칙 준수)
- `StartAutomationAsync` switch 문에 Yes24 분기:
  - `IsPageReadyAsync` 폴링 (50ms → Melon 대비 동일, 빠른 반응성 유지)
  - `RunAsync` 호출
- `HandleHotkeyAsync` (F8) 는 Melon 분기와 같은 루트에 Yes24 분기 추가.
- `TemplateOptions` 에 `Yes24` 추가.
- `SelectedTemplate` 변경 시 `IsYes24Template` 변경 이벤트 발생.

### 6.2 MainWindow.xaml

- `ItemsControl` / `ComboBox` 바인딩 이미 있음 → `Yes24` 추가하면 자동 표시.
- Melon 전용 패널 구조 복사해 YES24 전용 패널(날짜/시간/실행/준비) 생성.
- `Visibility="{Binding IsYes24Template, Converter=...}"` 로 가시성 제어.
- 핫키 안내 재사용.

### 6.3 App.xaml.cs / DI

- `ConfigureServices` 에 특별 추가 없음 (Infrastructure 의 `AddInfrastructure` 가 자동으로 `IYes24AutomationService` 등록).
- `MainWindowViewModel` 은 DI 로 Yes24 서비스 주입.

---

## 7. 테스트 및 검증 계획

### 7.1 단위 테스트 (선택)

- Melon 에 단위 테스트가 거의 없으므로 YES24 도 동일하게 **서비스 수준 단위 테스트는 미작성**(회귀 리스크가 적은 래퍼 코드 위주).
- 필요 시 `PlaywrightRuntime` 의 기존 fake 를 확장해 `Yes24AutomationService.RunAsync` 인자 검증만 테스트.

### 7.2 빌드 검증

- `dotnet build KillRiceMonkey.sln -c Release` 클린 빌드 성공
- `dotnet test` 기존 테스트가 모두 그린 상태 유지
- `dotnet publish src/KillRiceMonkey.App -c Release -r win-x64 --self-contained false -o publish/kill-rice-monkey-yes24` 릴리즈 산출물 생성 (AGENTS.md 정책)

### 7.3 E2E QA (CDP 직접 검증 — AGENTS.md "작업 검수" 규칙)

1. 작업 완료 후 Chrome remote-debugging 으로 `ticket.yes24.com/Perf/57553` 접속 (또는 yes24.com 로그인 상태 유지)
2. 본 문서의 CDP 헬퍼 스크립트(`.tmp-yes24-qa/eval.py`)를 사용해
   - 시간 선택 → 예매하기 → `PerfSaleProcess.aspx` 오픈 확인
   - 좌석 iframe 로드 확인
   - 가용 좌석 클릭 후 `.son` 클래스 변화 확인
   - `fdc_VerifySelSeatNumber()` 호출 후 `#step03` display 변화 확인
3. 두 번째 URL `ticket.yes24.com/Perf/55989?Gcode=009_300` 도 동일하게 검증 + `ChangeBlock(N)` 호출 후 좌석 로드 확인
4. 전부 성공해야 **최종 작업 종료**. 실패 시 수정 → 재검증 무한 반복(AGENTS.md 명시 규칙).

### 7.4 성능 벤치마크 (목표)

- 사람 기준(버튼 클릭 → 좌석 선택 → 확인 평균 3~5초)의 **1초 이내 완료** 를 목표.
- 주요 지연 요소:
  - `ShowPerfSaleProcess` → `PerfSaleProcess.aspx` 팝업 로드: 300~800ms (네트워크 의존)
  - `fdc_FlashSeatLoad` → iframe GetBookWhole AJAX: 200~500ms
  - 좌석 스캔 & 클릭: 20~50ms
  - "다음단계" 이동 검증: 30~100ms
- **기대 총 합**: 약 600ms ~ 1.5s (사람보다 2~5배 빠름)

---

## 8. 커밋 플랜 (한국어, 작게 나누기)

AGENTS.md 정책: 한국어 메시지, 변수명/기술 용어는 영어 유지.

1. **feat(app): YES24 티켓팅 템플릿 enum 및 요청 필드 확장**
   - `TicketingTemplateType.Yes24`
   - `TicketingJobRequest` 에 `DesiredIdTime`, `DesiredGrade`, `DesiredBlock`, `DesiredSeatIndex` 추가(전부 optional)
2. **feat(app): IYes24AutomationService 추상화 추가**
   - 신규 인터페이스 파일
3. **feat(infra): Yes24AutomationService 스켈레톤 및 DI 등록**
   - Launch / Prepare / IsReady / RunAsync 시그니처 + Melon 상수/구조 복사
   - DependencyInjection 등록
4. **feat(infra): YES24 예매 프로세스 트리거 및 팝업 연결**
   - 시간 선택 → `jsf_base_ShowPerfSaleProcess` 직접 호출 → 팝업 대기 & 캐시
5. **feat(infra): YES24 좌석 iframe 로드 및 구역 분기**
   - `fdc_FlashSeatLoad` 호출 → iframe 탐색 → 구역/무구역 분기 구현
6. **feat(infra): YES24 좌석 스캔·선택·확인 로직**
   - scanClickResult JS, 재시도 루프, `fdc_VerifySelSeatNumber` 확정
7. **feat(app): MainWindowViewModel YES24 상태·커맨드·분기 추가**
   - Launch/Prepare 커맨드, `IsYes24Template`, `StartAutomationAsync` 분기
8. **feat(app): MainWindow.xaml YES24 템플릿 옵션 및 전용 패널**
   - ComboBox 옵션, 가시성 바인딩, 패널 레이아웃
9. **docs: YES24 개발현황 및 플랜 문서 추가**
   - 본 문서 (첫 커밋에 포함 가능)
10. **chore: .tmp-yes24-qa 탐색 산출물 제거**
    - 임시 파이썬/자바스크립트 QA 스크립트 삭제
11. **build: dotnet publish 릴리즈 산출물 생성 후 확인**

> 커밋 단위는 구현 진행에 맞춰 더 잘게 쪼갤 수 있음. 각 커밋 전 `dotnet build`(가능하면 `dotnet test`) 로 검증.

---

## 9. 구현 현황 (2026-04-09 기준)

| 항목 | 상태 | 비고 |
|---|---|---|
| `TicketingTemplateType.Yes24` | ✅ | enum 추가 완료 |
| `IYes24AutomationService` | ✅ | 신규 인터페이스 완료 |
| `Yes24AutomationService` | ✅ | 완료, 2026-04-09 팝업 초기화 race condition 수정 |
| DI 등록 | ✅ | `AddInfrastructure` 등록 완료 |
| `TicketingJobRequest` 확장 | ✅ | `DesiredIdTime`, `DesiredGrade`, `DesiredBlock`, `DesiredSeatIndex` 추가 |
| `MainWindowViewModel` 분기 | ✅ | 상태 + 커맨드 + 분기 완료 |
| `MainWindow.xaml` UI | ✅ | 템플릿 옵션 + 전용 패널 완료 |
| 본 플랜 문서 | ✅ | 본 파일 |
| 탐색용 QA 헬퍼(`.tmp-yes24-qa`, `.tmp-yes24-bug`) | ⚠️ 임시 | 작업 종료 시 제거 |

## 11. 2026-04-09 버그 수정 로그

### 증상
- 57918 (`Gcode=009_210_001`) 자동화 실행 시 관람일/회차 선택 페이지에서 `fbk_Alert("공연회차를 선택하세요.")` jQuery 다이얼로그가 뜬 채 상태가 "좌석 선택 중" 으로 넘어감.
- 8초 정도 대기 후 자동으로 복구되어 좌석 선택까지 완료되기도 하지만 (Playwright frame 폴링 루프 동안 팝업 JS가 사용자 모르게 스스로 다시 초기화), 안정적이지 않고 시간 손실이 크다.

### 원인
1. `TriggerYes24BookingPopupAsync` 는 팝업의 URL 이 `/Pages/Perf/Sale/PerfSaleProcess.aspx` 로 바뀌는 순간 즉시 `IPage` 를 반환한다.
2. `TriggerYes24SeatLoadAsync` 는 `typeof fdc_FlashSeatLoad === 'function'` 만 확인하고 즉시 호출한다.
3. 팝업 HTML 파싱이 진행 중인 약 **520~617ms 구간**에서는 `fdc_FlashSeatLoad` 함수는 이미 정의되어 있으나 `#IdTime.value` 가 아직 비어 있고 `#ulTime > li.on` 이 0개다.
4. 이 구간에 호출되면 `fdc_FlashSeatLoad` 내부 가드가 실패해 `fbk_Alert("공연회차를 선택하세요.")` 가 발동.
5. `fbk_Alert` 은 `jgBookAlert === jcMODE_JQUERY (=1)` 이므로 native alert 가 아닌 `$j("#dialogAlert").jAlert(...)` HTML 다이얼로그로 표시 → Playwright `Page.Dialog` 이벤트에 잡히지 않음.
6. 결과: 팝업 내 화면은 관람일/회차 선택에 머물고, 코드는 "좌석 선택 중" 으로 진행, 이후 `FindYes24SeatFrameAsync` 가 seat iframe 을 찾지 못해 타임아웃까지 소모.

### 수정 (`TriggerYes24SeatLoadAsync`)
1. 팝업 진입 직후 `window.alert` + `fbk_Alert` 에 JS 훅을 심어 `window.__yes24AlertDetected` / `window.__yes24LastAlert` 를 갱신하도록 변경.
2. `fdc_FlashSeatLoad` 호출 전에 아래 조건을 `PlaywrightRuntime.WaitForConditionAsync` (30ms 폴링) 로 대기:
   - `typeof fdc_FlashSeatLoad === 'function'`
   - `#IdTime` 엘리먼트의 `value` 가 비어있지 않고 `'0'` 이 아님
   - `#ulTime > li.on` 이 1개 이상 존재
3. 호출 전에 `__yes24AlertDetected = false` 로 리셋, 호출 후 `__yes24LastAlert` 확인하여 알림 발생 시 `InvalidOperationException` 으로 상위 재시도 유도.
4. 호출 결과가 `missing` 이면 기존대로 `#StepCtrlBtn01 a` 폴백 클릭.

### 검증
- CDP 로 팝업 raw timeline 측정: 약 618ms 이후 `#IdTime.value = '1429899'`, `#ulTime > li.on = 1` 로 안정화 → 수정된 대기 루프가 정확히 이 시점에서 통과.
- 57918 시나리오에서 수정 후 알림 다이얼로그가 한 번도 발생하지 않고 좌석 iframe 생성 → 좌석 클릭 → `fdc_VerifySelSeatNumber` → step01→step03 전환이 순차적으로 성공.

## 13. 2026-04-10 성능 최적화 — 중복 좌석 처리 속도

### 증상
- QA 재현 시나리오: 좌석 선택 전 일시정지 기능으로 진입 → 타 디바이스에서 목표 좌석 선점 → 재개.
- 결과: 자동화가 3초 이상 대기한 뒤에야 다른 좌석을 재선택. "티켓팅은 속도가 생명" 요건을 근본적으로 위배.

### 원인 (개선 전 기준)
1. **`selectionVerified` 고정 300ms 대기** (`SelectYes24SeatAndAdvanceAsync` 중간 단계)
   - `await TryWaitForConditionAsync([name=tk].son OR #liSelSeat > li, 300ms)` 로 좌석 확정만 대기.
   - 중복 케이스에서 `.son` 이 절대 붙지 않으므로 300ms 전부 소진 후에야 alert 체크로 이동.
2. **`ConfirmYes24SeatAsync` 내부 두 폴링 루프가 request timeout(기본 8s) 까지 대기**
   - 중복 좌석 확정 시도 후 step01→step03 전환이 영영 오지 않아 8초 timeout 만료.
   - 연쇄: `좌석 선택 → 300ms 낭비 → ConfirmYes24SeatAsync 진입 → 8s 대기 → 예외 → 재시도`.
3. **`DismissYes24SeatConflictAlertAsync` 가 `__yes24AlertDetected` flag 만 리셋**
   - jQuery `#dialogAlert` HTML 다이얼로그는 DOM 에 그대로 남아서 다음 iteration 의 click 에 UI 오염 가능.
4. **`FindYes24SeatFrameAsync` + `EnsureYes24SeatInventoryReadyAsync` 매 iteration 반복**
   - iframe 이 유지되고 inventory 가 안정된 이후에도 매 재시도마다 폴링 왕복 발생.

### 수정

**1) `WaitForYes24SeatOutcomeAsync` 신규 — 20ms 폴링 통합 outcome loop**
- `selectionVerified` 고정 300ms 대기 제거. 대신 20ms 주기로 세 신호를 병렬 체크:
  1. `salePopup` 의 `window.__yes24AlertDetected` — jQuery/native alert 훅이 세팅하는 flag
  2. `seatFrame.Locator("[name=tk].son")` — 좌석에 son class 부여 (iframe)
  3. `seatFrame.Locator("#liSelSeat > li")` — 선택 좌석 요약 리스트 (iframe)
- 최대 대기 시간 1500ms. 실측: 정상 케이스는 tick=1 (2ms) 내 confirmed, 중복 케이스는 알림이 도착하자마자 즉시 탈출.
- 반환 타입 `Yes24SeatOutcome` (Confirmed / Conflict / Unknown) record struct 로 분기 명확화.

**2) `SelectYes24SeatAndAdvanceAsync` 리팩토링**
- `seatFrame` 과 `inventoryReady` 를 iteration 외부에 캐싱. `PlaywrightException` 발생 시에만 null 로 재설정해 재탐색 강제.
  - 매 재시도마다 `FindYes24SeatFrameAsync` + `EnsureYes24SeatInventoryReadyAsync` 호출을 **N회 → 1회 + 재연결 시 1회** 로 감축.
- click 직전 `window.__yes24AlertDetected / __yes24LastAlert` 를 리셋하여 outcome loop 가 이전 iteration 의 잔재 플래그를 감지하지 않도록 보호.
- outcome 분기를 `Confirmed → ConfirmYes24SeatAsync` / `Conflict → Dismiss + continue` / `Unknown → excludedSeats + continue` 로 단순화.

**3) `DismissYes24SeatConflictAlertAsync` 강화**
- `__yes24AlertDetected = false / __yes24LastAlert = null` 리셋 외에도:
  - `jQuery('#dialogAlert').dialog('close')` 시도 (`$j` / `jQuery` 둘 다 fallback)
  - DOM 레벨에서 `#dialogAlert` 및 `.ui-dialog` wrapper, `.ui-widget-overlay` 모두 `display:none` 으로 강제 hide
- 이로써 다음 iteration 의 click 이 기존 HTML 다이얼로그의 이벤트 버블링에 영향받지 않음.

**4) `ConfirmYes24SeatAsync` 내부 timeout cap**
- 기존 `timeout` (기본 8s) 를 `confirmTimeout = min(timeout, 2500ms)` 로 cap.
- `preconditionMet` (selSeatClass 등장) 과 `advanced` (step01→step03 전환) 두 폴링 루프 모두 `confirmTimeout` 사용.
- 정상 케이스는 100~500ms 내 완료되므로 2.5s 면 5x 여유. 중복·오류 케이스는 2.5s 로 cap 되어 상위 재시도가 빠르게 트리거.
- `advanced` 폴링의 조건 함수 내부에서도 `DetectYes24SeatConflictAsync` 확인 시 false 반환하여 알림 발동 시점에 조기 탈출.

### 검증

- **정상 baseline (57918)**: 
  - `좌석 선택 중 → 좌석 선택 완료 = 774~1156ms`
  - outcome loop 실측: tick=1, elapsed=2ms → Confirmed. 이전 300ms 고정 대기 대비 ~150x 빠름.
- **중복 시나리오 end-to-end**: `run_conflict_scenario.py` — iframe 내부 자체 realm 에서 `ChoiceSeat` 1회성 override 로 fake duplicate alert 주입. 
  - 총 `좌석 선택 중 → 좌석 선택 완료 = 1421ms`
  - **conflict 감지 → 재선택 완료 = 247ms** (기존 3000ms+ → 247ms, 10x+ 개선)
  - `conflicts observed: 1`, seat 4700074 제외 후 다른 좌석(1층 1구역04열 002번) 성공 선택
  - `[RESULT] PASS — conflict handled in 1422ms (target <1500ms)`
- **회귀**: dotnet build 0 errors (warnings 2 기존), dotnet test 4/4 pass, dotnet publish 성공.

### 동작 플로우 (2026-04-10 Phase 2B Dialog event 최종)

```
SelectYes24SeatAndAdvanceAsync:
  [1회 only] FindYes24SeatFrameAsync + EnsureYes24SeatInventoryReadyAsync
  [1회] C# Dialog event handler 등록 (dialogFlag int[] + phase2BHandler)
  for seatAttempt in 0..9:
    Phase 0: .son 클린업 + JS flag reset + C# dialogFlag reset
    Phase 1: ClickYes24SeatAsync → target.click()
    Phase 2: TryWaitForConditionAsync(.son > 0 OR #liSelSeat > 0, 100ms)
         - 미반영 → excludedSeats + continue (즉시 다음 좌석)
    Phase 2B: C# Dialog event 기반 중복 감지 (10ms polling, max 300ms)
         - Volatile.Read(dialogFlag[0]) == 1 → 즉시 reject (~10ms)
         - 300ms 내 미감지 → JS flag 보완 체크 (fbk_Alert jQuery dialog 대응)
         - rejected → excludedSeats + DismissYes24SeatConflictAlertAsync + continue
    Phase 3: ConfirmYes24SeatAsync (cap 1000ms, dialogFlag 전달)
         - ChoiceEnd() → selSeatClass 대기 → fdc_VerifySelSeatNumber → step 전환
         - polling 루프 내 C# dialogFlag 우선 체크로 native alert 교착 방지
         - 예외 시 excludedSeats + continue, 아니면 return title
```

### 근본 병목 분석 (0.83s → <0.05s)

**증상**: 중복 좌석 재시도에 좌석당 0.83s 소요. 사용자 평가 "인간이 더 빠름".

**근본 원인**: native `alert()` 가 JS main thread 를 차단하면 Playwright 의 CDP `Runtime.evaluate`
명령이 응답을 받을 수 없어 `EvaluateAsync` 가 교착. Phase 2B 의 15ms polling 이 전부 교착 상태로
300ms timeout 만료 → Phase 3 진입 → Phase 3 의 `DetectYes24SeatConflictAsync` 도 동일 교착 →
`confirmTimeout` (1000ms) 근처까지 소진 후에야 Dialog handler 의 `AcceptAsync` 가 처리됨.

**핵심 메커니즘**: Playwright 내부에서 CDP 명령이 직렬화되어 `Runtime.evaluate` (alert 차단) 와
`Page.handleJavaScriptDialog` (AcceptAsync) 가 상호 대기하는 사실상 교착 상태.

**해결**: Playwright `Page.Dialog` C# event 는 CDP `Page.javascriptDialogOpening` 알림으로
발동하므로 JS 차단과 무관하게 즉시 감지 가능. `EvaluateAsync` 를 제거하고 C# `int[]` flag +
`Volatile.Read` 10ms polling 으로 교체하여 교착 완전 해소.

- Phase 2B: `EvaluateAsync` 완전 제거 → `Volatile.Read(ref dialogFlag[0])` 10ms polling
- Phase 3: `DetectYes24SeatConflictAsync` 앞에 `Volatile.Read` 게이트 추가로 교착 방지
- fbk_Alert 보완: 300ms 후 JS flag 1회 체크 (jQuery dialog 는 Dialog event 미발동, JS 비차단)

**기대 성능**: 중복 좌석 감지 Phase 2B 내 ~10ms (이전 300ms+ 교착 → 10ms), 총 재시도 시간
Phase 2 (30ms) + Phase 2B (alert 시점 ~150ms + 감지 10ms) + dismiss (20ms) ≈ **0.2s/좌석** 목표.

## 12. 2026-04-09 후속 개선 로그 — 대기열 처리 + 제외 목록 리셋

초기 플랜에 "YES24 는 대기열 없음" 이라고 기재했으나, 재검토 결과 `jsf_base_ShowPerfSaleProcess` 가 일부 perf 에 대해 `NetFunnel_Action` 으로 대기열을 사용함을 발견. Melon/NOL 과 달리 YES24 는 대기열을 **별도 페이지가 아닌 메인 페이지에 inline 모달** (`#NetFunnel_Skin_Top`) 로 표시하기 때문에 기존 코드의 "새 팝업 URL 매칭" 로직으로는 감지 불가능했다. 또한 `SelectYes24SeatAndAdvanceAsync` 의 `excludedSeats` 는 한 번 채워지면 시도 종료까지 비워지지 않아 Melon/NOL 의 fallback 리셋 로직이 빠져있었다.

### 증상
1. **대기열 처리 누락** — 경쟁이 많은 공연에서 NetFunnel 대기열이 뜨면 `TryWaitForYes24SalePopupAsync` 가 `PerfSaleProcess.aspx` 만 기다리므로 `StepTimeoutSeconds` 후 `TimeoutException` 으로 실패.
2. **제외 목록 리셋 없음** — 10회 재시도 동안 중복 감지된 좌석이 누적되어 가용 좌석이 전부 `excludedSeats` 에 들어가면 `'not_found'` 가 반복되어 조기 실패. 실제로는 재시도 사이에 다른 고객이 결제 포기해 좌석이 풀릴 수 있으므로 리셋 기회가 필요.

### 수정
**1) `TriggerYes24BookingPopupAsync` + 신규 `WaitForYes24SalePopupWithQueueAsync` + 신규 `IsYes24QueueActiveAsync`**

```
TriggerYes24BookingPopupAsync(page, idTime, timeout, progress, ct)
  ├─ 캐시/기존 팝업 검사 → 재사용
  ├─ jsf_base_ShowPerfSaleProcess 직접 호출 시도 (popup blocker 로 보통 실패)
  ├─ a.rn-bb03 클릭 (trusted)
  └─ WaitForYes24SalePopupWithQueueAsync(page, timeout, progress, ct)
        ├─ Phase 1: 30ms 폴링, timeout 까지
        │     ├─ PerfSaleProcess.aspx 감지 → 반환
        │     └─ #NetFunnel_Skin_Top 감지 → Phase 2 진입
        ├─ Phase 1.5: timeout 만료 후 한번 더 NetFunnel 재확인
        └─ Phase 2: while (!ct.IsCancellationRequested) 무한 대기
              ├─ 10초마다 "대기열 대기 중... (N초)" progress
              ├─ PerfSaleProcess.aspx 감지 → "대기열 통과 (N초)" 후 반환
              └─ 메인 페이지 닫힘 감지 → InvalidOperationException
```

`IsYes24QueueActiveAsync` 는 `#NetFunnel_Skin_Top` 의 존재 + `display/visibility/bounding rect` 검사로 활성 여부 판정.

**2) `SelectYes24SeatAndAdvanceAsync` 제외 목록 리셋 fallback**

`ClickYes24SeatAsync` 가 반환하는 status:
- `'clicked'`: 정상 클릭
- `'empty'`: 가용 좌석 없음 (원본 카운트 0)
- `'not_found'`: 가용 좌석은 존재하나 전부 `excludedSeats` 에 포함

`'not_found'` 이고 `excludedSeats.Count > 0` 이면 Melon/NOL 과 동일하게 다음을 수행:
```csharp
_logger.LogWarning("[YES24] 제외 좌석 {Count}개를 빼면 선택 가능한 좌석 없음 — 제외 목록 초기화 후 재시도. ...");
excludedSeats.Clear();
progress?.Report(new AutomationProgress("좌석 재선택 중", "제외 목록 초기화 후 재시도"));
```

### 검증
- **대기열 end-to-end 시나리오** (`run_queue_scenario.py`):
  1. 메인 페이지에 `jsf_pdi_GoPerfSale` + `jsf_base_ShowPerfSaleProcess` 1회성 override 설치 — 첫 호출 시 `#NetFunnel_Skin_Top` fake 요소 주입 + 원본 suppress.
  2. 실제 `Yes24AutomationService.RunAsync` 실행.
  3. 자동화가 1.49s 에서 "대기열 감지" 로그 출력 및 `"대기열 대기 중..."` progress 보고.
  4. Python 테스트가 2.5초 후 fake 요소 제거 + CDP trusted click 으로 2차 클릭 발사.
  5. 원본 `jsf_pdi_GoPerfSale` 실행 → 실제 팝업 오픈.
  6. 자동화가 팝업 감지 → `"대기열 통과 (3초)"` 보고 → 좌석 iframe 로드 → 좌석 선택 → 완료.
  - 결과: `success=True`, `saw_queue=True`, `saw_release=True`, rc=0 — **PASS**.
- **제외 목록 리셋 JS 검증** (`test_js_synthetic.py`): 합성 DOM 3좌석(S001/S002 VIP, S003 R) 에 대해 `ClickYes24SeatAsync` JS 7 케이스 검증.
  - empty excluded → clicked S001 ✅
  - exclude S001 → clicked S002 ✅
  - exclude S001+S002 → clicked S003 ✅
  - exclude ALL → `'not_found'` ✅ **리셋 트리거**
  - grade=VIP → clicked S001 ✅
  - grade=VIP + exclude VIP → `'not_found'` ✅
  - grade=NONEXISTENT → `'empty'` ✅
- **회귀**: baseline 57918 정상 실행 (warm 2s, cold 7~16s with/without conflict retry), `dotnet test` 4/4 통과, `dotnet build` 0 errors.

---

## 10. 자동화 UX 요약 (사용자 시점)

1. 앱 실행 → 템플릿 드롭다운에서 **YES24** 선택.
2. **"Yes24 Remote Debug 브라우저 열기"** 클릭 → 프로필 영속 Chrome 실행.
3. 최초 1회: 사용자가 yes24 로그인. 세션 유지됨.
4. 공연 페이지로 이동 (`https://ticket.yes24.com/Perf/XXXXX`).
5. **"자동화 준비"** 클릭 → 백그라운드로 페이지 캐시 + 로그인 상태 검증 (약 10분 유효).
6. 앱에서 날짜/회차 입력 (선택적으로 IdTime 직접 지정, 선호 구역/등급 입력).
7. **F8** 또는 "시작" 버튼 → 자동화 트리거:
   - 시간 선택 → 예매 팝업 열림 → 좌석 iframe → 구역 분기 → 좌석 클릭 → "다음단계" 진입 → 완료.
8. 완료 후 사용자가 결제 진행(본 자동화는 결제까지 관여하지 않음 - 과제 범위 외).

---
