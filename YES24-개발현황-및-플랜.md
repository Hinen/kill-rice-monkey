# YES24 티켓 자동화 개발 현황 및 플랜

## 1. 조사 개요

- **조사일**: 2026-04-08
- **조사 방식**: Chrome remote-debugging(CDP) + websockets로 두 URL 모두 직접 접속 → 로그인 → 시간 선택 → 예매 버튼 → 좌석 iframe까지 전 과정 DOM 구조 확인
- **무구역 사이트**: https://ticket.yes24.com/Perf/57553 (Sou LIVE TOUR 2026 「Finder」 in Seoul)
- **구역 사이트**: https://ticket.yes24.com/Perf/55989?Gcode=009_300 (tuki. 1ST ASIA TOUR 2026 IN SEOUL)
- **로그인**: `yes24.txt` 자격증명(hinen/casa15)으로 실제 CDP 로그인 성공 확인
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
- **페이지 진입 시 상태**:
  - `jgCalSelDate = "YYYY-MM-DD"` (IdTime 으로부터 자동 결정됨)
  - `#IdTime.val() === {IdTime}` / `#ulTime > li.on` 이 이미 선택되어 있음
  - `#step01`, `#step01_date`, `#step01_time`, `#step01_notice` 가 `display: block`
  - `#step03`, `#step04`, `#step05` 는 `display: none`
- **StepCtrlBtn01 (다음단계 버튼)**: `a` onclick=`fdc_VerifySelSeatNumber()`
  - 현재 좌석 class 가 없으면 → `fdc_FlashSeatLoad()` 호출
  - `fdc_FlashSeatLoad()` → `fdc_CtrlStep(jcSTEP_SEAT)` → `#SeatFlashArea` 내부에 **좌석 iframe 생성**
- **좌석 iframe 생성 URL** (예시):
  - 무구역: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleHtmlSeat.aspx?idTime=1425353&idHall=14062&block=0&stMax=10&pHCardAppOpt=0`
  - 구역: `https://ticket.yes24.com/Pages/Perf/Sale/PerfSaleHtmlSeat.aspx?idTime=1397006&idHall=13754&block=2&stMax=10&pHCardAppOpt=0`
- **자동화 관점 최적화 규칙**:
  1. 팝업이 열리자마자 바로 `fdc_FlashSeatLoad()` 를 호출(검증 함수 경유 불필요).
  2. 동일 orgin 이므로 좌석 iframe 접근은 `iframe.contentDocument` 또는 Playwright `Frame` API 로 직접 가능.

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
2. SelectYes24TimeAsync (옵션 — DesiredIdTime 이 비어있으면 스킵 가능)
   - a[idTime="{DesiredIdTime}"] 찾기 + ScrollIntoViewIfNeededAsync + ClickAsync(force=true)
   - 실패 시 텍스트 매칭 fallback (Melon NormalizeText + idTime 텍스트 매칭)
3. TriggerYes24BookingPopupAsync  (Melon ClickMelonBookingAsync 상응)
   - 우선순위:
     (a) page.EvaluateAsync("() => typeof jsf_base_ShowPerfSaleProcess === 'function' && jsf_base_ShowPerfSaleProcess($('#HidIdPerf').val(), $('.rn-04-left-calist a.on').attr('idTime'))")
         → 이게 실패하면 (b)
     (b) a.rn-bb03 Playwright ClickAsync(force:true)
   - page.Context.Pages 에서 새 페이지 중 URL contains "/Pages/Perf/Sale/PerfSaleProcess.aspx" 인 것을 대기 (polling 30ms, timeout = request.StepTimeoutSeconds)
   - 대기열 없음 — Melon 의 queuePage 분기 전부 생략
4. TriggerYes24SeatLoadAsync
   - salePopup.BringToFrontAsync
   - salePopup.EvaluateAsync("() => { if (typeof fdc_FlashSeatLoad === 'function') { fdc_FlashSeatLoad(); return 'ok'; } return 'missing'; }")
   - 실패 시 fallback: salePopup.Locator("#StepCtrlBtn01 a").First.ClickAsync
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
- **frame detached 감지**: `PlaywrightException` 감지 시 seat iframe 재탐색
- **팝업 닫힘 감지**: `salePopup.IsClosed === true` 면 외부 재시도로 통째 다시 트리거
- **dialog/alert 감지**: `salePopup.Dialog += handler` 로 자동 accept + `window.__yes24AlertDetected = true` 플래그 저장 → 좌석 중복 시 제외 처리

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

## 9. 현재 미구현 요약 (시작 시점 기준)

| 항목 | 상태 | 비고 |
|---|---|---|
| `TicketingTemplateType.Yes24` | ❌ | enum 추가 필요 |
| `IYes24AutomationService` | ❌ | 신규 인터페이스 |
| `Yes24AutomationService` | ❌ | 신규 구현 |
| DI 등록 | ❌ | `AddInfrastructure` 에 1줄 추가 |
| `TicketingJobRequest` 확장 | ❌ | optional 필드 4개 |
| `MainWindowViewModel` 분기 | ❌ | 상태 + 커맨드 + 분기 |
| `MainWindow.xaml` UI | ❌ | 템플릿 옵션 + 전용 패널 |
| 본 플랜 문서 | ✅ | 본 파일 |
| 탐색용 QA 헬퍼(`.tmp-yes24-qa`) | ⚠️ 임시 | 작업 종료 시 제거 |

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
