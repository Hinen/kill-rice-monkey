# NOL 티켓 자동화 개발 현황 및 플랜

## 1. 조사 개요

- **조사일**: 2026-03-23
- **NEW 버전**: https://tickets.interpark.com/goods/25017910 (아마자라시 아시아투어 2026 in SEOUL)
- **OLD 버전**: https://tickets.interpark.com/goods/25016435 (tuki. 1ST ASIA TOUR 2026 IN SEOUL)
- **방법**: Chrome remote-debugging(CDP)으로 **두 페이지 모두** 직접 접속, 카카오 로그인 수행, 예매하기 클릭 → 대기열 통과 → CAPTCHA/좌석 선택 페이지까지 직접 진입하여 DOM 구조 분석

---

## 2. NEW vs OLD의 진짜 차이: 두 개의 완전히 다른 예매 시스템

### 상품 페이지 (`tickets.interpark.com/goods/*`)

상품 페이지 단계에서는 NEW/OLD **동일한 DOM 구조**:
- `#productSide`, `.sideCalendar`, `.sideTimeTable`, `a.sideBtn.is-primary`
- 모두 React 기반, 동일한 셀렉터
- 차이는 콘텐츠(공연명, 날짜, 좌석 잔여)뿐

### 예매하기 클릭 후: 완전히 다른 시스템으로 분기

| 항목 | OLD (레거시) | NEW (원스탑) |
|------|-------------|-------------|
| **예매 URL** | `poticket.interpark.com/Book/BookMain.asp` | `tickets.interpark.com/onestop/seat` |
| **프레임워크** | 레거시 HTML + 3단 iframe 중첩 | **Next.js React SPA** (`#__next`) |
| **좌석맵** | `<img usemap>` + `<area>` 42개 (이미지맵) | **SVG `<circle>` 2,926개** (벡터) |
| **CAPTCHA** | `#txtCaptcha` + `#imgCaptcha` (iframe 내부) | **React 모달** `ModalCaptchaText_*` |
| **구역 선택** | `GetBlockSeatList('','','BLOCK_CODE')` | SVG `<g>` 25개 그룹 (구역 없이 전체 좌석 표시) |
| **좌석 표현** | `<area>` on image map | `<circle>` in SVG (클릭 가능) |
| **선택 완료** | `formBook` 제출 (hidden inputs) | **`EntButton_*` "선택 완료" 버튼** |
| **iframe 구조** | `ifrmSeat` > `ifrmSeatDetail` > `ifrmInfo` | iframe 없음 (전부 단일 페이지) |

---

## 3. 실제 예매 플로우 (두 버전 모두 직접 확인)

### 3.1 공통 플로우 (상품 페이지 → 대기열)

```
[1] 상품 페이지 (tickets.interpark.com/goods/XXXXX)
    ├── 팝업 닫기: .popup.is-visible .popupCloseBtn
    ├── 날짜 선택: .sideCalendar ul[data-view='days'] > li
    ├── 회차 선택: .sideTimeTable .timeTableLabel[role='button']
    └── 예매하기 클릭: a.sideBtn.is-primary (React debounced onClick, CDP trusted click 필요)
            │
            ▼ (비로그인 시)
[2] 로그인 (accounts.yanolja.com)
    ├── "카카오로 시작하기" 버튼 (CDP trusted click, JS .click() 불가)
    ├── 카카오 로그인 팝업 (accounts.kakao.com, 새 창)
    │   ├── #loginId--1 (아이디) — nativeInputValueSetter 필요
    │   ├── #password--2 (비밀번호)
    │   └── button[type=submit] "로그인"
    └── OAuth 완료 → 팝업 닫힘 → 상품 페이지 복귀 → 예매하기 재클릭
            │
            ▼
[3] 대기열 (tickets.interpark.com/waiting?key=...)
    └── Page.windowOpen 이벤트로 새 창 감지
```

### 3.2 OLD 플로우 (레거시 시스템)

```
[4-OLD] poticket.interpark.com/Book/BookMain.asp
    ├── #divBookSeat (display:block)
    └── #ifrmSeat
        ├── CAPTCHA 오버레이
        │   ├── #divCaptchaWrap (opacity:0.7)
        │   ├── #imgCaptcha (base64)
        │   ├── #txtCaptcha
        │   ├── fnCheck() → "입력완료"
        │   ├── fnCapchaRefresh()
        │   └── CaptchaHide()/CaptchaShow()
        │
        ├── #ifrmSeatDetail (BookSeat.asp)
        │   ├── #TmgsTable
        │   ├── <img usemap="#MapMapMap"> (600×530)
        │   └── <map> → 42개 <area>
        │       └── GetBlockSeatList('','','BLOCK_CODE') ← 구역 선택
        │
        ├── #SeatGradeInfo, #SelectedSeat
        └── fnSeatUpdate(), fnSeatPrice()
```

### 3.3 NEW 플로우 (원스탑 시스템) ⭐

```
[4-NEW] tickets.interpark.com/onestop/seat (Next.js SPA)
    ├── CAPTCHA 모달 (최상위 레벨, iframe 없음)
    │   ├── .ModalCaptchaText_layerWrap__*
    │   ├── .ModalCaptchaText_captchaImage__* (base64 이미지)
    │   ├── .ModalCaptchaText_captchaInput__* (input, placeholder: "화면의 문자를 입력해주세요")
    │   ├── .ModalCaptchaText_buttonRefresh__* (새로고침)
    │   ├── .ModalCaptchaText_buttonSound__* (음성)
    │   └── "입력완료" 버튼 (EntButton_*, disabled until input)
    │
    ├── SVG 좌석맵 (CAPTCHA 뒤에 이미 로드됨)
    │   ├── .SeatPlan_seatPlan__*
    │   ├── .SeatMap_seatGroup__*
    │   ├── <svg viewBox="0 0 459 321">
    │   │   ├── 25개 <g> 그룹 (구역)
    │   │   └── 2,926개 <circle> (개별 좌석)
    │   ├── 줌 컨트롤: .SeatPlan_zoomButton__* (in/out/refresh)
    │   └── .EntZoomableWrapper_* (핀치줌/드래그)
    │
    ├── 등급 정보
    │   ├── .SeatGradeLayer_* ("등급 별 가격" 토글)
    │   └── R석 132,000원 / S석 121,000원 / A석(시야제한) 110,000원
    │
    ├── 선택 좌석 정보
    │   ├── .InfoSelected_outerWrap__*
    │   └── "선택한 좌석이 없습니다."
    │
    ├── "예매대기" 버튼: .ButtonWaiting_buttonWaiting__*
    └── "선택 완료" 버튼: .EntButton_* (disabled until seat selected)
```

---

## 4. 현재 개발 완료된 기능

### 4.1 NOL 자동화 플로우 (`NolAutomationService.cs`, ~2,000줄)

| 단계 | 기능 | 상태 | OLD 호환 | NEW 호환 |
|------|------|------|---------|---------|
| 0 | 브라우저 실행 | ✅ | ✅ | ✅ |
| 1 | 자동화 준비 | ✅ | ✅ | ✅ |
| 2 | 팝업 닫기 | ✅ | ✅ | ✅ |
| 3 | 날짜 선택 | ✅ | ✅ | ✅ |
| 4 | 회차 선택 | ✅ | ✅ | ✅ |
| 5 | 예매 클릭 + 대기열 | ✅ | ✅ | ✅ |
| 6 | CAPTCHA 풀기 | ✅ | ✅ (OLD iframe + NEW 모달 모두 대응) | ✅ |
| 7 | 로그인 | ⚠️ 기존 세션 활용 (자동 로그인 미구현) | - | - |
| 8 | 구역/좌석 선택 | ✅ | ✅ (레거시 span[onclick] 클릭) | ✅ (SVG circle 클릭) |
| 9 | 선택 완료 | ✅ | ✅ (fnSelect/NextStepImage) | ✅ (Playwright trusted click) |

### 4.2 CAPTCHA 코드 호환성 비교

| 현재 코드 셀렉터 | OLD 페이지 | NEW 페이지 |
|---|---|---|
| `#txtCaptcha` | ✅ `ifrmSeat` 내 존재 | ❌ 없음 (`.ModalCaptchaText_captchaInput__*`) |
| `#imgCaptcha` | ✅ `ifrmSeat` 내 존재 | ❌ 없음 (`.ModalCaptchaText_captchaImage__* img`) |
| `fnCheck()` | ✅ 존재 | ❌ 없음 (React 버튼 클릭) |
| `fnCapchaRefresh()` | ✅ 존재 | ❌ 없음 (`.ModalCaptchaText_buttonRefresh__*`) |
| `button:text-is('입력완료')` | ✅ | ✅ (동일 텍스트, 다른 class) |

**결론**: CAPTCHA 코드는 OLD/NEW 모두 대응 완료. `FindNolCaptchaInputAsync`가 OLD iframe 내 `#txtCaptcha`와 NEW 모달의 `[class*='captchaInput']` 양쪽을 검색한다.

---

## 5. 미구현 / 향후 개선 사항

### 5.1 [미구현] 카카오 로그인 자동화

현재는 사용자가 미리 브라우저에서 로그인한 세션을 재사용한다.

**직접 확인한 플로우 기반**:
```
1. 예매 클릭 → accounts.yanolja.com 리다이렉트 감지
2. "카카오로 시작하기" 버튼 CDP trusted click
3. 새 창 감지: accounts.kakao.com (Page.windowOpen 이벤트 아닌 CDP Target 이벤트)
4. #loginId--1 에 ID 입력 (nativeInputValueSetter)
5. #password--2 에 PW 입력
6. button[type=submit] "로그인" 클릭
7. 팝업 닫힘 대기 → 원래 페이지 복귀 → 예매하기 재클릭
```

### 5.2 [참고] NEW/OLD 시스템 분기 감지 (구현 완료)

`FindNolSeatPageAsync`에서 URL 패턴으로 자동 분기:
- `tickets.interpark.com/onestop` → NEW 원스탑 플로우 (`SelectNolOnestopSeatAndCompleteAsync`)
- `poticket.interpark.com` → OLD 레거시 플로우 (`SelectNolLegacySeatAndCompleteAsync`)

### 5.3 [참고] NEW 원스탑 좌석 선택 (구현 완료)

```
구현 메서드:
├── WaitForNolOnestopSeatMapAsync()       -- SVG/SeatPlan 로드 대기
├── IsNolOnestopZoneSelectionRequiredAsync() -- 구역 선택 필요 여부 판단
├── WaitForNolOnestopZoneSelectionAsync() -- 구역 선택 대기 (사용자 클릭)
├── SelectNolOnestopSeatAsync()           -- SVG circle 스캔 + 빈 좌석 클릭
├── ClickNolOnestopSeatCompleteAsync()    -- "선택 완료" Playwright trusted click + 검증
└── IsOnestopSeatCompleteConfirmedAsync() -- 클릭 후 페이지 전환/이탈 확인
```

### 5.4 [참고] OLD 레거시 좌석 선택 (구현 완료)

```
구현 메서드:
├── FindNolLegacyFrameAsync()             -- ifrmSeat/ifrmSeatDetail 프레임 탐색
├── HasNolLegacySelectSeatSpansAsync()    -- SelectSeat span 존재 여부
├── WaitForNolLegacySelectSeatSpansAsync() -- 구역 선택 대기
├── SelectNolLegacySeatAsync()            -- span[onclick*=SelectSeat] 클릭
└── ClickNolLegacySeatCompleteAsync()     -- fnSelect()/NextStepImage 호출
```

---

## 6. 향후 개발 플랜

### Phase 1: 카카오 로그인 자동화 (미구현)

**파일**: `NolAutomationService.cs`

```
추가 메서드:
├── DetectLoginRedirectAsync()           -- yanolja.com 리다이렉트 감지
├── LoginWithKakaoAsync()
│   ├── ClickKakaoButtonAsync()          -- CDP trusted click (React 이벤트)
│   ├── WaitForKakaoPopupAsync()         -- 새 창 감지 (browser.Contexts pages)
│   ├── FillKakaoCredentialsAsync()      -- nativeInputValueSetter 방식
│   └── WaitForPopupCloseAsync()         -- 팝업 닫힘 + 원래 페이지 복귀
└── RetryBookingAsync()                  -- 로그인 후 예매하기 재클릭
```

### Phase 2: Mock 서버 확장

**파일**: `tools/mock-ticket-server/Pages/NolPages.cs`

- NEW 원스탑 좌석 선택 페이지 Mock 추가
- SVG circle 좌석맵 + CAPTCHA 모달

### Phase 3: UI 확장

- `TicketingJobRequest`에 `PreferredGrade` (선호 등급: R/S/A) 추가
- 카카오 로그인 정보 설정 (환경변수 또는 설정 파일)

---

## 7. 기술적 주의사항

1. **React debounced onClick**: 상품 페이지 예매 버튼과 NOL 로그인 페이지 버튼 모두 JS `.click()` 불가 → **CDP `Input.dispatchMouseEvent`** (trusted click) 필수
2. **카카오 로그인 팝업**: `window.open`으로 별도 창 → `browser.Contexts.Pages`에서 새 페이지 감지 가능
3. **nativeInputValueSetter**: 카카오 로그인 + NEW CAPTCHA 모달의 React controlled input에 값 설정 시 `Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set` 사용 필요
4. **NEW 시스템은 iframe 없음**: `page.Locator()`로 직접 접근 가능 — 기존 iframe 탐색 로직 불필요
5. **SVG circle 좌석**: 좌석 상태는 `fill`, `stroke`, `class`, `data-*` 속성으로 구분될 가능성 높음 — 실제 CAPTCHA 통과 후 추가 조사 필요
6. **CSS module 클래스명**: `ModalCaptchaText_captchaInput__DC7Gz` 같은 해시 suffix가 빌드마다 변경 가능 → `[class*=ModalCaptchaText_captchaInput]` 패턴 사용 권장
7. **시스템 분기 시점**: 대기열 통과 후 URL로 판별 (`/onestop` vs `poticket.interpark.com`)

---

## 8. 우선순위 요약

| 순위 | 항목 | 대상 | 상태 | 난이도 |
|------|------|------|------|--------|
| ~~1~~ | ~~시스템 분기 감지~~ | 공통 | ✅ 완료 | 하 |
| 2 | 카카오 로그인 | 공통 | 미구현 | 중 |
| ~~3~~ | ~~NEW CAPTCHA 처리~~ | NEW | ✅ 완료 | 중 |
| ~~4~~ | ~~NEW SVG 좌석 선택~~ | NEW | ✅ 완료 | 상 |
| ~~5~~ | ~~NEW "선택 완료" 클릭~~ | NEW | ✅ 완료 | 하 |
| ~~6~~ | ~~OLD 구역/좌석 선택~~ | OLD | ✅ 완료 | 중 |
| 7 | Mock 서버 확장 | 테스트 | 미구현 | 중 |
| 8 | UI 확장 | 앱 | 미구현 | 하 |

---

## 9. 최종 요약

```
현재 코드 커버리지:

상품 페이지 ────────────── ✅ 완료 (NEW/OLD 공통)
├── 팝업 닫기              ✅
├── 날짜 선택              ✅ (DOM + Screen)
├── 회차 선택              ✅ (DOM + Screen)
├── 예매 클릭 + 대기열     ✅
│
├── 카카오 로그인          ⚠️ 기존 세션 활용 (자동 로그인 미구현)
│
├── [OLD] CAPTCHA          ✅ (OLD iframe + NEW 모달 모두 대응)
├── [OLD] 구역 선택        ✅ (사용자 대기 후 SelectSeat span 감지)
├── [OLD] 좌석 선택        ✅ (span[onclick*=SelectSeat] 클릭)
│
├── [NEW] CAPTCHA          ✅ (React 모달 셀렉터 대응)
├── [NEW] 좌석 선택        ✅ (SVG circle 스캔 + 클릭)
└── [NEW] 선택 완료        ✅ (Playwright trusted click + 페이지 전환 검증)
```

**핵심 발견**: NEW와 OLD는 상품 페이지는 동일하지만, **예매 시스템이 완전히 다르다**.
- OLD: `poticket.interpark.com` 레거시 iframe 시스템
- NEW: `tickets.interpark.com/onestop` Next.js React SPA

**현재 상태**: NEW 원스탑 + OLD 레거시 모두 좌석 선택 및 완료까지 자동화 구현 완료. 카카오 로그인 자동화만 미구현 (기존 브라우저 세션 활용으로 대체).
