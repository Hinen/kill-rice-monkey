# CDP(Chrome DevTools Protocol) 브라우저 조작 가이드

이 문서는 Chrome remote-debugging을 통해 웹 페이지에 접속하고, DOM 구조를 파악하고, 로그인/클릭 등을 수행하는 **검증된 절차**를 기록한 것이다. 다른 세션에서 이 방식을 그대로 따라하면 된다.

---

## 0. 사전 준비

### Python websockets 설치 (최초 1회)
```bash
pip install websockets
```

### 환경변수 (모든 Python 명령에 적용)
```bash
PYTHONIOENCODING=utf-8
```
Python 3.7(cp949)에서 유니코드 출력 에러를 방지한다. 모든 python 명령 앞에 `PYTHONIOENCODING=utf-8`을 붙인다.

---

## 1. Chrome remote-debugging 실행

```bash
PROFILE_DIR="$LOCALAPPDATA/KillRiceMonkey/MelonRemoteDebugProfile"
mkdir -p "$PROFILE_DIR"
CHROME_PATH="/c/Program Files (x86)/Google/Chrome/Application/chrome.exe"
if [ ! -f "$CHROME_PATH" ]; then
  CHROME_PATH="/c/Program Files/Google/Chrome/Application/chrome.exe"
fi
"$CHROME_PATH" --remote-debugging-port=9222 --user-data-dir="$PROFILE_DIR" --new-window "TARGET_URL" &
sleep 4
curl -s http://127.0.0.1:9222/json/version | head -3
```

- `TARGET_URL`에 접속할 URL을 넣는다.
- 프로필 디렉토리를 지정하면 로그인 세션이 유지된다.
- 포트 9222가 이미 사용 중이면 Chrome이 이미 실행 중이므로 바로 2단계로 간다.

---

## 2. 열린 페이지 목록 확인

```bash
curl -s http://127.0.0.1:9222/json | PYTHONIOENCODING=utf-8 python -c "
import json, sys
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)
pages = json.load(sys.stdin)
for p in pages:
    if p.get('type') == 'page':
        print(f\"ID: {p['id']} | URL: {p.get('url','')[:130]} | Title: {p.get('title','')[:60]}\")
"
```

여기서 얻은 **page ID**를 이후 모든 WebSocket 연결에 사용한다.

---

## 3. 새 탭 열기

```bash
PYTHONIOENCODING=utf-8 python -c "
import asyncio, json, websockets, sys
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

async def main():
    # browser WebSocket URL은 json/version에서 확인 가능
    browser_uri = 'ws://127.0.0.1:9222/devtools/browser/BROWSER_ID'
    async with websockets.connect(browser_uri, max_size=2**22) as ws:
        await ws.send(json.dumps({'id':1,'method':'Target.createTarget','params':{'url':'TARGET_URL'}}))
        resp = json.loads(await ws.recv())
        print(json.dumps(resp, indent=2))

asyncio.run(main())
"
```

- `BROWSER_ID`는 Chrome 실행 시 출력되는 `DevTools listening on ws://127.0.0.1:9222/devtools/browser/XXXXX`에서 얻는다.
- 새 탭의 page ID는 2단계로 다시 확인한다.

---

## 4. 페이지 DOM 구조 파악 (핵심 템플릿)

### 4.1 기본 구조 (안전한 eval 패턴)

**반드시 이 패턴을 사용한다.** Page 이벤트와 eval 응답이 섞이는 문제를 방지한다.

```bash
PYTHONIOENCODING=utf-8 python -c "
import asyncio, json, websockets, sys
sys.stdout = open(sys.stdout.fileno(), mode='w', encoding='utf-8', buffering=1)

async def main():
    uri = 'ws://127.0.0.1:9222/devtools/page/PAGE_ID'
    async with websockets.connect(uri, max_size=2**24) as ws:
        # 1) 대기 중인 이벤트 비우기 (필수)
        while True:
            try: await asyncio.wait_for(ws.recv(), timeout=0.3)
            except: break

        # 2) eval 함수 (이벤트 건너뛰고 응답만 받기)
        async def eval_js(expr, msg_id):
            await ws.send(json.dumps({
                'id': msg_id,
                'method': 'Runtime.evaluate',
                'params': {'expression': expr, 'returnByValue': True}
            }))
            deadline = asyncio.get_event_loop().time() + 5
            while asyncio.get_event_loop().time() < deadline:
                try:
                    msg = await asyncio.wait_for(ws.recv(), timeout=1)
                    d = json.loads(msg)
                    if d.get('id') == msg_id:
                        return d.get('result',{}).get('result',{}).get('value','')
                except asyncio.TimeoutError:
                    break
            return None

        # 3) 여기에 eval_js 호출을 작성
        r = await eval_js('JSON.stringify({url: location.href, title: document.title})', 1)
        print(r)

asyncio.run(main())
"
```

**핵심 포인트**:
- `while True: try/except` 루프로 **대기 이벤트를 반드시 비운 후** eval을 시작한다.
- `eval_js` 함수 내부에서 `d.get('id') == msg_id`로 **응답만 필터링**한다 (Page 이벤트 무시).
- JSON 문자열 내부의 따옴표는 `\"` 또는 `'`로 이스케이프한다.
- `\n`, `\r` 등 특수문자가 포함된 텍스트는 `replace()` 로 정리한다.

### 4.2 DOM 구조 파악용 JS 표현식

**페이지 전체 구조 파악**:
```javascript
JSON.stringify({
    url: location.href,
    title: document.title,
    iframes: [...document.querySelectorAll('iframe')].map(f => ({src:f.src?.substring(0,120), name:f.name, id:f.id})),
    buttons: [...document.querySelectorAll('button')].filter(b => {
        var s = window.getComputedStyle(b); return s.display !== 'none';
    }).map(b => ({text: b.innerText?.trim()?.substring(0,40), cls: b.className?.substring(0,60)})).slice(0,15),
    inputs: [...document.querySelectorAll('input:not([type=hidden])')].map(i => ({
        type:i.type, name:i.name, id:i.id, placeholder:i.placeholder?.substring(0,40)
    })),
    divIds: [...document.querySelectorAll('div[id]')].map(d => d.id).slice(0,25),
    hasReact: !!document.querySelector('#__next, #root, [data-reactroot]'),
})
```

**특정 요소 상세 분석**:
```javascript
JSON.stringify({
    el: !!document.querySelector('SELECTOR'),
    visible: window.getComputedStyle(document.querySelector('SELECTOR'))?.display !== 'none',
    text: document.querySelector('SELECTOR')?.innerText?.trim()?.substring(0,100),
    cls: document.querySelector('SELECTOR')?.className,
    children: [...(document.querySelector('SELECTOR')?.children||[])].map(c => ({
        tag: c.tagName, cls: c.className?.substring(0,50), id: c.id
    })).slice(0,10),
})
```

**iframe 내부 DOM 접근** (같은 origin일 때만 가능):
```javascript
(function(){
    var f = document.querySelector('#IFRAME_ID');
    var doc = f.contentDocument;
    return JSON.stringify({
        url: f.src,
        inputs: [...doc.querySelectorAll('input')].map(i => ({id:i.id, name:i.name, type:i.type})),
        buttons: [...doc.querySelectorAll('button,a[onclick]')].map(b => ({
            text:(b.innerText||'').substring(0,30),
            onclick: b.getAttribute('onclick')?.substring(0,80)
        })).slice(0,10),
    });
})()
```

---

## 5. 클릭 수행

### 5.1 일반 JS 클릭 (비-React 요소)

```javascript
document.querySelector('SELECTOR').click(); 'clicked'
```

### 5.2 CDP Trusted Click (React 요소, 필수)

React 이벤트 핸들러가 붙은 버튼은 JS `.click()`이 동작하지 않는다. **반드시 CDP Input.dispatchMouseEvent를 사용해야 한다.**

```python
# 1단계: 버튼 좌표 구하기
pos_str = await eval_js('''
    var btn = document.querySelector("SELECTOR");
    var r = btn.getBoundingClientRect();
    JSON.stringify({x: r.x + r.width/2, y: r.y + r.height/2})
''', MSG_ID)
pos = json.loads(pos_str)

# 2단계: CDP trusted click
await ws.send(json.dumps({'id': ID1, 'method': 'Input.dispatchMouseEvent',
    'params': {'type':'mousePressed','x':pos['x'],'y':pos['y'],'button':'left','clickCount':1}}))
await ws.recv()
await ws.send(json.dumps({'id': ID2, 'method': 'Input.dispatchMouseEvent',
    'params': {'type':'mouseReleased','x':pos['x'],'y':pos['y'],'button':'left','clickCount':1}}))
await ws.recv()
```

---

## 6. 입력 필드에 값 넣기

### 6.1 React controlled input (nativeInputValueSetter)

React가 value를 관리하는 input에 직접 `.value =`을 하면 React가 감지하지 못한다. **반드시 이 패턴을 사용한다:**

```javascript
var input = document.querySelector('SELECTOR');
input.focus();
var nativeInputValueSetter = Object.getOwnPropertyDescriptor(
    window.HTMLInputElement.prototype, 'value').set;
nativeInputValueSetter.call(input, 'VALUE');
input.dispatchEvent(new Event('input', {bubbles: true}));
input.dispatchEvent(new Event('change', {bubbles: true}));
'filled'
```

### 6.2 일반 input

```javascript
document.querySelector('SELECTOR').value = 'VALUE';
document.querySelector('SELECTOR').dispatchEvent(new Event('input', {bubbles: true}));
'filled'
```

---

## 7. 페이지 이동/팝업 감지

### 7.1 클릭 후 네비게이션 모니터링

```python
await ws.send(json.dumps({'id': ID, 'method': 'Page.enable', 'params': {}}))
# ... (drain events) ...

# 클릭 수행 후 이벤트 모니터링
deadline = asyncio.get_event_loop().time() + TIMEOUT_SECONDS
while asyncio.get_event_loop().time() < deadline:
    try:
        msg = await asyncio.wait_for(ws.recv(), timeout=0.5)
        d = json.loads(msg)
        m = d.get('method', '')
        if m == 'Page.frameNavigated':
            fid = d['params']['frame']['id']
            furl = d['params']['frame']['url']
            if fid == 'PAGE_ID':  # 메인 프레임만
                print(f'NAVIGATE: {furl[:150]}')
        elif m == 'Page.windowOpen':
            print(f'NEW WINDOW: {json.dumps(d["params"])[:200]}')
        elif m == 'Page.javascriptDialogOpening':
            print(f'DIALOG: {d["params"].get("message","")[:100]}')
            # 자동 수락
            await ws.send(json.dumps({'id':99,'method':'Page.handleJavaScriptDialog','params':{'accept':True}}))
    except asyncio.TimeoutError:
        continue
```

### 7.2 새 창(팝업) 감지 후 조작

클릭 후 새 창이 열리면:
1. `curl -s http://127.0.0.1:9222/json`으로 새 page ID 확인
2. 새 page ID로 별도 WebSocket 연결
3. 새 창에서 DOM 조작 수행

---

## 8. 카카오 로그인 전체 예시

```python
# 1. 로그인 페이지에서 "카카오로 시작하기" 클릭 (CDP trusted click)
pos_str = await eval_js('''
    var btns = [...document.querySelectorAll('button')];
    var kakaoBtn = btns.find(b => b.innerText.includes('카카오'));
    kakaoBtn.scrollIntoView({block:'center'});
    var rect = kakaoBtn.getBoundingClientRect();
    JSON.stringify({x: rect.x + rect.width/2, y: rect.y + rect.height/2})
''', 10)
# ... CDP trusted click ...

# 2. 새 창(accounts.kakao.com) page ID 확인 (curl json)
# 3. 새 WebSocket 연결

# 4. ID 입력 (nativeInputValueSetter)
await eval_js('''
    var idInput = document.querySelector('#loginId--1');
    idInput.focus();
    var setter = Object.getOwnPropertyDescriptor(
        window.HTMLInputElement.prototype, 'value').set;
    setter.call(idInput, 'KAKAO_ID');
    idInput.dispatchEvent(new Event('input', {bubbles: true}));
    idInput.dispatchEvent(new Event('change', {bubbles: true}));
    'done'
''', 20)

# 5. PW 입력 (동일 패턴, #password--2)
# 6. 로그인 버튼 클릭 (CDP trusted click, button[type=submit] "로그인")
# 7. 팝업 닫히면 WebSocket 끊어짐 (정상) → 메인 페이지 URL 확인
```

---

## 9. 흔한 오류와 해결

| 오류 | 원인 | 해결 |
|------|------|------|
| `KeyError: 'result'` 또는 `KeyError: 'value'` | Page 이벤트가 eval 응답보다 먼저 도착 | **4.1의 eval_js 패턴 사용** (id로 필터링) |
| `UnicodeEncodeError: 'cp949'` | Python stdout이 cp949 | `PYTHONIOENCODING=utf-8` + `sys.stdout = open(...)` |
| `ConnectionClosedError: no close frame` | 페이지가 navigate하거나 팝업이 닫힘 | 정상 동작. try/except로 감싸기 |
| JS `.click()` 후 아무 반응 없음 | React synthetic event handler | **CDP trusted click 사용** (5.2) |
| input에 값 입력 후 React가 인식 못함 | React controlled input | **nativeInputValueSetter 사용** (6.1) |
| `SyntaxError: Invalid regular expression` | JS 내 `\n`, 정규식 특수문자 | JS 문자열에서 `replace()` 대신 `substring()` 사용 |
| eval 결과가 빈 문자열/None | JS 표현식이 undefined 반환 | 마지막 줄에 명시적 반환값 추가 (예: `'done'`) |
| iframe contentDocument가 null | cross-origin iframe | `page.Frames`로 접근하거나, CDP Target으로 접근 |

---

## 10. 체크리스트

페이지 조사 시 반드시 확인할 항목:

- [ ] URL, title
- [ ] iframe 유무 및 src/name/id
- [ ] React 여부 (`#__next`, `__reactEventHandlers`)
- [ ] 주요 버튼 (text, class, onclick, disabled 상태)
- [ ] input 필드 (id, name, type, placeholder)
- [ ] div[id] 목록 (페이지 구조 파악)
- [ ] SVG/canvas 유무 (좌석맵 등)
- [ ] 팝업/모달 (`.popup.is-visible`, `[class*=Modal]`)
- [ ] 로그인 상태 (로그인 링크 존재 여부)
- [ ] CSS module 클래스명 패턴 (`ClassName_property__hash`)
