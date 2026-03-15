namespace MockTicketServer.Pages;

public static class MelonPages
{
    public static string PerformancePage(int queueSeconds) => $$"""
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 공연 정보</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", sans-serif; background: #f4f6f8; color: #1f2933; }
    main { max-width: 720px; margin: 48px auto; padding: 32px; background: #ffffff; border-radius: 20px; box-shadow: 0 18px 48px rgba(15, 23, 42, 0.12); }
    h1 { margin: 0 0 12px; font-size: 32px; }
    p { margin: 0 0 28px; color: #52606d; }
    h2 { margin: 0 0 12px; font-size: 18px; }
    .section { margin-bottom: 28px; }
    .option-list { display: flex; gap: 12px; flex-wrap: wrap; }
    .item_date, .item_time { min-width: 120px; padding: 14px 18px; border: 1px solid #d9e2ec; border-radius: 14px; background: #f8fafc; cursor: pointer; text-align: center; font-weight: 600; transition: all 0.15s ease; }
    .item_date.on, .item_time.on { background: #00c73c; border-color: #00c73c; color: #ffffff; box-shadow: 0 10px 20px rgba(0, 199, 60, 0.22); }
    #ticketReservation_Btn { width: 100%; padding: 16px 20px; border: none; border-radius: 16px; background: #1f8f4d; color: #ffffff; font-size: 18px; font-weight: 700; cursor: pointer; }
  </style>
</head>
<body>
  <main>
    <h1>Melon Mock - 공연 정보</h1>
    <p>대기열 {{queueSeconds}}초 후 예매 팝업으로 이동하는 멜론 티켓 테스트 페이지입니다.</p>

    <section class="section">
      <h2>관람일 선택</h2>
      <div id="list_date" class="option-list">
        <div class="item_date" data-perfday="20260410">04.10(금)</div>
        <div class="item_date" data-perfday="20260411">04.11(토)</div>
        <div class="item_date" data-perfday="20260412">04.12(일)</div>
      </div>
    </section>

    <section class="section">
      <h2>회차 선택</h2>
      <div id="list_time" class="option-list">
        <div class="item_time">18:00</div>
        <div class="item_time">20:00</div>
      </div>
    </section>

    <button id="ticketReservation_Btn" type="button">예매하기</button>
  </main>

  <script>
    const dateItems = Array.from(document.querySelectorAll('#list_date .item_date'));
    const timeItems = Array.from(document.querySelectorAll('#list_time .item_time'));

    dateItems.forEach((item) => {
      item.addEventListener('click', () => {
        dateItems.forEach((dateItem) => dateItem.classList.remove('on'));
        item.classList.add('on');
      });
    });

    timeItems.forEach((item) => {
      item.addEventListener('click', () => {
        timeItems.forEach((timeItem) => timeItem.classList.remove('on'));
        item.classList.add('on');
      });
    });

    document.getElementById('ticketReservation_Btn').addEventListener('click', () => {
      window.open('/queue/popup', 'melonPopup', 'width=800,height=600');
    });
  </script>
</body>
</html>
""";

    public static string QueuePopup(int queueSeconds) => $$"""
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 대기열</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; min-height: 100vh; display: grid; place-items: center; font-family: "Segoe UI", sans-serif; background: linear-gradient(135deg, #f1f5f9, #d9fbe6); color: #102a43; }
    .panel { width: min(420px, calc(100vw - 32px)); padding: 40px 28px; border-radius: 24px; background: rgba(255, 255, 255, 0.95); text-align: center; box-shadow: 0 20px 45px rgba(15, 23, 42, 0.12); }
    h1 { margin: 0 0 12px; font-size: 28px; }
    p { margin: 0; color: #486581; font-size: 18px; font-weight: 600; }
  </style>
</head>
<body>
  <div class="panel">
    <h1>Melon Mock - 대기열</h1>
    <p id="countdown">대기 중... {{queueSeconds}}초 남음</p>
  </div>

  <script>
    let remainingSeconds = {{queueSeconds}};
    const countdown = document.getElementById('countdown');

    const timer = window.setInterval(() => {
      remainingSeconds -= 1;

      if (remainingSeconds <= 0) {
        countdown.textContent = '대기 중... 0초 남음';
        window.clearInterval(timer);
        window.location.href = '/reservation/popup/onestop.htm';
        return;
      }

      countdown.textContent = `대기 중... ${remainingSeconds}초 남음`;
    }, 1000);
  </script>
</body>
</html>
""";

    public static string OnestopPopup(bool hasCaptcha) => hasCaptcha ? OnestopWithCaptcha() : OnestopNoCaptcha();

    private static string OnestopNoCaptcha() => """
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 예매 팝업</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", sans-serif; background: #eef2f6; color: #102a43; }
    .layout { max-width: 980px; margin: 0 auto; padding: 24px; }
    .card { background: #ffffff; border-radius: 20px; box-shadow: 0 16px 36px rgba(15, 23, 42, 0.12); padding: 24px; }
    h1 { margin: 0 0 16px; font-size: 28px; }
    #seatFrame { width: 100%; height: 500px; border: none; border-radius: 18px; background: #f8fafc; }
  </style>
</head>
<body>
  <div class="layout">
    <div class="card">
      <h1>Melon Mock - 예매 팝업 (캡차 없음)</h1>
      <iframe id="seatFrame" src="/reservation/popup/stepSeat.htm" style="width:100%;height:500px;border:none;"></iframe>
    </div>
  </div>
  <script>window.__melonAlertDetected = false;</script>
</body>
</html>
""";

    private static string OnestopWithCaptcha() => """
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 예매 팝업</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", sans-serif; background: #eef2f6; color: #102a43; }
    .layout { max-width: 980px; margin: 0 auto; padding: 24px; }
    .card { background: #ffffff; border-radius: 20px; box-shadow: 0 16px 36px rgba(15, 23, 42, 0.12); padding: 24px; }
    h1 { margin: 0 0 16px; font-size: 28px; }
    #divRecaptcha { display: grid; grid-template-columns: 1fr auto; gap: 14px 18px; align-items: center; margin-bottom: 18px; }
    #txtCaptcha { width: 100%; padding: 12px 14px; border: 1px solid #cbd2d9; border-radius: 12px; font-size: 16px; }
    #imgCaptcha { display: block; border-radius: 12px; border: 1px solid #bcccdc; background: #f8fafc; }
    .capchaBtns a { color: #1f8f4d; text-decoration: none; font-weight: 600; }
    button { padding: 12px 18px; border: none; border-radius: 12px; background: #00c73c; color: #ffffff; font-weight: 700; cursor: pointer; }
    #captchaStatus { display: none; margin: 0 0 18px; color: #1f8f4d; font-weight: 700; }
    #seatFrame { width: 100%; height: 500px; border: none; border-radius: 18px; background: #f8fafc; }
    @media (max-width: 700px) {
      #divRecaptcha { grid-template-columns: 1fr; }
    }
  </style>
</head>
<body>
  <div class="layout">
    <div class="card">
      <h1>Melon Mock - 예매 팝업</h1>

      <div id="divRecaptcha">
        <input id="txtCaptcha" type="text" placeholder="보안문자 입력" maxlength="6" />
        <img id="imgCaptcha" width="200" height="60" alt="captcha" />
        <div class="capchaBtns">
          <a href="javascript:void(0)" onclick="fnCapchaRefresh()">새로고침</a>
        </div>
        <button id="captchaSubmitButton" type="button">입력완료</button>
      </div>

      <div id="captchaStatus">인증 완료</div>

      <iframe id="seatFrame" src="/reservation/popup/stepSeat.htm" style="width:100%;height:500px;border:none;"></iframe>
    </div>
  </div>

  <script>
    window.__melonAlertDetected = false;

    const captchaBase64 = 'PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIyMDAiIGhlaWdodD0iNjAiIHZpZXdCb3g9IjAgMCAyMDAgNjAiPjxyZWN0IHdpZHRoPSIyMDAiIGhlaWdodD0iNjAiIHJ4PSIxMiIgZmlsbD0iI2Y4ZmFmYyIvPjx0ZXh0IHg9IjI2IiB5PSIzOCIgc3R5bGU9ImZvbnQ6IGJvbGQgMzBweCBzYW5zLXNlcmlmOyBmaWxsOiAjMTAyYTQzOyBsZXR0ZXItc3BhY2luZzogNnB4OyI+QTFCMkMzPC90ZXh0Pjwvc3ZnPg==';

    function fnCapchaRefresh() {
      document.getElementById('imgCaptcha').src = 'data:image/svg+xml;base64,' + captchaBase64;
    }

    document.getElementById('captchaSubmitButton').addEventListener('click', () => {
      window.__melonAlertDetected = false;
      document.getElementById('divRecaptcha').style.display = 'none';
      document.getElementById('captchaStatus').style.display = 'block';
    });

    fnCapchaRefresh();
  </script>
</body>
</html>
""";

    public static string SeatFrame(bool hasZone) => hasZone ? SeatFrameWithZone() : SeatFrameNoZone();

    private static string SeatFrameNoZone() => """
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 좌석 선택 (구역 없음)</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", sans-serif; background: #ffffff; color: #102a43; }
    .wrap { padding: 20px; text-align: center; }
    svg { width: min(100%, 620px); height: 430px; border: 1px solid #d9e2ec; border-radius: 18px; background: #f8fafc; margin: 0 auto 16px; display: block; }
    #partSeatSelected { width: min(100%, 620px); margin: 0 auto 16px; padding: 16px; border: 1px solid #d9e2ec; border-radius: 16px; text-align: left; background: #f8fafc; }
    #partSeatSelected ul { margin: 0; padding-left: 20px; min-height: 24px; }
    #nextTicketSelection { padding: 12px 20px; border: none; border-radius: 12px; background: #00c73c; color: #ffffff; font-weight: 700; cursor: pointer; }
    #seatCompleteMessage { margin-top: 14px; color: #1f8f4d; font-weight: 700; }
  </style>
</head>
<body>
  <div class="wrap">
    <svg id="ez_canvas" viewBox="0 0 560 430" xmlns="http://www.w3.org/2000/svg">
    </svg>

    <div id="partSeatSelected">
      <ul></ul>
    </div>

    <button id="nextTicketSelection" type="button" style="display:none;">좌석 선택 완료</button>
    <div id="seatCompleteMessage"></div>
  </div>

  <script>
    window.__melonAlertDetected = false;

    const svgNamespace = 'http://www.w3.org/2000/svg';
    const canvas = document.getElementById('ez_canvas');
    const selectedSeatList = document.querySelector('#partSeatSelected ul');
    const completeButton = document.getElementById('nextTicketSelection');
    const completeMessage = document.getElementById('seatCompleteMessage');

    for (let row = 0; row < 8; row++) {
      for (let col = 0; col < 10; col++) {
        const rect = document.createElementNS(svgNamespace, 'rect');
        rect.setAttribute('x', String(50 + col * 20));
        rect.setAttribute('y', String(50 + row * 20));
        rect.setAttribute('width', '12');
        rect.setAttribute('height', '12');
        rect.setAttribute('fill', '#4488CC');
        rect.addEventListener('click', function() {
          if (this.dataset.selected === 'true') return;
          this.dataset.selected = 'true';
          this.setAttribute('fill', '#FFD700');
          const item = document.createElement('li');
          item.textContent = 'A석 ' + (row + 1) + '열 ' + (col + 1) + '번';
          selectedSeatList.appendChild(item);
          completeButton.style.display = 'inline-block';
          window.__melonAlertDetected = false;
        });
        canvas.appendChild(rect);
      }
    }

    completeButton.addEventListener('click', function() {
      completeMessage.textContent = '좌석 선택 완료!';
    });
  </script>
</body>
</html>
""";

    private static string SeatFrameWithZone() => """
<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Melon Mock - 좌석 선택</title>
  <style>
    * { box-sizing: border-box; }
    body { margin: 0; font-family: "Segoe UI", sans-serif; background: #ffffff; color: #102a43; }
    .wrap { padding: 20px; text-align: center; }
    svg { width: min(100%, 620px); height: 430px; border: 1px solid #d9e2ec; border-radius: 18px; background: #f8fafc; margin: 0 auto 16px; display: block; }
    #txtSelectSeatInfo { margin-bottom: 16px; color: #486581; font-weight: 600; }
    #partSeatSelected { width: min(100%, 620px); margin: 0 auto 16px; padding: 16px; border: 1px solid #d9e2ec; border-radius: 16px; text-align: left; background: #f8fafc; }
    #partSeatSelected ul { margin: 0; padding-left: 20px; min-height: 24px; }
    #nextTicketSelection { padding: 12px 20px; border: none; border-radius: 12px; background: #00c73c; color: #ffffff; font-weight: 700; cursor: pointer; }
    #seatCompleteMessage { margin-top: 14px; color: #1f8f4d; font-weight: 700; }
    #ez_canvas_zone rect { cursor: pointer; }
  </style>
</head>
<body>
  <div class="wrap">
    <svg id="ez_canvas" viewBox="0 0 560 430" xmlns="http://www.w3.org/2000/svg">
      <g id="ez_canvas_zone">
        <rect x="50" y="50" width="200" height="150" fill="#4488CC" data-zone="A"></rect>
        <rect x="300" y="50" width="200" height="150" fill="#CC4444" data-zone="B"></rect>
        <rect x="175" y="250" width="200" height="150" fill="#44CC44" data-zone="C"></rect>
      </g>
    </svg>

    <div id="txtSelectSeatInfo">구역을 먼저 선택해 주세요.</div>

    <div id="partSeatSelected">
      <ul></ul>
    </div>

    <button id="nextTicketSelection" type="button" style="display:none;">좌석 선택 완료</button>
    <div id="seatCompleteMessage"></div>
  </div>

  <script>
    window.__melonAlertDetected = false;

    const svgNamespace = 'http://www.w3.org/2000/svg';
    const canvas = document.getElementById('ez_canvas');
    const zoneGroup = document.getElementById('ez_canvas_zone');
    const info = document.getElementById('txtSelectSeatInfo');
    const selectedSeatList = document.querySelector('#partSeatSelected ul');
    const completeButton = document.getElementById('nextTicketSelection');
    const completeMessage = document.getElementById('seatCompleteMessage');

    function createSeat(row, col, zone) {
      const rect = document.createElementNS(svgNamespace, 'rect');
      rect.setAttribute('x', String(50 + col * 20));
      rect.setAttribute('y', String(50 + row * 20));
      rect.setAttribute('width', '12');
      rect.setAttribute('height', '12');
      rect.setAttribute('fill', '#4488CC');
      rect.setAttribute('data-seat-zone', zone);
      rect.setAttribute('data-seat-row', String(row + 1));
      rect.setAttribute('data-seat-number', String(col + 1));

      rect.addEventListener('click', () => {
        if (rect.dataset.selected === 'true') {
          return;
        }

        rect.dataset.selected = 'true';
        rect.setAttribute('fill', '#FFD700');
        const item = document.createElement('li');
        item.textContent = zone + '석 ' + (row + 1) + '열 ' + (col + 1) + '번';
        selectedSeatList.appendChild(item);
        completeButton.style.display = 'inline-block';
        window.__melonAlertDetected = false;
      });

      canvas.appendChild(rect);
    }

    function renderSeats(zone) {
      const currentZoneGroup = document.getElementById('ez_canvas_zone');
      if (currentZoneGroup) {
        currentZoneGroup.remove();
      }

      info.style.display = 'none';

      for (let row = 0; row < 8; row += 1) {
        for (let col = 0; col < 10; col += 1) {
          createSeat(row, col, zone);
        }
      }
    }

    zoneGroup.querySelectorAll('rect').forEach((zoneRect) => {
      zoneRect.addEventListener('click', () => {
        renderSeats(zoneRect.dataset.zone || 'A');
      });
    });

    completeButton.addEventListener('click', () => {
      completeMessage.textContent = '좌석 선택 완료!';
    });
  </script>
</body>
</html>
""";
}
