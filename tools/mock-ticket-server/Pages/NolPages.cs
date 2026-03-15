namespace MockTicketServer.Pages;

public static class NolPages
{
    public static string GoodsPage(int queueSeconds) => $$"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>NOL Mock - 공연 상품</title>
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0;
            min-height: 100vh;
            font-family: "Segoe UI", sans-serif;
            background: linear-gradient(135deg, #151922 0%, #222b3b 100%);
            color: #fff;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }

        #productSide {
            width: min(100%, 420px);
            background: rgba(12, 16, 24, 0.92);
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 20px;
            padding: 24px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.35);
        }

        h1 {
            margin: 0 0 8px;
            font-size: 28px;
        }

        .subtitle {
            margin: 0 0 24px;
            color: #c0cadb;
            font-size: 14px;
        }

        .sideCalendar ul,
        .sideTimeTable {
            list-style: none;
            padding: 0;
            margin: 0;
        }

        .sideCalendar > ul:first-child {
            display: grid;
            grid-template-columns: 44px 1fr 44px;
            gap: 8px;
            align-items: center;
            margin-bottom: 16px;
        }

        .sideCalendar > ul:first-child li {
            background: rgba(255, 255, 255, 0.08);
            border-radius: 10px;
            min-height: 40px;
            display: flex;
            align-items: center;
            justify-content: center;
            cursor: pointer;
            transition: background 0.2s ease;
        }

        .sideCalendar > ul:first-child li:hover {
            background: rgba(255, 255, 255, 0.16);
        }

        .sideCalendar > ul:first-child li[data-view="month current"] {
            cursor: default;
            font-weight: 700;
            background: rgba(255, 255, 255, 0.12);
        }

        ul[data-view="days"] {
            display: grid;
            grid-template-columns: repeat(5, minmax(0, 1fr));
            gap: 10px;
        }

        ul[data-view="days"] li,
        .timeTableLabel,
        .sideBtn {
            transition: transform 0.2s ease, background 0.2s ease, border-color 0.2s ease;
        }

        ul[data-view="days"] li {
            min-height: 48px;
            display: flex;
            align-items: center;
            justify-content: center;
            border-radius: 12px;
            background: rgba(255, 255, 255, 0.06);
            border: 1px solid transparent;
            cursor: pointer;
        }

        ul[data-view="days"] li:hover,
        .timeTableLabel:hover,
        .sideBtn:hover {
            transform: translateY(-1px);
        }

        ul[data-view="days"] li:hover {
            background: rgba(255, 255, 255, 0.12);
        }

        ul[data-view="days"] li.is-selected {
            background: #4aa3ff;
            border-color: #80c2ff;
            color: #081019;
            font-weight: 700;
        }

        .containerTop,
        .containerMiddle {
            margin-top: 18px;
            padding: 14px 16px;
            border-radius: 14px;
            background: rgba(255, 255, 255, 0.06);
        }

        .selectedData {
            display: flex;
            align-items: center;
            min-height: 24px;
            font-weight: 600;
        }

        .sideTimeTable {
            display: grid;
            gap: 10px;
            margin-top: 18px;
        }

        .timeTableLabel {
            border-radius: 12px;
            padding: 14px 16px;
            background: rgba(255, 255, 255, 0.08);
            border: 1px solid rgba(255, 255, 255, 0.1);
            cursor: pointer;
        }

        .timeTableLabel.is-toggled {
            background: #ffd76a;
            border-color: #ffe7a4;
            color: #1a1505;
            font-weight: 700;
        }

        .sideBtn {
            display: block;
            margin-top: 22px;
            text-align: center;
            text-decoration: none;
            border-radius: 14px;
            padding: 16px;
            background: #17c964;
            color: #08150d;
            font-weight: 800;
        }

        .sideBtn:hover {
            background: #3ddc84;
        }
    </style>
</head>
<body>
    <div id="productSide">
        <h1>NOL Mock</h1>
        <p class="subtitle">공연 상품 페이지 · 대기열 {{queueSeconds}}초</p>

        <div class="sideCalendar">
            <ul>
                <li data-view="month prev">&lt;</li>
                <li data-view="month current">2026.04</li>
                <li data-view="month next">&gt;</li>
            </ul>
            <ul data-view="days">
                <li>1</li><li>2</li><li>3</li><li>4</li><li>5</li>
                <li>6</li><li>7</li><li>8</li><li>9</li><li>10</li>
                <li>11</li><li>12</li><li>13</li><li>14</li><li>15</li>
                <li>16</li><li>17</li><li>18</li><li>19</li><li>20</li>
                <li>21</li><li>22</li><li>23</li><li>24</li><li>25</li>
                <li>26</li><li>27</li><li>28</li><li>29</li><li>30</li>
            </ul>
        </div>

        <div class="containerTop">
            <div class="selectedData">
                <span class="date"></span>
            </div>
        </div>

        <div class="sideTimeTable">
            <div class="timeTableLabel" role="button">1회 19:00</div>
            <div class="timeTableLabel" role="button">2회 21:00</div>
        </div>

        <div class="containerMiddle">
            <div class="selectedData">
                <span class="time"></span>
            </div>
        </div>

        <a class="sideBtn is-primary" href="javascript:void(0)">예매하기</a>
    </div>

    <script>
        const dayItems = document.querySelectorAll('ul[data-view="days"] li');
        const dateOutput = document.querySelector('.containerTop .selectedData .date');
        const timeButtons = document.querySelectorAll('.timeTableLabel');
        const timeOutput = document.querySelector('.containerMiddle .selectedData .time');
        const bookingButton = document.querySelector('.sideBtn');

        dayItems.forEach(function (item) {
            item.addEventListener('click', function () {
                dayItems.forEach(function (candidate) {
                    candidate.classList.remove('is-selected');
                });

                item.classList.add('is-selected');
                const day = String(item.textContent.trim()).padStart(2, '0');
                dateOutput.textContent = '2026.04.' + day;
            });
        });

        timeButtons.forEach(function (button) {
            button.addEventListener('click', function () {
                timeButtons.forEach(function (candidate) {
                    candidate.classList.remove('is-toggled');
                });

                button.classList.add('is-toggled');
                timeOutput.textContent = button.textContent.trim();
            });
        });

        bookingButton.addEventListener('click', function () {
            window.location.href = '/queue';
        });
    </script>
</body>
</html>
""";

    public static string QueuePage(int queueSeconds) => $$"""
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>NOL Mock - 대기열</title>
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0;
            min-height: 100vh;
            font-family: "Segoe UI", sans-serif;
            background: radial-gradient(circle at top, #27405f 0%, #0d1420 58%, #090d15 100%);
            color: #fff;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }

        .queue-shell {
            width: min(100%, 520px);
            text-align: center;
            padding: 36px 28px;
            border-radius: 24px;
            background: rgba(7, 11, 18, 0.82);
            border: 1px solid rgba(255, 255, 255, 0.12);
            box-shadow: 0 24px 70px rgba(0, 0, 0, 0.35);
        }

        h1 {
            margin: 0 0 12px;
            font-size: 32px;
        }

        .message {
            margin: 0;
            color: #d1ddf3;
            font-size: 20px;
        }

        .countdown {
            display: block;
            margin-top: 20px;
            font-size: clamp(56px, 12vw, 92px);
            font-weight: 800;
            line-height: 1;
            color: #7bc4ff;
        }
    </style>
</head>
<body>
    <main class="queue-shell">
        <h1>예매 대기열</h1>
        <p class="message">대기 중... 남은 시간: <span id="countdownText">{{queueSeconds}}</span>초</p>
        <span class="countdown" id="countdownNumber">{{queueSeconds}}</span>
    </main>

    <script>
        let remainingSeconds = {{queueSeconds}};
        const countdownText = document.getElementById('countdownText');
        const countdownNumber = document.getElementById('countdownNumber');

        function updateCountdown() {
            countdownText.textContent = String(remainingSeconds);
            countdownNumber.textContent = String(remainingSeconds);
        }

        updateCountdown();

        const timer = window.setInterval(function () {
            remainingSeconds -= 1;

            if (remainingSeconds <= 0) {
                window.clearInterval(timer);
                countdownText.textContent = '0';
                countdownNumber.textContent = '0';
                window.location.href = '/captcha';
                return;
            }

            updateCountdown();
        }, 1000);
    </script>
</body>
</html>
""";

    public static string CaptchaPage() => """
<!DOCTYPE html>
<html lang="ko">
<head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>NOL Mock - 예매</title>
    <style>
        * { box-sizing: border-box; }
        body {
            margin: 0;
            min-height: 100vh;
            font-family: "Segoe UI", sans-serif;
            background: linear-gradient(160deg, #eef4fb 0%, #dbe7f5 100%);
            color: #162033;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 24px;
        }

        .booking-shell {
            width: min(100%, 420px);
            background: rgba(255, 255, 255, 0.92);
            border-radius: 24px;
            padding: 32px;
            box-shadow: 0 22px 60px rgba(31, 55, 90, 0.18);
            text-align: center;
        }

        h1 {
            margin: 0 0 8px;
            font-size: 30px;
        }

        p {
            margin: 0 0 20px;
            color: #52627a;
        }

        #imgCaptcha {
            display: block;
            width: 200px;
            height: 60px;
            margin: 0 auto 16px;
            border: 1px solid #cdd8e8;
            border-radius: 12px;
            background: #f3f7fc;
        }

        #txtCaptcha {
            width: 100%;
            padding: 14px 16px;
            border-radius: 12px;
            border: 1px solid #c6d3e4;
            font-size: 16px;
            margin-bottom: 12px;
        }

        .actions {
            display: grid;
            grid-template-columns: 1fr 1fr;
            gap: 10px;
        }

        button {
            border: none;
            border-radius: 12px;
            padding: 14px 16px;
            font-size: 15px;
            font-weight: 700;
            cursor: pointer;
        }

        button[type="submit"] {
            background: #206cff;
            color: #fff;
        }

        #btnReload {
            background: #e8eef7;
            color: #234063;
        }

        #resultMessage {
            margin-top: 18px;
            min-height: 24px;
            font-weight: 700;
            color: #17804b;
        }
    </style>
</head>
<body>
    <main class="booking-shell">
        <h1>보안문자 확인</h1>
        <p>보안문자를 입력하고 예매를 완료하세요.</p>
        <form id="captchaForm">
            <img id="imgCaptcha" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAMgAAAA8CAIAAACsOWLGAAAClElEQVR4nO3cQW6jQBRF0VRx/5e5s0m2W0Q8JXEh0fZrlmjlI4D8a7xwyWOOj66fHwC3M1cHwM8iLBAsECwQLBAsECwQLBAsECwQLBAsECwQLBAsECwQLBAsECwQLHjL+S1J7bI3f2Zr4O87DsP7M2l3PH7R9vZc7wOt7YvTe3N1vtt2N8+L12f8GrM+qjz6Q7tTnF8qRr7x7jXtlG7b8iV0bSx+vFn2fXzvw9d90Mx3f2sNp1f5S7wdc89v2n9hd3l2U6Zf7h4b3Gx7LJwXlP5BszfV48sL6N8cN2+Kx6zQzH6uF0ztuM2xvLPdX5R8F6+X1g+YmWmLwA4Hy+bt8y+dS5bP0sz/V7bqf5v0Vv2bC6m5H3r7A+1D4lWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBY8GL90l0n2f3Tg4lfM2Pw3r6Pq/LPZ6dY3mWkYb8fJcY1m1Z5f7Qf8n4q1vXwK5cVle3n4GvB35h5H5aR5qT6zS3m3jG2W8wKfY4v6zF3k9eN6mPXwV7W9t7NqjR7gOMZr3Hx5P2cB3xj2d5m6V+U3j5hWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBYIFggWCBY8CP9AT/lcS4I1k4tAAAAAElFTkSuQmCC" width="200" height="60" />
            <input id="txtCaptcha" type="text" placeholder="보안문자 입력" />
            <div class="actions">
                <button type="submit">입력완료</button>
                <button type="button" id="btnReload">새로고침</button>
            </div>
        </form>
        <div id="resultMessage"></div>
    </main>

    <script>
        const captchaImage = document.getElementById('imgCaptcha');
        const captchaForm = document.getElementById('captchaForm');
        const captchaInput = document.getElementById('txtCaptcha');
        const resultMessage = document.getElementById('resultMessage');
        const captchaCodes = ['A1B2C3', 'C3D4E5', 'M7N8P9', 'Q2R4T6'];
        let captchaIndex = 0;

        function buildCaptchaDataUrl(code) {
            const canvas = document.createElement('canvas');
            canvas.width = 200;
            canvas.height = 60;

            const context = canvas.getContext('2d');
            context.fillStyle = '#eef4ff';
            context.fillRect(0, 0, canvas.width, canvas.height);
            context.fillStyle = '#d8e6ff';
            context.fillRect(8, 8, 184, 44);
            context.strokeStyle = '#7aa6ff';
            context.lineWidth = 2;
            context.strokeRect(8, 8, 184, 44);
            context.fillStyle = '#1d3354';
            context.font = 'bold 28px Segoe UI';
            context.textAlign = 'center';
            context.textBaseline = 'middle';
            context.fillText(code, 100, 32);

            return canvas.toDataURL('image/png');
        }

        function fnCapchaRefresh() {
            captchaIndex = (captchaIndex + 1) % captchaCodes.length;
            const nextCode = captchaCodes[captchaIndex];
            captchaImage.src = buildCaptchaDataUrl(nextCode) + '#ts=' + Date.now();
            captchaInput.value = '';
            resultMessage.textContent = '';
        }

        document.getElementById('btnReload').addEventListener('click', fnCapchaRefresh);

        captchaForm.addEventListener('submit', function (event) {
            event.preventDefault();
            resultMessage.textContent = '예매 완료!';
        });
    </script>
</body>
</html>
""";
}
