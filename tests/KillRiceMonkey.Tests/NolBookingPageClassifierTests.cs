using KillRiceMonkey.Infrastructure.Services;

namespace KillRiceMonkey.Tests;

public class NolBookingPageClassifierTests
{
    [Fact]
    public void Queue_url_is_not_booking_ready()
    {
        var state = NolBookingPageClassifier.Classify(
            currentUrl: "https://tickets.interpark.com/queue",
            beforeUrl: "https://tickets.interpark.com/goods/25017910",
            currentTitle: "NOL Mock - 대기열",
            beforeTitle: "상품",
            hasProductSide: false,
            hasCaptchaInput: false,
            hasCaptchaModal: false,
            hasLegacySeatFrame: false,
            hasOnestopSeatMap: false);

        Assert.Equal(NolBookingPageState.Queue, state);
    }

    [Fact]
    public void Captcha_modal_on_onestop_page_is_booking_ready()
    {
        var state = NolBookingPageClassifier.Classify(
            currentUrl: "https://tickets.interpark.com/onestop/seat",
            beforeUrl: "https://tickets.interpark.com/goods/25017910",
            currentTitle: "선택 좌석",
            beforeTitle: "상품",
            hasProductSide: false,
            hasCaptchaInput: true,
            hasCaptchaModal: true,
            hasLegacySeatFrame: false,
            hasOnestopSeatMap: true);

        Assert.Equal(NolBookingPageState.BookingReady, state);
    }

    [Fact]
    public void Legacy_ifrmSeat_page_is_booking_ready()
    {
        var state = NolBookingPageClassifier.Classify(
            currentUrl: "https://poticket.interpark.com/Book/BookSession.asp",
            beforeUrl: "https://tickets.interpark.com/goods/25017910",
            currentTitle: "좌석선택",
            beforeTitle: "상품",
            hasProductSide: false,
            hasCaptchaInput: false,
            hasCaptchaModal: false,
            hasLegacySeatFrame: true,
            hasOnestopSeatMap: false);

        Assert.Equal(NolBookingPageState.BookingReady, state);
    }

    [Fact]
    public void Unknown_transition_without_ready_markers_stays_unknown()
    {
        var state = NolBookingPageClassifier.Classify(
            currentUrl: "https://tickets.interpark.com/booking/loading",
            beforeUrl: "https://tickets.interpark.com/goods/25017910",
            currentTitle: "로딩",
            beforeTitle: "상품",
            hasProductSide: false,
            hasCaptchaInput: false,
            hasCaptchaModal: false,
            hasLegacySeatFrame: false,
            hasOnestopSeatMap: false);

        Assert.Equal(NolBookingPageState.Unknown, state);
    }
}
