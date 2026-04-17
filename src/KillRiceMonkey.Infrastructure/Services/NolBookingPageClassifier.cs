namespace KillRiceMonkey.Infrastructure.Services;

internal enum NolBookingPageState
{
    Goods,
    Queue,
    BookingReady,
    Unknown
}

internal static class NolBookingPageClassifier
{
    public static NolBookingPageState Classify(
        string currentUrl,
        string beforeUrl,
        string currentTitle,
        string beforeTitle,
        bool hasProductSide,
        bool hasCaptchaInput,
        bool hasCaptchaModal,
        bool hasLegacySeatFrame,
        bool hasOnestopSeatMap)
    {
        if (IsQueuePage(currentUrl, currentTitle))
            return NolBookingPageState.Queue;

        if (hasCaptchaInput || hasCaptchaModal || hasLegacySeatFrame || hasOnestopSeatMap)
            return NolBookingPageState.BookingReady;

        if (currentUrl.Contains("/onestop", StringComparison.OrdinalIgnoreCase) ||
            currentUrl.Contains("poticket.interpark.com", StringComparison.OrdinalIgnoreCase))
            return NolBookingPageState.BookingReady;

        var urlChanged = !string.Equals(currentUrl, beforeUrl, StringComparison.OrdinalIgnoreCase) &&
                         !string.Equals(StripUrlFragment(currentUrl), StripUrlFragment(beforeUrl), StringComparison.OrdinalIgnoreCase);
        var titleChanged = !string.IsNullOrWhiteSpace(currentTitle) &&
                           !string.Equals(currentTitle, beforeTitle, StringComparison.Ordinal);

        if (!urlChanged && !titleChanged && hasProductSide)
            return NolBookingPageState.Goods;

        return NolBookingPageState.Unknown;
    }

    private static bool IsQueuePage(string currentUrl, string currentTitle)
    {
        if (currentUrl.Contains("/queue", StringComparison.OrdinalIgnoreCase))
            return true;

        return currentTitle.Contains("대기열", StringComparison.OrdinalIgnoreCase) ||
               currentTitle.Contains("queue", StringComparison.OrdinalIgnoreCase);
    }

    private static string StripUrlFragment(string? url)
    {
        if (string.IsNullOrWhiteSpace(url))
            return string.Empty;

        var hashIndex = url.IndexOf('#');
        return hashIndex >= 0 ? url[..hashIndex] : url;
    }
}
