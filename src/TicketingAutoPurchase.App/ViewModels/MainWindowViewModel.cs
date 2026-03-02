using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using TicketingAutoPurchase.Application.Abstractions;
using TicketingAutoPurchase.Application.Models;

namespace TicketingAutoPurchase.App.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private readonly ITicketingAutomationService _ticketingAutomationService;

    [ObservableProperty]
    private string _eventKeyword = "sample concert";

    [ObservableProperty]
    private string _targetUrl = "https://example.com/ticket";

    [ObservableProperty]
    private string _statusMessage = "준비 완료";

    [ObservableProperty]
    private string _lastRunSummary = "자동화 실행 기록이 없습니다.";

    public MainWindowViewModel(ITicketingAutomationService ticketingAutomationService)
    {
        _ticketingAutomationService = ticketingAutomationService;
        StartAutomationCommand = new AsyncRelayCommand(StartAutomationAsync);
    }

    public IAsyncRelayCommand StartAutomationCommand { get; }

    private async Task StartAutomationAsync()
    {
        StatusMessage = "실행 중...";

        var request = new TicketingJobRequest(EventKeyword, TargetUrl);
        var result = await _ticketingAutomationService.RunAsync(request, CancellationToken.None);

        StatusMessage = result.IsSuccess ? "실행 완료" : "실행 실패";
        LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
    }
}
