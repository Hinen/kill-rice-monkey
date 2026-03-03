using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using System.Windows;

namespace KillRiceMonkey.App.ViewModels;

public partial class MainWindowViewModel : ObservableObject
{
    private const string Yes24FixedImageDirectory = "button-images/yes24";

    private readonly ITicketingAutomationService _ticketingAutomationService;

    public IReadOnlyList<string> TemplateOptions { get; } = ["Yes24", "Custom"];

    [ObservableProperty]
    private string _selectedTemplate = "Custom";

    [ObservableProperty]
    private string _imageDirectory = "button-images";

    public bool IsImageDirectoryEditable => !string.Equals(SelectedTemplate, "Yes24", StringComparison.OrdinalIgnoreCase);

    [ObservableProperty]
    private double _matchThreshold = 0.86;

    [ObservableProperty]
    private int _stepTimeoutSeconds = 8;

    [ObservableProperty]
    private string _hotkeyText = "F8";

    [ObservableProperty]
    private bool _isRunning;

    [ObservableProperty]
    private string _statusMessage = "준비 완료";

    [ObservableProperty]
    private string _lastRunSummary = "자동화 실행 기록이 없습니다.";

    public MainWindowViewModel(ITicketingAutomationService ticketingAutomationService)
    {
        _ticketingAutomationService = ticketingAutomationService;
        StartAutomationCommand = new AsyncRelayCommand(StartAutomationAsync, () => !IsRunning);
    }

    public IAsyncRelayCommand StartAutomationCommand { get; }

    public async Task HandleHotkeyAsync()
    {
        if (StartAutomationCommand.CanExecute(null))
        {
            await StartAutomationCommand.ExecuteAsync(null);
        }
    }

    partial void OnIsRunningChanged(bool value)
    {
        StartAutomationCommand.NotifyCanExecuteChanged();
    }

    partial void OnSelectedTemplateChanged(string value)
    {
        if (value.Equals("Yes24", StringComparison.OrdinalIgnoreCase))
        {
            ImageDirectory = Yes24FixedImageDirectory;
        }

        OnPropertyChanged(nameof(IsImageDirectoryEditable));
    }

    private async Task StartAutomationAsync()
    {
        var mainWindow = System.Windows.Application.Current?.MainWindow;
        var originalState = mainWindow?.WindowState ?? WindowState.Normal;

        if (mainWindow is not null)
        {
            mainWindow.WindowState = WindowState.Minimized;
        }

        IsRunning = true;
        StatusMessage = "active 상태: 다단계 이미지 클릭 실행 중";

        try
        {
            var templateType = ParseTemplateType(SelectedTemplate);
            var imageDirectory = templateType == TicketingTemplateType.Yes24 ? Yes24FixedImageDirectory : ImageDirectory;
            var request = new TicketingJobRequest(templateType, imageDirectory, MatchThreshold, StepTimeoutSeconds);
            var result = await _ticketingAutomationService.RunAsync(request, CancellationToken.None);

            StatusMessage = result.IsSuccess ? "성공 종료" : "예외 종료";
            LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 예외 발생: {ex.Message}";
        }
        finally
        {
            IsRunning = false;

            if (mainWindow is not null)
            {
                mainWindow.WindowState = originalState;
                mainWindow.Activate();
            }
        }
    }

    private static TicketingTemplateType ParseTemplateType(string value)
    {
        return value.Equals("Yes24", StringComparison.OrdinalIgnoreCase)
            ? TicketingTemplateType.Yes24
            : TicketingTemplateType.Custom;
    }
}
