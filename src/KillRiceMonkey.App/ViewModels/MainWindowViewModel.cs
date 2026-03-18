using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Windows.Input;

namespace KillRiceMonkey.App.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private const string Yes24FixedImageDirectory = "button-images/yes24";
    private const string BoothFixedImageDirectory = "button-images/booth";
    private const string MelonFixedImageDirectory = "button-images/melon";

    private readonly INolAutomationService _nolAutomationService;
    private readonly IMelonAutomationService _melonAutomationService;
    private readonly IImageAutomationService _imageAutomationService;
    private string? _selectedTemplate;
    private string _imageDirectory = "button-images";
    private DateTime? _desiredDate = DateTime.Today;
    private string _desiredTime = "18시 00분";
    private int _melonHour = 18;
    private int _melonMinute;
    private string _desiredRound = string.Empty;
    private double _matchThreshold = 0.86;
    private int _stepTimeoutSeconds = 8;
    private string _hotkeyText = "F8";
    private bool _isRunning;
    private string _statusMessage = "템플릿 선택 필요";
    private string _lastRunSummary = "자동화 실행 기록이 없습니다.";
    private readonly ObservableCollection<string> _runLogs = [];
    private DateTimeOffset _nolAutomationReadyAt;
    private DateTimeOffset _melonAutomationReadyAt;
    private CancellationTokenSource? _runCts;
    private bool _pauseBeforeSeatSelection;
    private ManualResetEventSlim? _pauseGate;
    private bool _isPaused;

    public MainWindowViewModel(
        INolAutomationService nolAutomationService,
        IMelonAutomationService melonAutomationService,
        IImageAutomationService imageAutomationService)
    {
        _nolAutomationService = nolAutomationService;
        _melonAutomationService = melonAutomationService;
        _imageAutomationService = imageAutomationService;
        LaunchNolRemoteDebugCommand = new AsyncCommand(LaunchNolRemoteDebugAsync, () => !IsRunning && IsNolTemplate);
        PrepareNolAutomationCommand = new AsyncCommand(PrepareNolAutomationAsync, () => !IsRunning && IsNolTemplate);
        LaunchMelonRemoteDebugCommand = new AsyncCommand(LaunchMelonRemoteDebugAsync, () => !IsRunning && IsMelonTemplate);
        PrepareMelonAutomationCommand = new AsyncCommand(PrepareMelonAutomationAsync, () => !IsRunning && IsMelonTemplate);
        StartAutomationCommand = new AsyncCommand(StartAutomationAsync, () => !IsRunning && HasSelectedTemplate);
        ResumeAutomationCommand = new AsyncCommand(ResumeAutomationAsync, () => IsPaused);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public IReadOnlyList<string> TemplateOptions { get; } = ["Yes24", "Booth", "NOL", "Melon", "Custom"];

    public AsyncCommand StartAutomationCommand { get; }

    public AsyncCommand LaunchNolRemoteDebugCommand { get; }

    public AsyncCommand PrepareNolAutomationCommand { get; }

    public AsyncCommand LaunchMelonRemoteDebugCommand { get; }

    public AsyncCommand PrepareMelonAutomationCommand { get; }

    public AsyncCommand ResumeAutomationCommand { get; }

    public string? SelectedTemplate
    {
        get => _selectedTemplate;
        set
        {
            if (!SetProperty(ref _selectedTemplate, value))
            {
                return;
            }

            HandleSelectedTemplateChanged(value);
        }
    }

    public string ImageDirectory
    {
        get => _imageDirectory;
        set => SetProperty(ref _imageDirectory, value);
    }

    public DateTime? DesiredDate
    {
        get => _desiredDate;
        set => SetProperty(ref _desiredDate, value);
    }

    public string DesiredTime
    {
        get => _desiredTime;
        set => SetProperty(ref _desiredTime, value);
    }

    public IReadOnlyList<int> HourOptions { get; } = Enumerable.Range(0, 24).ToList();

    public IReadOnlyList<int> MinuteOptions { get; } = Enumerable.Range(0, 60).ToList();

    public int MelonHour
    {
        get => _melonHour;
        set
        {
            if (SetProperty(ref _melonHour, value))
                DesiredTime = $"{_melonHour:D2}시 {_melonMinute:D2}분";
        }
    }

    public int MelonMinute
    {
        get => _melonMinute;
        set
        {
            if (SetProperty(ref _melonMinute, value))
                DesiredTime = $"{_melonHour:D2}시 {_melonMinute:D2}분";
        }
    }

    public string DesiredRound
    {
        get => _desiredRound;
        set => SetProperty(ref _desiredRound, value);
    }

    public double MatchThreshold
    {
        get => _matchThreshold;
        set => SetProperty(ref _matchThreshold, value);
    }

    public int StepTimeoutSeconds
    {
        get => _stepTimeoutSeconds;
        set => SetProperty(ref _stepTimeoutSeconds, value);
    }

    public bool PauseBeforeSeatSelection
    {
        get => _pauseBeforeSeatSelection;
        set => SetProperty(ref _pauseBeforeSeatSelection, value);
    }

    public string HotkeyText
    {
        get => _hotkeyText;
        set => SetProperty(ref _hotkeyText, value);
    }

    public bool IsRunning
    {
        get => _isRunning;
        private set
        {
            if (!SetProperty(ref _isRunning, value))
            {
                return;
            }

            StartAutomationCommand.NotifyCanExecuteChanged();
            LaunchNolRemoteDebugCommand.NotifyCanExecuteChanged();
            PrepareNolAutomationCommand.NotifyCanExecuteChanged();
            LaunchMelonRemoteDebugCommand.NotifyCanExecuteChanged();
            PrepareMelonAutomationCommand.NotifyCanExecuteChanged();
        }
    }

    public string StatusMessage
    {
        get => _statusMessage;
        private set
        {
            if (!SetProperty(ref _statusMessage, value))
            {
                return;
            }

            OnPropertyChanged(nameof(IsSuccessStatus));
            OnPropertyChanged(nameof(IsErrorStatus));
        }
    }

    public string LastRunSummary
    {
        get => _lastRunSummary;
        private set => SetProperty(ref _lastRunSummary, value);
    }

    public bool IsPaused
    {
        get => _isPaused;
        private set
        {
            if (!SetProperty(ref _isPaused, value))
            {
                return;
            }

            ResumeAutomationCommand.NotifyCanExecuteChanged();
        }
    }

    public ObservableCollection<string> RunLogs => _runLogs;

    public bool IsImageDirectoryEditable => HasSelectedTemplate && ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Custom;

    public bool HasSelectedTemplate => !string.IsNullOrWhiteSpace(SelectedTemplate);

    public bool IsNolTemplate => ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Nol;

    public bool IsMelonTemplate => ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Melon;

    public bool IsSeatPauseSupported => IsMelonTemplate;

    public bool IsSuccessStatus => string.Equals(StatusMessage, "성공 종료", StringComparison.Ordinal);

    public bool IsErrorStatus => string.Equals(StatusMessage, "예외 종료", StringComparison.Ordinal);

    public async Task HandleHotkeyAsync()
    {
        if (StartAutomationCommand.CanExecute(null))
        {
            await StartAutomationCommand.ExecuteAsync();
        }
    }

    public void CancelAutomation()
    {
        if (!IsRunning || _runCts is null)
            return;

        _runCts.Cancel();
        StatusMessage = "취소 중";
    }

    private async Task StartAutomationAsync()
    {
        if (!HasSelectedTemplate)
        {
            StatusMessage = "템플릿 선택 필요";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 실행 전 템플릿을 선택하세요.";
            return;
        }

        var templateType = ParseTemplateType(SelectedTemplate);

        if (templateType == TicketingTemplateType.Nol)
        {
            if (DesiredDate is null)
            {
                StatusMessage = "입력 확인 필요";
                LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 관람일을 선택하세요.";
                return;
            }

            if (string.IsNullOrWhiteSpace(DesiredRound))
            {
                StatusMessage = "입력 확인 필요";
                LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 회차를 입력하세요. 예: 1회 19:00";
                return;
            }

            var readyCacheValid = (DateTimeOffset.UtcNow - _nolAutomationReadyAt).TotalMinutes < 10;
            if (!readyCacheValid)
            {
                if (!await _nolAutomationService.IsRemoteDebugBrowserAvailableAsync(CancellationToken.None))
                {
                    StatusMessage = "remote debug 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 먼저 'NOL Remote Debug 열기' 버튼으로 브라우저를 실행하세요.";
                    return;
                }

                if (!await _nolAutomationService.IsAutomationPreparedAsync(CancellationToken.None))
                {
                    StatusMessage = "NOL 준비 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 상품 페이지를 연 뒤 'NOL 준비' 버튼으로 연결을 미리 준비하세요.";
                    return;
                }
            }
        }

        if (templateType == TicketingTemplateType.Melon)
        {
            if (DesiredDate is null)
            {
                StatusMessage = "입력 확인 필요";
                LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 관람일을 선택하세요.";
                return;
            }

            if (string.IsNullOrWhiteSpace(DesiredTime))
            {
                StatusMessage = "입력 확인 필요";
                LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 시간을 선택하세요.";
                return;
            }

            var readyCacheValid = (DateTimeOffset.UtcNow - _melonAutomationReadyAt).TotalMinutes < 10;
            if (!readyCacheValid)
            {
                if (!await _melonAutomationService.IsRemoteDebugBrowserAvailableAsync(CancellationToken.None))
                {
                    StatusMessage = "remote debug 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 먼저 'Melon Remote Debug 열기' 버튼으로 브라우저를 실행하세요.";
                    return;
                }

                if (!await _melonAutomationService.IsAutomationPreparedAsync(CancellationToken.None))
                {
                    StatusMessage = "Melon 준비 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 상품 페이지를 연 뒤 'Melon 준비' 버튼으로 연결을 미리 준비하세요.";
                    return;
                }
            }
        }

        _runCts?.Dispose();
        _runCts = new CancellationTokenSource();
        ResetPauseGate();
        if (PauseBeforeSeatSelection && IsSeatPauseSupported)
        {
            _pauseGate = new ManualResetEventSlim(false);
        }

        _runLogs.Clear();
        AppendLog($"{SelectedTemplate} 자동화 시작");

        IsRunning = true;

        try
        {
            var imageDirectory = templateType switch
            {
                TicketingTemplateType.Yes24 => Yes24FixedImageDirectory,
                TicketingTemplateType.Booth => BoothFixedImageDirectory,
                TicketingTemplateType.Melon => MelonFixedImageDirectory,
                _ => ImageDirectory
            };

            var request = new TicketingJobRequest(
                templateType,
                imageDirectory,
                MatchThreshold,
                StepTimeoutSeconds,
                IsNolTemplate || IsMelonTemplate ? DesiredDate?.ToString("yyyy.MM.dd") : null,
                IsNolTemplate ? DesiredRound : (IsMelonTemplate ? DesiredTime : null),
                PauseBeforeSeatSelection && IsSeatPauseSupported,
                _pauseGate);

            var progress = new Progress<AutomationProgress>(automationProgress =>
            {
                StatusMessage = automationProgress.Stage;
                UpdatePauseState(automationProgress.Stage);
                if (!string.IsNullOrWhiteSpace(automationProgress.LogMessage))
                {
                    AppendLog(automationProgress.LogMessage);
                }
            });

            if (templateType == TicketingTemplateType.Nol)
            {
                var attempt = 0;
                while (true)
                {
                    _runCts.Token.ThrowIfCancellationRequested();
                    attempt++;

                    if (!await _nolAutomationService.IsPageReadyAsync(_runCts.Token))
                    {
                        StatusMessage = $"NOL 페이지 대기 중 ({attempt}회 폴링)";
                        await Task.Delay(200, _runCts.Token);
                        continue;
                    }

                    StatusMessage = "NOL 자동화 실행 중";
                    var result = await _nolAutomationService.RunAsync(request, progress, _runCts.Token);

                    if (result.IsSuccess)
                    {
                        AppendLog(result.Message);
                        StatusMessage = "성공 종료";
                        LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
                        return;
                    }

                    AppendLog($"시도 실패 — {result.Message}");
                    LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {attempt}회 시도 실패 — {result.Message}";
                    await Task.Delay(500, _runCts.Token);
                }
            }
            else if (templateType == TicketingTemplateType.Melon)
            {
                var attempt = 0;
                while (true)
                {
                    _runCts.Token.ThrowIfCancellationRequested();
                    attempt++;

                    if (!await _melonAutomationService.IsPageReadyAsync(_runCts.Token))
                    {
                        StatusMessage = $"Melon 페이지 대기 중 ({attempt}회 폴링)";
                        await Task.Delay(50, _runCts.Token);
                        continue;
                    }

                    StatusMessage = "Melon 자동화 실행 중";
                    var result = await _melonAutomationService.RunAsync(request, progress, _runCts.Token);

                    if (result.IsSuccess)
                    {
                        AppendLog(result.Message);
                        StatusMessage = "성공 종료";
                        LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
                        return;
                    }

                    AppendLog($"시도 실패 — {result.Message}");
                    LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {attempt}회 시도 실패 — {result.Message}";
                    await Task.Delay(100, _runCts.Token);
                }
            }
            else
            {
                StatusMessage = "active 상태: 다단계 이미지 클릭 실행 중";
                var result = await _imageAutomationService.RunAsync(request, progress, _runCts.Token);
                AppendLog(result.Message);
                StatusMessage = result.IsSuccess ? "성공 종료" : "예외 종료";
                LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
            }
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "취소 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 사용자에 의해 취소되었습니다.";
            AppendLog("사용자에 의해 취소되었습니다.");
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 예외 발생: {ex.Message}";
            AppendLog($"예외 발생: {ex.Message}");
        }
        finally
        {
            IsPaused = false;
            IsRunning = false;
            ResetPauseGate();
            _runCts?.Dispose();
            _runCts = null;
        }
    }

    private Task ResumeAutomationAsync()
    {
        _pauseGate?.Set();
        IsPaused = false;
        return Task.CompletedTask;
    }

    private async Task LaunchNolRemoteDebugAsync()
    {
        if (!IsNolTemplate)
        {
            return;
        }

        IsRunning = true;
        StatusMessage = "remote debug 브라우저 실행 중";

        try
        {
            var message = await _nolAutomationService.LaunchRemoteDebugBrowserAsync(CancellationToken.None);
            StatusMessage = "성공 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | {message}";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | remote debug 브라우저 실행 실패: {ex.Message}";
        }
        finally
        {
            IsRunning = false;
        }
    }

    private async Task PrepareNolAutomationAsync()
    {
        if (!IsNolTemplate)
        {
            return;
        }

        IsRunning = true;
        StatusMessage = "NOL 자동화 준비 중";

        try
        {
            var message = await _nolAutomationService.PrepareAutomationAsync(CancellationToken.None);
            _nolAutomationReadyAt = DateTimeOffset.UtcNow;
            StatusMessage = "성공 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | {message}";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | NOL 준비 실패: {ex.Message}";
        }
        finally
        {
            IsRunning = false;
        }
    }

    private async Task LaunchMelonRemoteDebugAsync()
    {
        if (!IsMelonTemplate)
        {
            return;
        }

        IsRunning = true;
        StatusMessage = "remote debug 브라우저 실행 중";

        try
        {
            var message = await _melonAutomationService.LaunchRemoteDebugBrowserAsync(CancellationToken.None);
            StatusMessage = "성공 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | {message}";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | remote debug 브라우저 실행 실패: {ex.Message}";
        }
        finally
        {
            IsRunning = false;
        }
    }

    private async Task PrepareMelonAutomationAsync()
    {
        if (!IsMelonTemplate)
        {
            return;
        }

        IsRunning = true;
        StatusMessage = "Melon 자동화 준비 중";

        try
        {
            var message = await _melonAutomationService.PrepareAutomationAsync(CancellationToken.None);
            _melonAutomationReadyAt = DateTimeOffset.UtcNow;
            StatusMessage = "성공 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | {message}";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | Melon 준비 실패: {ex.Message}";
        }
        finally
        {
            IsRunning = false;
        }
    }

    private void HandleSelectedTemplateChanged(string? value)
    {
        var templateType = ParseTemplateType(value);
        if (templateType == TicketingTemplateType.Yes24)
        {
            ImageDirectory = Yes24FixedImageDirectory;
        }
        else if (templateType == TicketingTemplateType.Booth)
        {
            ImageDirectory = BoothFixedImageDirectory;
        }
        else if (templateType == TicketingTemplateType.Melon)
        {
            ImageDirectory = MelonFixedImageDirectory;
        }
        else if (templateType == TicketingTemplateType.Nol)
        {
            ImageDirectory = "button-images";
        }
        else if (!HasSelectedTemplate)
        {
            ImageDirectory = "button-images";
            StatusMessage = "템플릿 선택 필요";
        }

        OnPropertyChanged(nameof(IsImageDirectoryEditable));
        OnPropertyChanged(nameof(HasSelectedTemplate));
        OnPropertyChanged(nameof(IsNolTemplate));
        OnPropertyChanged(nameof(IsMelonTemplate));
        OnPropertyChanged(nameof(IsSeatPauseSupported));
        StartAutomationCommand.NotifyCanExecuteChanged();
        LaunchNolRemoteDebugCommand.NotifyCanExecuteChanged();
        PrepareNolAutomationCommand.NotifyCanExecuteChanged();
        LaunchMelonRemoteDebugCommand.NotifyCanExecuteChanged();
        PrepareMelonAutomationCommand.NotifyCanExecuteChanged();
    }

    private void AppendLog(string message)
    {
        _runLogs.Add($"{DateTimeOffset.Now:HH:mm:ss} | {message}");
    }

    private void UpdatePauseState(string stage)
    {
        IsPaused = string.Equals(stage, "좌석 선택 대기 — 일시정지", StringComparison.Ordinal);
    }

    private void ResetPauseGate()
    {
        IsPaused = false;
        _pauseGate?.Dispose();
        _pauseGate = null;
    }

    private bool SetProperty<T>(ref T field, T value, [CallerMemberName] string? propertyName = null)
    {
        if (EqualityComparer<T>.Default.Equals(field, value))
        {
            return false;
        }

        field = value;
        OnPropertyChanged(propertyName);
        return true;
    }

    private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
    {
        if (propertyName is null)
        {
            return;
        }

        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }

    private static TicketingTemplateType ParseTemplateType(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return TicketingTemplateType.Custom;
        }

        if (value.Equals("Yes24", StringComparison.OrdinalIgnoreCase))
        {
            return TicketingTemplateType.Yes24;
        }

        if (value.Equals("Booth", StringComparison.OrdinalIgnoreCase))
        {
            return TicketingTemplateType.Booth;
        }

        if (value.Equals("NOL", StringComparison.OrdinalIgnoreCase))
        {
            return TicketingTemplateType.Nol;
        }

        if (value.Equals("Melon", StringComparison.OrdinalIgnoreCase))
        {
            return TicketingTemplateType.Melon;
        }

        return TicketingTemplateType.Custom;
    }

    public sealed class AsyncCommand : ICommand
    {
        private readonly Func<Task> _executeAsync;
        private readonly Func<bool>? _canExecute;
        private bool _isExecuting;

        public AsyncCommand(Func<Task> executeAsync, Func<bool>? canExecute = null)
        {
            _executeAsync = executeAsync;
            _canExecute = canExecute;
        }

        public event EventHandler? CanExecuteChanged;

        public bool CanExecute(object? parameter)
        {
            return !_isExecuting && (_canExecute?.Invoke() ?? true);
        }

        public async void Execute(object? parameter)
        {
            await ExecuteAsync();
        }

        public async Task ExecuteAsync()
        {
            if (!CanExecute(null))
            {
                return;
            }

            _isExecuting = true;
            NotifyCanExecuteChanged();
            try
            {
                await _executeAsync();
            }
            finally
            {
                _isExecuting = false;
                NotifyCanExecuteChanged();
            }
        }

        public void NotifyCanExecuteChanged()
        {
            CanExecuteChanged?.Invoke(this, EventArgs.Empty);
        }
    }
}
