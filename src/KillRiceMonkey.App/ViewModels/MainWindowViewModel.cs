using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Input;

namespace KillRiceMonkey.App.ViewModels;

public sealed class MainWindowViewModel : INotifyPropertyChanged
{
    private const string Yes24FixedImageDirectory = "button-images/yes24";
    private const string BoothFixedImageDirectory = "button-images/booth";
    private const string MelonFixedImageDirectory = "button-images/melon";

    private readonly ITicketingAutomationService _ticketingAutomationService;
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
    private DateTimeOffset _nolAutomationReadyAt;
    private DateTimeOffset _melonAutomationReadyAt;
    private CancellationTokenSource? _runCts;

    public MainWindowViewModel(ITicketingAutomationService ticketingAutomationService)
    {
        _ticketingAutomationService = ticketingAutomationService;
        LaunchNolRemoteDebugCommand = new AsyncCommand(LaunchNolRemoteDebugAsync, () => !IsRunning && IsNolTemplate);
        PrepareNolAutomationCommand = new AsyncCommand(PrepareNolAutomationAsync, () => !IsRunning && IsNolTemplate);
        LaunchMelonRemoteDebugCommand = new AsyncCommand(LaunchMelonRemoteDebugAsync, () => !IsRunning && IsMelonTemplate);
        PrepareMelonAutomationCommand = new AsyncCommand(PrepareMelonAutomationAsync, () => !IsRunning && IsMelonTemplate);
        StartAutomationCommand = new AsyncCommand(StartAutomationAsync, () => !IsRunning && HasSelectedTemplate);
    }

    public event PropertyChangedEventHandler? PropertyChanged;

    public IReadOnlyList<string> TemplateOptions { get; } = ["Yes24", "Booth", "NOL", "Melon", "Custom"];

    public AsyncCommand StartAutomationCommand { get; }

    public AsyncCommand LaunchNolRemoteDebugCommand { get; }

    public AsyncCommand PrepareNolAutomationCommand { get; }

    public AsyncCommand LaunchMelonRemoteDebugCommand { get; }

    public AsyncCommand PrepareMelonAutomationCommand { get; }

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

    public bool IsImageDirectoryEditable => HasSelectedTemplate && ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Custom;

    public bool HasSelectedTemplate => !string.IsNullOrWhiteSpace(SelectedTemplate);

    public bool IsNolTemplate => ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Nol;

    public bool IsMelonTemplate => ParseTemplateType(SelectedTemplate) == TicketingTemplateType.Melon;

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

        var mainWindow = System.Windows.Application.Current?.MainWindow;
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
                if (!await _ticketingAutomationService.IsNolRemoteDebugBrowserAvailableAsync(CancellationToken.None))
                {
                    StatusMessage = "remote debug 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 먼저 'NOL Remote Debug 열기' 버튼으로 브라우저를 실행하세요.";
                    return;
                }

                if (!await _ticketingAutomationService.IsNolAutomationPreparedAsync(CancellationToken.None))
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
                if (!await _ticketingAutomationService.IsMelonRemoteDebugBrowserAvailableAsync(CancellationToken.None))
                {
                    StatusMessage = "remote debug 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 먼저 'Melon Remote Debug 열기' 버튼으로 브라우저를 실행하세요.";
                    return;
                }

                if (!await _ticketingAutomationService.IsMelonAutomationPreparedAsync(CancellationToken.None))
                {
                    StatusMessage = "Melon 준비 필요";
                    LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 상품 페이지를 연 뒤 'Melon 준비' 버튼으로 연결을 미리 준비하세요.";
                    return;
                }
            }
        }

        if (mainWindow is not null)
        {
            mainWindow.WindowState = WindowState.Minimized;
        }

        _runCts?.Dispose();
        _runCts = new CancellationTokenSource();

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
                IsNolTemplate ? DesiredRound : (IsMelonTemplate ? DesiredTime : null));

            if (templateType == TicketingTemplateType.Nol)
            {
                var attempt = 0;
                while (true)
                {
                    _runCts.Token.ThrowIfCancellationRequested();
                    attempt++;

                    if (!await _ticketingAutomationService.IsNolPageReadyAsync(_runCts.Token))
                    {
                        StatusMessage = $"NOL 페이지 대기 중 ({attempt}회 폴링)";
                        await Task.Delay(200, _runCts.Token);
                        continue;
                    }

                    StatusMessage = "NOL 자동화 실행 중";
                    var result = await _ticketingAutomationService.RunAsync(request, _runCts.Token);

                    if (result.IsSuccess)
                    {
                        StatusMessage = "성공 종료";
                        LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
                        return;
                    }

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

                    if (!await _ticketingAutomationService.IsMelonPageReadyAsync(_runCts.Token))
                    {
                        StatusMessage = $"Melon 페이지 대기 중 ({attempt}회 폴링)";
                        await Task.Delay(50, _runCts.Token);
                        continue;
                    }

                    StatusMessage = "Melon 자동화 실행 중";
                    var result = await _ticketingAutomationService.RunAsync(request, _runCts.Token);

                    if (result.IsSuccess)
                    {
                        StatusMessage = "성공 종료";
                        LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
                        return;
                    }

                    LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {attempt}회 시도 실패 — {result.Message}";
                    await Task.Delay(100, _runCts.Token);
                }
            }
            else
            {
                StatusMessage = "active 상태: 다단계 이미지 클릭 실행 중";
                var result = await _ticketingAutomationService.RunAsync(request, _runCts.Token);
                StatusMessage = result.IsSuccess ? "성공 종료" : "예외 종료";
                LastRunSummary = $"{result.ExecutedAt:yyyy-MM-dd HH:mm:ss} | {result.Message}";
            }
        }
        catch (OperationCanceledException)
        {
            StatusMessage = "취소 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 사용자에 의해 취소되었습니다.";
        }
        catch (Exception ex)
        {
            StatusMessage = "예외 종료";
            LastRunSummary = $"{DateTimeOffset.Now:yyyy-MM-dd HH:mm:ss} | 예외 발생: {ex.Message}";
        }
        finally
        {
            IsRunning = false;
            _runCts?.Dispose();
            _runCts = null;
        }
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
            var message = await _ticketingAutomationService.LaunchNolRemoteDebugBrowserAsync(CancellationToken.None);
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
            var message = await _ticketingAutomationService.PrepareNolAutomationAsync(CancellationToken.None);
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
            var message = await _ticketingAutomationService.LaunchMelonRemoteDebugBrowserAsync(CancellationToken.None);
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
            var message = await _ticketingAutomationService.PrepareMelonAutomationAsync(CancellationToken.None);
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
        StartAutomationCommand.NotifyCanExecuteChanged();
        LaunchNolRemoteDebugCommand.NotifyCanExecuteChanged();
        PrepareNolAutomationCommand.NotifyCanExecuteChanged();
        LaunchMelonRemoteDebugCommand.NotifyCanExecuteChanged();
        PrepareMelonAutomationCommand.NotifyCanExecuteChanged();
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
