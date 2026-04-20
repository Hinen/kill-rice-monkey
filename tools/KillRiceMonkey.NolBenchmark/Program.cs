using System.Diagnostics;
using KillRiceMonkey.Application;
using KillRiceMonkey.Application.Abstractions;
using KillRiceMonkey.Application.Models;
using KillRiceMonkey.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;

namespace KillRiceMonkey.NolBenchmark;

// NOL 자동화 실측 벤치마크 러너
// - 실제 Chrome 9225 + NEW URL에 연결하여 NolAutomationService.RunAsync를 1회 호출
// - 각 Stage/로그 기반으로 구간별 경과시간 측정
// - 좌석 선점이 1회 발생할 수 있으므로 반복 실행은 피함
// - desiredBlock=1 로 자동 구역 선택, Chrome 세션은 로그인된 상태여야 함
internal static class Program
{
    private static async Task<int> Main(string[] args)
    {
        var baseDir = AppContext.BaseDirectory;
        var logDirectory = Path.Combine(baseDir, "logs");
        Directory.CreateDirectory(logDirectory);
        var logPath = Path.Combine(logDirectory, "bench-.log");

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .Enrich.FromLogContext()
            .WriteTo.Console()
            .WriteTo.File(
                logPath,
                rollingInterval: RollingInterval.Hour,
                shared: true,
                outputTemplate: "{Timestamp:HH:mm:ss.fff} [{Level:u3}] {Message:lj} {Properties:j}{NewLine}{Exception}")
            .CreateLogger();

        try
        {
            return await RunAsync(args);
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "Benchmark failed");
            return 1;
        }
        finally
        {
            await Log.CloseAndFlushAsync();
        }
    }

    private static async Task<int> RunAsync(string[] args)
    {
        var options = BenchmarkOptions.Parse(args);
        Log.Information("Benchmark options: {@Options}", options);

        using var host = Host.CreateDefaultBuilder()
            .UseSerilog()
            .ConfigureServices(services =>
            {
                services.AddApplication();
                services.AddInfrastructure();
            })
            .Build();

        await host.StartAsync();

        var service = host.Services.GetRequiredService<INolAutomationService>();
        var appLogger = host.Services.GetRequiredService<ILogger<BenchmarkRunner>>();

        var runner = new BenchmarkRunner(service, appLogger, options);
        var report = await runner.RunOnceAsync();

        Console.WriteLine();
        Console.WriteLine("============================================================");
        Console.WriteLine(" NOL Benchmark Report");
        Console.WriteLine("============================================================");
        Console.WriteLine($" Final result   : {(report.Success ? "SUCCESS" : "FAILURE")}");
        Console.WriteLine($" Final message  : {report.Message}");
        Console.WriteLine($" Total elapsed  : {report.TotalElapsedMs:N0} ms");
        Console.WriteLine();
        Console.WriteLine(" Stage breakdown");
        Console.WriteLine($" {"Stage",-30} {"Duration (ms)",15} {"Total (ms)",15}");
        foreach (var (stage, d, t) in report.Stages)
        {
            Console.WriteLine($" {Truncate(stage, 30),-30} {d,15:N0} {t,15:N0}");
        }
        Console.WriteLine("============================================================");
        Console.WriteLine($" Human average baseline (captcha solve + seat select + complete): ~8,000~12,000 ms");
        if (report.TotalElapsedMs > 0)
        {
            var ratio = 10000.0 / report.TotalElapsedMs;
            Console.WriteLine($" Automation speed vs human 10s baseline: {ratio:F2}x faster");
        }
        Console.WriteLine("============================================================");
        Console.WriteLine($" Log file: {logPath(host)}");

        await host.StopAsync();
        return report.Success ? 0 : 2;
    }

    private static string logPath(IHost host)
    {
        var logDir = Path.Combine(AppContext.BaseDirectory, "logs");
        var files = Directory.Exists(logDir) ? Directory.GetFiles(logDir, "bench-*.log") : Array.Empty<string>();
        return files.Length > 0 ? files[^1] : "<no log>";
    }

    private static string Truncate(string input, int maxLength)
        => input.Length <= maxLength ? input : input[..(maxLength - 1)] + "…";
}

internal sealed class BenchmarkOptions
{
    public required string DesiredDate { get; init; }
    public required string DesiredRound { get; init; }
    public string? DesiredBlock { get; init; }
    public int StepTimeoutSeconds { get; init; } = 8;
    public double MatchThreshold { get; init; } = 0.86;
    public bool LaunchIfMissing { get; init; } = true;
    // "좌석 선택 중" 스테이지 진입 즉시 CTS 취소. 구역 선택/fill 판정까지만 확인하고
    // 좌석 클릭을 하지 않아 좌석 선점을 방지한다(반복 실측용).
    public bool DryRunAfterZone { get; init; } = false;

    public static BenchmarkOptions Parse(string[] args)
    {
        // 기본값: NEW URL 26004589 코다라인 공연 2026.08.12 1회 20:00 (사용자 시나리오)
        // - 실제 벤치마크 시 CLI로 덮어쓰기
        string desiredDate = "2026.08.12";
        string desiredRound = "1회 20:00";
        string? desiredBlock = "1";
        int stepTimeoutSeconds = 8;
        double matchThreshold = 0.86;
        bool dryRunAfterZone = false;

        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "--date" when i + 1 < args.Length:
                    desiredDate = args[++i];
                    break;
                case "--round" when i + 1 < args.Length:
                    desiredRound = args[++i];
                    break;
                case "--block" when i + 1 < args.Length:
                    desiredBlock = args[++i];
                    break;
                case "--manual-zone":
                    // 명시적 수동 모드: DesiredBlock을 null로 강제 (PowerShell의 빈 문자열 인자 drop 회피).
                    desiredBlock = null;
                    break;
                case "--timeout" when i + 1 < args.Length:
                    stepTimeoutSeconds = int.Parse(args[++i]);
                    break;
                case "--threshold" when i + 1 < args.Length:
                    matchThreshold = double.Parse(args[++i]);
                    break;
                case "--dry-run-after-zone":
                    dryRunAfterZone = true;
                    break;
            }
        }

        return new BenchmarkOptions
        {
            DesiredDate = desiredDate,
            DesiredRound = desiredRound,
            DesiredBlock = desiredBlock,
            StepTimeoutSeconds = stepTimeoutSeconds,
            MatchThreshold = matchThreshold,
            DryRunAfterZone = dryRunAfterZone,
        };
    }
}

internal sealed record BenchmarkReport(
    bool Success,
    string Message,
    long TotalElapsedMs,
    IReadOnlyList<(string Stage, long DurationMs, long TotalMs)> Stages);

internal sealed class BenchmarkRunner(
    INolAutomationService service,
    ILogger<BenchmarkRunner> logger,
    BenchmarkOptions options)
{
    public async Task<BenchmarkReport> RunOnceAsync()
    {
        // 1. Chrome 9225 준비 확인
        var available = await service.IsRemoteDebugBrowserAvailableAsync(CancellationToken.None);
        logger.LogInformation("Remote debug browser available: {Available}", available);

        if (!available && options.LaunchIfMissing)
        {
            logger.LogInformation("Chrome 9225 미실행 — 자동 실행 중...");
            var url = await service.LaunchRemoteDebugBrowserAsync(CancellationToken.None);
            logger.LogInformation("Chrome launched: {Url}", url);
            Console.WriteLine();
            Console.WriteLine("== 벤치마크 준비 ==");
            Console.WriteLine("Chrome 9225 세션이 새로 열렸습니다.");
            Console.WriteLine("(이미 로그인된 프로필이면 바로 상품 페이지가 열립니다.)");
            Console.WriteLine("준비가 완료되면 Enter를 눌러주세요.");
            Console.ReadLine();
        }

        // 2. 상품 페이지가 열려있는지 확인. PageReady = 예매/캡챠/좌석/결제 어느 단계든 OK
        var readyWaitSw = Stopwatch.StartNew();
        while (!await service.IsPageReadyAsync(CancellationToken.None))
        {
            if (readyWaitSw.Elapsed > TimeSpan.FromSeconds(60))
            {
                logger.LogError("IsPageReadyAsync 대기 60초 초과 — 상품/예매 페이지가 로드되지 않음.");
                return new BenchmarkReport(false, "page not ready", 0, Array.Empty<(string, long, long)>());
            }
            await Task.Delay(500);
        }
        logger.LogInformation("IsPageReady=true (prep waited {Ms}ms)", readyWaitSw.ElapsedMilliseconds);

        // 3. Request 구성
        var request = new TicketingJobRequest(
            TicketingTemplateType.Nol,
            ImageDirectory: "button-images",
            MatchThreshold: options.MatchThreshold,
            StepTimeoutSeconds: options.StepTimeoutSeconds,
            DesiredDate: options.DesiredDate,
            DesiredRound: options.DesiredRound,
            PauseBeforeSeatSelection: false,
            PauseGate: null,
            DesiredGrade: null,
            DesiredBlock: options.DesiredBlock);

        logger.LogInformation("Issuing RunAsync with request {@Request}", new
        {
            request.TemplateType,
            request.DesiredDate,
            request.DesiredRound,
            request.DesiredBlock,
            request.StepTimeoutSeconds,
            request.MatchThreshold,
        });

        // 4. Stage 추적: Progress 콜백으로 stage 전환 시점 측정
        var stageTimes = new List<(string Stage, long DurationMs, long TotalMs)>();
        var totalSw = Stopwatch.StartNew();
        var stageSw = Stopwatch.StartNew();
        string? prevStage = null;
        // "좌석 선택 중"은 RunNolAutomation 상위/SelectNolOnestopSeatAndCompleteAsync 하위에서 두 번 보고된다.
        // dry-run-after-zone은 '구역 선택 대기 중' stage를 최초로 본 뒤 '좌석 선택 중'으로 전환되는 두 번째 시점에 cancel해야 한다.
        var sawZoneManualWait = false;
        using var runCts = new CancellationTokenSource();

        var progress = new Progress<AutomationProgress>(p =>
        {
            var stage = string.IsNullOrWhiteSpace(p.Stage) ? "(no-stage)" : p.Stage;
            if (stage != prevStage)
            {
                if (prevStage is not null)
                {
                    stageTimes.Add((prevStage, stageSw.ElapsedMilliseconds, totalSw.ElapsedMilliseconds));
                }
                Console.WriteLine($"  [T+{totalSw.ElapsedMilliseconds,7:N0} ms] Stage: {stage}");
                if (!string.IsNullOrWhiteSpace(p.LogMessage))
                    Console.WriteLine($"                 └─ {p.LogMessage}");

                // '구역 선택 대기 중' (수동 모드 마커) 등장 기록
                if (stage.Contains("구역 선택 대기"))
                    sawZoneManualWait = true;

                // '구역 선택 중'(자동 모드)도 마커로 취급
                if (stage.Contains("구역 선택 중") || stage.Contains("구역 선택 완료"))
                    sawZoneManualWait = true;

                // --dry-run-after-zone: 구역 단계 이후 '좌석 선택' stage 진입 시 cancel
                if (options.DryRunAfterZone && sawZoneManualWait && stage.Contains("좌석 선택") && !stage.Contains("완료"))
                {
                    logger.LogInformation("[dry-run-after-zone] 구역 이후 좌석 선택 stage 감지 — 취소 발사(선점 방지).");
                    try { runCts.Cancel(); } catch { }
                }

                prevStage = stage;
                stageSw.Restart();
            }
            else if (!string.IsNullOrWhiteSpace(p.LogMessage))
            {
                Console.WriteLine($"  [T+{totalSw.ElapsedMilliseconds,7:N0} ms]   └─ {p.LogMessage}");
            }
        });

        // 5. RunAsync 실행
        bool success;
        string message;
        try
        {
            var result = await service.RunAsync(request, progress, runCts.Token);
            success = result.IsSuccess;
            message = result.Message ?? "<no message>";
        }
        catch (OperationCanceledException) when (options.DryRunAfterZone)
        {
            success = true;
            message = "dry-run-after-zone: 좌석 선택 직전 취소(의도적).";
            logger.LogInformation(message);
        }
        catch (Exception ex)
        {
            success = false;
            message = ex.Message;
            logger.LogError(ex, "RunAsync threw");
        }

        if (prevStage is not null)
        {
            stageTimes.Add((prevStage, stageSw.ElapsedMilliseconds, totalSw.ElapsedMilliseconds));
        }
        totalSw.Stop();

        return new BenchmarkReport(success, message, totalSw.ElapsedMilliseconds, stageTimes);
    }
}
