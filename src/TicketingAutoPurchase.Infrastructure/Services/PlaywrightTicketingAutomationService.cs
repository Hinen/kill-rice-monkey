using Microsoft.Extensions.Logging;
using Polly;
using TicketingAutoPurchase.Application.Abstractions;
using TicketingAutoPurchase.Application.Models;

namespace TicketingAutoPurchase.Infrastructure.Services;

public sealed class PlaywrightTicketingAutomationService : ITicketingAutomationService
{
    private readonly ILogger<PlaywrightTicketingAutomationService> _logger;
    private readonly ResiliencePipeline _pipeline;

    public PlaywrightTicketingAutomationService(ILogger<PlaywrightTicketingAutomationService> logger)
    {
        _logger = logger;
        _pipeline = new ResiliencePipelineBuilder()
            .AddRetry(new Polly.Retry.RetryStrategyOptions
            {
                MaxRetryAttempts = 2,
                Delay = TimeSpan.FromMilliseconds(300)
            })
            .Build();
    }

    public async Task<AutomationRunResult> RunAsync(TicketingJobRequest request, CancellationToken cancellationToken)
    {
        return await _pipeline.ExecuteAsync(async token =>
        {
            _logger.LogInformation("Automation started. keyword={Keyword}, url={Url}", request.EventKeyword, request.TargetUrl);

            await Task.Delay(800, token);

            var message = $"초기 자동화 파이프라인 실행 완료 (keyword: {request.EventKeyword}, url: {request.TargetUrl})";
            _logger.LogInformation(message);

            return new AutomationRunResult(true, message, DateTimeOffset.Now);
        }, cancellationToken);
    }
}
