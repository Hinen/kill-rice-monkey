namespace KillRiceMonkey.Application.Models;

public sealed record TicketingJobRequest(
    TicketingTemplateType TemplateType,
    string ImageDirectory,
    double MatchThreshold,
    int StepTimeoutSeconds,
    string? DesiredDate = null,
    string? DesiredRound = null,
    bool PauseBeforeSeatSelection = false,
    ManualResetEventSlim? PauseGate = null);
