namespace KillRiceMonkey.Application.Models;

public sealed record TicketingJobRequest(TicketingTemplateType TemplateType, string ImageDirectory, double MatchThreshold, int StepTimeoutSeconds);
