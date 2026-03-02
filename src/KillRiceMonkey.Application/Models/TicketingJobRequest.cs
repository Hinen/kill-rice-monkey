namespace KillRiceMonkey.Application.Models;

public sealed record TicketingJobRequest(string ImageDirectory, double MatchThreshold, int StepTimeoutSeconds);
