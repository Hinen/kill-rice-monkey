namespace KillRiceMonkey.Application.Models;

public sealed record AutomationRunResult(bool IsSuccess, string Message, DateTimeOffset ExecutedAt);
