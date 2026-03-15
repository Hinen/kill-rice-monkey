namespace KillRiceMonkey.Application.Models;

public sealed record AutomationProgress(string Stage, string? LogMessage = null);
