# AGENTS.md

This file stores repository-level persistent instructions for the coding agent across sessions.

## Scope
- This repository-specific file focuses on project context and constraints.
- Cross-project common rules are defined in `C:\Users\Hinen\.config\opencode\AGENTS.md`.

## Project Context (Ticketing Helper)
- Goal: build a Windows C# GUI helper for ticketing actions.
- The actual ticketing website is opened and handled by the user.
- Do not implement site login automation in this project.
- Use explicit user-triggered actions (button/hotkey) as the start point for automation steps.

## Security and Compliance Constraints
- Do not implement anti-detection, stealth, or macro-evasion behavior.
- Do not claim or attempt "undetectable" automation.
- Prefer transparent, user-controlled flows and compliant automation boundaries.

## Current Architecture Snapshot
- Solution: `KillRiceMonkey.sln`
- App layer: WPF + MVVM (`src/KillRiceMonkey.App`)
- Core layers: Application/Domain/Infrastructure under `src/`
- Test layer: xUnit project under `tests/KillRiceMonkey.Tests`
