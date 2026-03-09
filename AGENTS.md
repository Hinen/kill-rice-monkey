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

## Commit Execution Policy (Project-Level Override)
- Plan 요청과 실제 작업 요청을 구분한다.
- 실제 작업 요청을 수행한 경우, 별도 추가 지시가 없어도 작업 종료 전에 반드시 커밋까지 완료한다.
- 커밋은 작업 단위를 작게 나누어 단계별로 수행한다.
- 커밋 메시지는 한국어로 작성하고, 변수명/기술 용어는 영어를 유지한다.
- 검증(가능한 범위의 build/test/lint) 후 커밋한다.
- 실제 작업 요청으로 변경이 발생했다면, 작업 종료 전에 `dotnet publish` 기준의 릴리즈 산출물까지 생성하고 확인한다.
