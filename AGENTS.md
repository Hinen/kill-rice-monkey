# AGENTS.md

This file stores repository-level persistent instructions for the coding agent across sessions.

## Scope
- This repository-specific file focuses on project context and constraints.
- Cross-project common rules are defined in `C:\Users\Hinen\.config\opencode\AGENTS.md`.

## Project Context (Ticketing Helper)
- Goal: 티켓팅, 선착순 구매를 DOM or Image 캡쳐로 자동화 처리를 지원하는 Windows C# GUI 프로그램
- The actual ticketing website is opened and handled by the user.
- Do not implement site login automation in this project.
- Use explicit user-triggered actions (button/hotkey) as the start point for automation steps.

## DOM 자동화 작업 규칙
- DOM 방식으로 진행하는 자동화 템플릿에서 플랜, 계획, 설계를 세울 때는 반드시 Chrome remote-debugging을 스스로 실행하여 사용자가 전달한 URL에 접속
- Chrome remote-debugging에서 사용자가 요청한 시나리오대로 스스로 홈페이지 조작 및 이동하며 플랜, 계획, 설계, 현재 작업 방식을 검토
- Chrome remote-debugging에서 Login이 필요한 상황에서는 사용자에게 질문을 통해 Login 요청을 보냄
- 사용자가 미리 Login 정보를 전달한 상황이라면, 그 정보를 활용하여 Login 수행
- 이 상황에서 정말 예외적인 상황이 아니라면, 사용자의 개입은 하나도 없어야 함

## CAPTCHA 표본 수집
- CAPTCHA 표본을 수집할때에도 마찬가지로 반드시 Chrome remote-debugging을 스스로 실행하여 사용자가 전달한 URL에 접속
- Chrome remote-debugging에서 직접 CAPTCHA 발생하는 순간까지 직접 홈페이지 조작 및 이동 수행
- Chrome remote-debugging에서 Login이 필요한 상황에서는 사용자에게 질문을 통해 Login 요청을 보냄
- 사용자가 미리 Login 정보를 전달한 상황이라면, 그 정보를 활용하여 Login 수행
- 이 상황에서 정말 예외적인 상황이 아니라면, 사용자의 개입은 하나도 없어야 함

## 성능 최적화
- 티켓팅, 선착순 구매는 속도가 생명이기에, 최소한 사람이 직접 하는 것 보다는 빠른 성능이 나와야함
- 최소 성능이 사람보다 빠른 것이지, 이 최소값을 목표로 잡으면 안됨
- 경쟁하는 것은 결국엔 이런 자동화 프로그램끼리 경쟁을 할테니, 최대한 자동화의 처리 속도가 빨라야 함

## 작업 검수
- 어떠한 작업이라도 작업이 완료되면 반드시 Chrome remote-debugging을 실행하여 작업에 문제가 없는지 직접 자동화 검증을 해야함
- 검증 과정에서 문제가 발생했다면, 다시 작업을 수정하고 작업 검수를 무한 반복해야함
- 이 과정에서 검수를 성공적으로 끝냈을 경우에만 최종 작업 종료로 판단

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
- 빌드, 테스트, 검증을 위해서 사용한 임시 데이터들은 작업 종료 후 반드시 제거한다.
- git worktree에 잔여물이 최대한 남지 않도록 한다.
- 커밋 대상이 애매한 작업물은 사용자에게 질문 후 응답에 맞게 처리한다.
