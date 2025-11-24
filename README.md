**KoreanIMEFixer**

간단 소개
- KoreanIMEFixer는 Windows용 WPF 유틸리티로, 한글 입력(IME) 관련 입력 이상 현상(특히 Notion의 테이블 셀 등에서 엔터 직후 초성 누락 등)을 완화하기 위한 여러 계층의 입력 복원(fallback) 기법을 제공합니다.

주요 기능
- 엔터 후 빠른 복원(스캔코드 SendInput 빠른 경로)
- 여러 입력 재주입(fallback): scancode SendInput, virtual-key SendInput, PostMessage, 클립보드 붙여넣기(Ctrl+V) 등
- Notion 테이블과 유사한 환경을 감지하는 UIA 기반 휴리스틱(빠른 캐시형 및 샘플링 방식)
- 전역 키후킹 및 입력 분석 로직 분리
- 상세 로깅(`%APPDATA%\\KoreanIMEFixer\\app.log`)으로 성능/동작 분석 지원

빌드 요구사항
- Windows 10/11
- .NET SDK 10 (net10.0-windows)

빌드 및 실행
1. 저장소 루트에서 빌드:
```powershell
dotnet build "KoreanIMEFixer.sln" -c Debug
```

2. 실행 (디버그):
```powershell
dotnet run --project "KoreanIMEFixer\\KoreanIMEFixer.csproj" -c Debug
```

테스트
```powershell
dotnet test "KoreanIMEFixer.Tests\\KoreanIMEFixer.Tests.csproj"
```

로그
- 동작 로그 파일: `%APPDATA%\\KoreanIMEFixer\\app.log` — heuristic 실행 시간, 분기 로그 등을 남깁니다. 문제 재현 시 최근 로그 200줄을 첨부해 주세요.

성능 / 트러블슈팅 팁
- Notion 감지(휴리스틱)가 느리다고 느껴지면 `Input/AutomationHelpers.cs`의 파라미터를 조절하세요 (샘플/방문 최대값, 캐시 TTL). 또한 `MainWindow.xaml.cs`에서 빠른 캐시형 검사(`IsLikelyNotionTableFast`)가 사용되도록 설정되어 있는지 확인하세요.
- 엔터 복원 동작을 더 공격적으로(더 빠르게) 하려면 `IME/IMEStateRestorer.cs`의 `SendEnterScancodeDoubleFast()` 및 관련 미세대기(`Thread.Sleep`) 값을 검토하세요.

기여
- 기능 개선, 휴리스틱 튜닝, 버그 리포트 환영합니다. PR 또는 issue를 열어 주세요.

라이선스
- 본 프로젝트는 아래의 MIT 라이선스에 따라 배포됩니다. 자세한 내용은 `LICENSE` 파일을 확인하세요.
