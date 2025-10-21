# INCA COM-API 자동화 스크립트

INCA 소프트웨어를 위한 Python 기반 자동화 스크립트입니다. Excel 파일에서 캘리브레이션 변수를 읽어 ECU에 적용하고, 실시간 측정 데이터를 CSV 파일로 저장합니다.

[![Python Version](https://img.shields.io/badge/python-3.7%2B-blue)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Code Style](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

## 📋 목차

- [주요 기능](#-주요-기능)
- [필수 요구사항](#-필수-요구사항)
- [설치 방법](#-설치-방법)
- [사용 방법](#-사용-방법)
- [입력 파일 형식](#-입력-파일-형식)
- [출력 파일 형식](#-출력-파일-형식)
- [프로젝트 구조](#-프로젝트-구조)
- [고급 기능](#-고급-기능)
- [문제 해결](#-문제-해결)
- [기여 방법](#-기여-방법)
- [라이선스](#-라이선스)

## 🚀 주요 기능

- ✅ **Excel 기반 캘리브레이션**: Excel 파일에서 변수와 값을 읽어 자동 적용
- ✅ **실시간 측정 및 로깅**: 지정된 간격으로 측정 데이터를 수집하고 CSV로 저장
- ✅ **자동 메모리 동기화**: ECU 메모리 페이지 자동 동기화
- ✅ **강력한 예외 처리**: 상세한 오류 메시지 및 해결 방법 제시
- ✅ **파일 충돌 방지**: 파일이 사용 중일 경우 자동으로 백업 파일명 생성 (`_1`, `_2`, ...)
- ✅ **상세한 진행 상황**: 실시간으로 진행 상황 표시
- ✅ **클린 코드 아키텍처**: 함수 80라인 이하, 순환복잡도 10 이하

## 📦 필수 요구사항

### 소프트웨어
- **Python**: 3.7 이상
- **INCA**: 실행 중이어야 하며 Experiment가 열려 있어야 함
- **Windows OS**: INCA COM-API는 Windows에서만 동작

### 하드웨어
- INCA 호환 ECU 또는 시뮬레이터 (ETK test device)

## 🔧 설치 방법

### 1. Python 라이브러리 설치

```bash
# 필수 패키지 설치
pip install pywin32 openpyxl

# 또는 requirements.txt 사용
pip install -r requirements.txt
```

#### requirements.txt 내용
```txt
pywin32>=305
openpyxl>=3.1.2
```

### 2. 저장소 클론

```bash
git clone https://github.com/your-username/inca-automation.git
cd inca-automation
```

## 📖 사용 방법

### 기본 사용법

```bash
python inca_refactored.py \
    --calib sample_input.xlsx \
    --measure "Input_1,Input_2,Output" \
    --duration 10 \
    --interval 0.2 \
    --output output.csv
```

### 명령줄 옵션

| 옵션 | 단축 | 필수 | 설명 | 예시 |
|------|------|------|------|------|
| `--calib` | `-c` | ✅ | 캘리브레이션 Excel 파일 경로 | `sample_input.xlsx` |
| `--measure` | `-m` | ✅ | 측정할 변수명 (콤마로 구분) | `"Input_1,Output"` |
| `--duration` | `-d` | ✅ | 측정 시간 (초) | `10` |
| `--interval` | `-i` | ✅ | 샘플링 간격 (초) | `0.2` |
| `--output` | `-o` | ✅ | 결과 CSV 파일명 | `output.csv` |
| `--project` | `-p` | ❌ | INCA 프로젝트 이름 | `Demo3` (기본값) |
| `--version` | `-v` | ❌ | 버전 정보 표시 | - |

### 사용 예시

#### 예시 1: 기본 측정
```bash
python inca_refactored.py \
    -c sample_input.xlsx \
    -m "Input_1,Input_2,Output" \
    -d 10 \
    -i 0.2 \
    -o output.csv
```

#### 예시 2: 빠른 샘플링
```bash
python inca_refactored.py \
    -c sample_input.xlsx \
    -m "Input_1,Output" \
    -d 5 \
    -i 0.1 \
    -o fast_sampling.csv
```

#### 예시 3: 여러 변수 측정
```bash
python inca_refactored.py \
    -c sample_input.xlsx \
    -m "Input_1,Input_2,Output,B_RED,B_GREEN,B_YELLOW" \
    -d 60 \
    -i 1.0 \
    -o long_test.csv
```

## 📄 입력 파일 형식

### Excel 파일 (`sample_input.xlsx`)

Excel 파일은 다음과 같은 형식이어야 합니다:

| 변수명 | 값 |
|--------|-----|
| DEMO_CONSTANT_1 | 100 |
| DEMO_CONSTANT_2 | 75.5 |
| DEMO_CONSTANT_3 | 50.0 |

#### 규칙
- **첫 번째 행**: 헤더 (자동으로 건너뜀)
- **첫 번째 열**: 변수명 (문자열)
- **두 번째 열**: 값 (숫자)
- **빈 행**: 데이터 끝을 표시

#### Excel 파일 예시 생성

```python
import openpyxl

# 새 워크북 생성
wb = openpyxl.Workbook()
ws = wb.active

# 헤더
ws['A1'] = '변수명'
ws['B1'] = '값'

# 데이터
ws['A2'] = 'DEMO_CONSTANT_1'
ws['B2'] = 100
ws['A3'] = 'DEMO_CONSTANT_2'
ws['B3'] = 75.5

# 저장
wb.save('sample_input.xlsx')
```

## 📊 출력 파일 형식

### CSV 파일 (`output.csv`)

측정 결과는 다음과 같은 형식의 CSV 파일로 저장됩니다:

```csv
시간(초),타임스탬프,Input_1,Input_2,Output
0.2,2025-10-22 01:00:00.123,1200.50,2.45,45.20
0.4,2025-10-22 01:00:00.323,1210.30,2.46,45.80
0.6,2025-10-22 01:00:00.523,1220.10,2.47,46.40
```

#### 컬럼 설명
- **시간(초)**: 경과 시간 (interval의 배수)
- **타임스탬프**: 측정 시각 (밀리초 포함)
- **변수명**: 각 측정 변수의 값 (측정 실패 시 `N/A`)

## 🏗️ 프로젝트 구조

```
inca-automation/
│
├── inca_refactored.py          # 메인 스크립트 (리팩토링 버전)
├── sample_input.xlsx            # 샘플 입력 파일
├── output.csv                   # 샘플 출력 파일
├── requirements.txt             # Python 의존성
├── README.md                    # 프로젝트 문서
├── LICENSE                      # 라이선스 파일
│
└── docs/                        # 추가 문서
    ├── architecture.md          # 아키텍처 설명
    ├── api_reference.md         # API 레퍼런스
    └── troubleshooting.md       # 문제 해결 가이드
```

### 코드 구조

```
inca_refactored.py
│
├── 유틸리티 함수
│   ├── print_section_header()       # 섹션 헤더 출력
│   ├── print_error()                # 오류 메시지 출력
│   ├── print_success()              # 성공 메시지 출력
│   └── print_warning()              # 경고 메시지 출력
│
├── FileValidator                    # 파일 검증 클래스
│   ├── validate_file_exists()       # 파일 존재 확인
│   ├── validate_file_readable()     # 읽기 권한 확인
│   ├── is_file_writable()           # 쓰기 권한 확인
│   └── get_available_filename()     # 사용 가능한 파일명 생성
│
├── ExcelCalibrationLoader           # Excel 로더 클래스
│   ├── load()                       # Excel 파일 로드
│   ├── _parse_excel()               # Excel 파싱
│   ├── _parse_row()                 # 행 파싱
│   └── _handle_*_error()            # 오류 처리
│
├── INCADemoController               # INCA 컨트롤러 클래스
│   ├── set_measurement_vars()       # 측정 변수 설정
│   ├── connect_to_inca()            # INCA 연결
│   ├── attach_to_experiment()       # Experiment 연결
│   ├── start_measurement()          # 측정 시작
│   ├── stop_measurement()           # 측정 중지
│   └── disconnect()                 # 연결 해제
│
├── CalibrationApplicator            # 캘리브레이션 적용 클래스
│   ├── apply_all()                  # 전체 적용
│   ├── _apply_single()              # 단일 변수 적용
│   ├── _verify_calibration()        # 검증
│   ├── _sync_memory()               # 메모리 동기화
│   └── _print_summary()             # 결과 요약
│
├── MeasurementCollector             # 측정 수집 클래스
│   ├── collect_and_save()           # 수집 및 저장
│   ├── _collect_samples()           # 샘플 수집
│   ├── _check_connections()         # 연결 확인
│   ├── _read_measurement()          # 측정값 읽기
│   └── _read_all_measurements()     # 전체 측정값 읽기
│
└── main()                           # 메인 실행 함수
```

### 코드 품질 지표

- ✅ **함수 라인 수**: 최대 70라인 (목표: 80라인 이하)
- ✅ **순환복잡도**: 최대 9 (목표: 10 이하)
- ✅ **주석 스타일**: Javadoc (`@param`, `@return`)
- ✅ **타입 힌트**: 모든 함수에 적용
- ✅ **디자인 패턴**: SRP, 의존성 주입

## 🎯 고급 기능

### 1. 자동 파일 백업

파일이 이미 열려있을 경우 자동으로 백업 파일명을 생성합니다:

```
output.csv      → 사용 중
output_1.csv    → 사용 중
output_2.csv    → 생성 ✓
```

### 2. 실시간 진행 상황 표시

```
================================================================================
실시간 측정 및 CSV 저장
================================================================================
측정 설정:
  - 측정 시간: 10초
  - 샘플링 간격: 0.2초
  - 예상 샘플 수: 50개
  - 출력 파일: output.csv
  - 시작 시각: 2025-10-22 01:00:00

측정 연결 확인 중...
  장치: ETK test device:1
  변수 개수: 3개
  ✓ 모든 측정 변수 연결 확인 완료

================================================================================
측정 시작
================================================================================
   시간(초)         Input_1         Input_2          Output
--------------------------------------------------------------------------------
     0.2s          1200.50            2.45           45.20
     0.4s          1210.30            2.46           45.80
```

### 3. 상세한 오류 메시지

```
✗ 오류: Excel 파일을 찾을 수 없습니다!
  전체 경로: C:\Projects\sample_input.xlsx

해결 방법:
  1. 파일 경로가 올바른지 확인하세요
  2. 파일 이름의 철자를 확인하세요
  3. 현재 디렉토리: C:\Projects
```

### 4. 캘리브레이션 검증

```
[1/2] DEMO_CONSTANT_1
────────────────────────────────────────────────────────────────
  현재 값 읽기 중...
  현재 값: 70.00
  새 값 설정: 100.00
  값 쓰기 시도 중...
  
  메모리 페이지 동기화 시도 중...
  ✓ Working Page 다운로드 완료
  
  ✓ 값 쓰기 완료
  검증 중...
  ✓ 확인 완료: 100.00 (변경 성공!)
```

## 🔧 문제 해결

### INCA 연결 오류

**증상**: `✗ 오류: INCA 연결 실패!`

**해결 방법**:
1. INCA가 실행 중인지 확인
2. `pip install pywin32` 실행
3. 다른 프로그램이 INCA COM-API를 사용 중인지 확인
4. INCA 재시작

### Experiment 없음 오류

**증상**: `✗ 오류: 열려 있는 Experiment가 없습니다!`

**해결 방법**:
1. INCA에서 Experiment 열기
2. Database → Workspace → Experiment 순서 확인
3. 프로젝트가 올바르게 로드되었는지 확인

### 측정값 읽기 실패

**증상**: 모든 측정값이 `N/A`로 표시

**해결 방법**:
1. ECU 시뮬레이터 실행 확인
2. 실제 하드웨어 연결 확인
3. 변수명이 A2L 파일과 일치하는지 확인
4. 측정이 시작되었는지 확인 (INCA GUI)

### CSV 파일 생성 실패

**증상**: `✗ 오류: CSV 파일을 생성할 수 없습니다!`

**해결 방법**:
1. 파일이 Excel이나 다른 프로그램에서 열려있는지 확인
2. 쓰기 권한 확인
3. 디렉토리가 존재하는지 확인

### Excel 파일 읽기 실패

**증상**: `✗ 오류: Excel 파일 읽기 실패!`

**해결 방법**:
1. Excel 파일이 열려있다면 닫기
2. 파일 형식 확인 (첫 행: 헤더, 첫 열: 변수명, 두 번째 열: 값)
3. `openpyxl` 재설치: `pip install --upgrade openpyxl`

## 🤝 기여 방법

기여를 환영합니다! 다음 절차를 따라주세요:

1. 저장소 Fork
2. 기능 브랜치 생성 (`git checkout -b feature/amazing-feature`)
3. 변경 사항 커밋 (`git commit -m 'Add amazing feature'`)
4. 브랜치에 Push (`git push origin feature/amazing-feature`)
5. Pull Request 생성

### 코딩 스타일

- 함수 라인 수: 80라인 이하
- 순환복잡도: 10 이하
- Javadoc 스타일 주석 사용
- Type hints 사용
- PEP 8 준수

## 📝 변경 로그

### v2.0.0 (2025-10-22)
- 🎨 전체 코드 리팩토링
- ✨ 클래스 기반 아키텍처로 전환
- 📝 Javadoc 주석 추가
- 🏷️ Type hints 추가
- ✅ 함수 80라인 이하, 순환복잡도 10 이하 달성

### v1.0.0 (2025-10-21)
- 🎉 초기 릴리스
- ✨ Excel 기반 캘리브레이션
- ✨ 실시간 측정 및 CSV 저장
- ✨ 자동 파일 백업

## 📄 라이선스

이 프로젝트는 MIT 라이선스 하에 배포됩니다. 자세한 내용은 [LICENSE](LICENSE) 파일을 참조하세요.

## 👥 개발자

- **Synetics Co., ltd.** - 초기 작업 및 유지보수

## 📞 연락처

질문이나 제안사항이 있으시면 다음으로 연락주세요:

- 📧 Email: dongjoon.han@synetics.kr
- 🐛 Issues: [GitHub Issues](https://github.com/SyneticsCorp/inca_automation/issues)
- 💬 Discussions: [GitHub Discussions](https://github.com/SyneticsCorp/inca_automation/discussions)

---

**⭐ 이 프로젝트가 도움이 되셨다면 Star를 눌러주세요!**
