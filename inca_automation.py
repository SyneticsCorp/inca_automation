"""
INCA COM-API DEMO 프로젝트 제어 스크립트
(리팩토링 버전 - 함수 80라인 이하, 순환복잡도 10 이하)

주요 기능:
1. Excel 파일에서 캘리브레이션 변수 및 값 읽기
2. 표준 명령줄 옵션으로 측정 설정
3. 측정 결과를 화면과 CSV 파일로 동시 출력
4. 자동 메모리 동기화 및 측정 중지
5. 강화된 예외 처리 및 사용자 친화적 오류 메시지
6. 파일 충돌 시 자동 백업 파일명 생성 (_1, _2, ...)
7. 상세한 진행 상황 메시지 및 UX 개선

사용 방법:
python inca_refactored.py -c calib.xlsx -m "Input_1,Input_2,Output" -d 10 -i 0.2 -o result.csv

필수 요구사항:
- Python 3.7 이상
- pip install pywin32 openpyxl

작성자: INCA Automation Team
버전: 2.0
날짜: 2025-10-22
"""

import win32com.client
import time
import sys
import os
import argparse
from datetime import datetime
import csv
from typing import Dict, List, Optional, Tuple

# Excel 파일 읽기를 위한 라이브러리
try:
    import openpyxl
except ImportError:
    print("✗ 오류: openpyxl이 설치되지 않았습니다.")
    print("  해결 방법: pip install openpyxl")
    sys.exit(1)


# ============================================================================
# 유틸리티 함수
# ============================================================================

def print_section_header(title: str, char: str = '=') -> None:
    """
    섹션 헤더를 출력합니다.
    
    @param title: 헤더 제목
    @param char: 구분선 문자 (기본값: '=')
    @return: None
    """
    print(f"\n{char * 80}")
    print(title)
    print(char * 80)


def print_error(message: str, solutions: List[str] = None) -> None:
    """
    오류 메시지와 해결 방법을 출력합니다.
    
    @param message: 오류 메시지
    @param solutions: 해결 방법 리스트
    @return: None
    """
    print(f"\n✗ 오류: {message}")
    if solutions:
        print("\n해결 방법:")
        for idx, solution in enumerate(solutions, 1):
            print(f"  {idx}. {solution}")


def print_success(message: str) -> None:
    """
    성공 메시지를 출력합니다.
    
    @param message: 성공 메시지
    @return: None
    """
    print(f"✓ {message}")


def print_warning(message: str) -> None:
    """
    경고 메시지를 출력합니다.
    
    @param message: 경고 메시지
    @return: None
    """
    print(f"⚠ {message}")


# ============================================================================
# 파일 검증 유틸리티
# ============================================================================

class FileValidator:
    """파일 검증을 위한 유틸리티 클래스"""
    
    @staticmethod
    def validate_file_exists(filepath: str) -> bool:
        """
        파일 존재 여부를 확인합니다.
        
        @param filepath: 확인할 파일 경로
        @return: 파일이 존재하면 True, 아니면 False
        """
        if not os.path.exists(filepath):
            print_error(
                "Excel 파일을 찾을 수 없습니다!",
                [
                    f"파일 경로가 올바른지 확인하세요: {os.path.abspath(filepath)}",
                    "파일 이름의 철자를 확인하세요",
                    f"현재 디렉토리: {os.getcwd()}"
                ]
            )
            return False
        print_success("파일 존재 확인 완료")
        return True
    
    @staticmethod
    def validate_file_readable(filepath: str) -> bool:
        """
        파일 읽기 권한을 확인합니다.
        
        @param filepath: 확인할 파일 경로
        @return: 읽기 가능하면 True, 아니면 False
        """
        if not os.access(filepath, os.R_OK):
            print_error(
                "Excel 파일을 읽을 수 없습니다!",
                [
                    "파일이 다른 프로그램에서 열려있는지 확인하세요",
                    "파일 읽기 권한을 확인하세요"
                ]
            )
            return False
        print_success("파일 접근 권한 확인 완료")
        return True
    
    @staticmethod
    def is_file_writable(filename: str) -> bool:
        """
        파일이 쓰기 가능한지 확인합니다.
        
        @param filename: 확인할 파일명
        @return: 쓰기 가능하면 True, 아니면 False
        """
        if not os.path.exists(filename):
            try:
                directory = os.path.dirname(filename)
                if directory and not os.path.exists(directory):
                    return False
                
                with open(filename, 'w') as f:
                    pass
                os.remove(filename)
                return True
            except:
                return False
        
        try:
            with open(filename, 'a'):
                pass
            return True
        except:
            return False
    
    @staticmethod
    def get_available_filename(filename: str) -> Optional[str]:
        """
        사용 가능한 파일명을 생성합니다 (충돌 시 _1, _2 등 추가).
        
        @param filename: 원본 파일명
        @return: 사용 가능한 파일명 또는 None
        """
        print(f"\n출력 파일 확인 중: {filename}")
        
        if FileValidator.is_file_writable(filename):
            print_success("출력 파일 사용 가능")
            return filename
        
        print_warning("원본 파일을 사용할 수 없습니다 (열려있거나 권한 없음)")
        print("  대체 파일명 검색 중...")
        
        base_name, ext = os.path.splitext(filename)
        
        for counter in range(1, 101):
            new_filename = f"{base_name}_{counter}{ext}"
            if FileValidator.is_file_writable(new_filename):
                print(f"\n✓ 대체 파일명 사용: {new_filename}")
                print("  이유: 원본 파일이 다른 프로그램에서 사용 중")
                return new_filename
        
        print_error(
            "사용 가능한 파일명을 찾을 수 없습니다!",
            [
                f"{base_name}_*.{ext} 파일들을 닫거나 삭제하세요",
                "다른 출력 파일명을 지정하세요"
            ]
        )
        return None


# ============================================================================
# Excel 데이터 로더
# ============================================================================

class ExcelCalibrationLoader:
    """Excel 파일에서 캘리브레이션 데이터를 로드하는 클래스"""
    
    @staticmethod
    def load(excel_path: str) -> Optional[Dict[str, float]]:
        """
        Excel 파일에서 캘리브레이션 변수와 값을 읽습니다.
        
        @param excel_path: Excel 파일 경로
        @return: {변수명: 값} 딕셔너리 또는 None (실패 시)
        """
        print_section_header("Excel 파일 로드")
        print(f"파일 경로: {excel_path}")
        
        if not FileValidator.validate_file_exists(excel_path):
            return None
        
        if not FileValidator.validate_file_readable(excel_path):
            return None
        
        try:
            return ExcelCalibrationLoader._parse_excel(excel_path)
        except PermissionError:
            ExcelCalibrationLoader._handle_permission_error(excel_path)
            return None
        except Exception as e:
            ExcelCalibrationLoader._handle_generic_error(excel_path, e)
            return None
    
    @staticmethod
    def _parse_excel(excel_path: str) -> Dict[str, float]:
        """
        Excel 파일을 파싱합니다.
        
        @param excel_path: Excel 파일 경로
        @return: {변수명: 값} 딕셔너리
        @raises: Exception if parsing fails
        """
        print("\nExcel 데이터 파싱 중...")
        
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
        
        calib_data = {}
        row_count = 0
        skip_count = 0
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            if row[0] is None:
                break
            
            result = ExcelCalibrationLoader._parse_row(row, row_idx)
            if result:
                var_name, var_value = result
                calib_data[var_name] = var_value
                row_count += 1
                print(f"  ✓ {var_name} = {var_value}")
            else:
                skip_count += 1
        
        if row_count == 0:
            ExcelCalibrationLoader._handle_empty_data()
            raise ValueError("No valid calibration data found")
        
        print(f"\n{'=' * 80}")
        print(f"✓ Excel 로드 완료: {row_count}개 변수, {skip_count}개 건너뜀")
        print(f"{'=' * 80}")
        
        return calib_data
    
    @staticmethod
    def _parse_row(row: Tuple, row_idx: int) -> Optional[Tuple[str, float]]:
        """
        Excel 행을 파싱합니다.
        
        @param row: Excel 행 데이터
        @param row_idx: 행 번호
        @return: (변수명, 값) 튜플 또는 None
        """
        var_name = str(row[0]).strip() if row[0] else ""
        var_value = row[1]
        
        if not var_name:
            print(f"  ⚠ {row_idx}행: 변수명이 비어있습니다 (건너뜀)")
            return None
        
        if var_value is None:
            print(f"  ⚠ {row_idx}행 '{var_name}': 값이 비어있습니다 (건너뜀)")
            return None
        
        try:
            return (var_name, float(var_value))
        except (ValueError, TypeError):
            print(f"  ✗ {row_idx}행 '{var_name}': 숫자 변환 실패 '{var_value}' (건너뜀)")
            return None
    
    @staticmethod
    def _handle_permission_error(excel_path: str) -> None:
        """권한 오류를 처리합니다."""
        print_error(
            "Excel 파일에 접근할 수 없습니다!",
            [
                "파일이 Excel에서 열려있다면 닫아주세요",
                "파일이 읽기 전용인지 확인하세요"
            ]
        )
    
    @staticmethod
    def _handle_generic_error(excel_path: str, error: Exception) -> None:
        """일반 오류를 처리합니다."""
        print_error(
            "Excel 파일 읽기 실패!",
            [
                "Excel 파일이 손상되지 않았는지 확인하세요",
                "Excel에서 파일을 열어 정상적으로 열리는지 확인하세요"
            ]
        )
        print(f"  오류 내용: {error}")
        import traceback
        traceback.print_exc()
    
    @staticmethod
    def _handle_empty_data() -> None:
        """빈 데이터 오류를 처리합니다."""
        print_error(
            "Excel 파일에서 유효한 캘리브레이션 데이터를 찾을 수 없습니다!",
            [
                "Excel 파일 형식을 확인하세요:",
                "  | 변수명            | 값     |",
                "  |------------------|--------|",
                "  | DEMO_CONSTANT_1  | 100    |",
                "  | DEMO_CONSTANT_2  | 75.5   |",
                "첫 번째 행은 헤더로 건너뜁니다",
                "변수명과 값이 모두 입력되어 있는지 확인하세요"
            ]
        )


# ============================================================================
# INCA 컨트롤러
# ============================================================================

class INCADemoController:
    """INCA DEMO 프로젝트 자동 제어 클래스"""
    
    def __init__(self, project_name: str = "Demo3"):
        """
        컨트롤러를 초기화합니다.
        
        @param project_name: INCA에서 사용 중인 프로젝트 이름
        """
        self.project_name = project_name
        self.inca = None
        self.experiment = None
        self.device = None
        self.device_name = None
        self.measurement_started = False
        self.measurement_vars = {}
    
    def set_measurement_vars(self, var_names_str: str) -> bool:
        """
        측정 변수를 설정합니다.
        
        @param var_names_str: 측정 변수명 문자열 (콤마로 구분)
        @return: 성공 여부
        """
        print_section_header("측정 변수 설정")
        
        if not var_names_str or not var_names_str.strip():
            print_error(
                "측정 변수가 비어있습니다!",
                ['-m 옵션에 변수명을 입력하세요 (예: -m "Input_1,Input_2,Output")']
            )
            return False
        
        var_names = [v.strip() for v in var_names_str.split(',') if v.strip()]
        
        if not var_names:
            print_error(
                "유효한 측정 변수가 없습니다!",
                ['변수명을 콤마로 구분하여 입력하세요 (예: -m "Input_1,Input_2,Output")']
            )
            return False
        
        print(f"등록할 변수: {len(var_names)}개")
        for idx, var_name in enumerate(var_names, 1):
            key = var_name.lower().replace('_', '')
            self.measurement_vars[key] = var_name
            print(f"  [{idx}] {var_name}")
        
        print_success("측정 변수 설정 완료")
        return True
    
    def connect_to_inca(self) -> bool:
        """
        INCA에 연결합니다.
        
        @return: 성공 여부
        """
        print_section_header("INCA COM-API 연결")
        print("연결 시도 중...")
        
        try:
            self.inca = win32com.client.Dispatch("Inca.Inca")
            print_success("INCA COM-API 연결 성공")
            
            self.inca.WriteToMonitor("Python 자동화 스크립트가 연결되었습니다.")
            print_success("INCA 모니터에 메시지 전송 완료")
            
            return True
        except Exception as e:
            print_error(
                "INCA 연결 실패!",
                [
                    "INCA가 실행 중인지 확인하세요",
                    "pywin32가 설치되어 있는지 확인하세요 (pip install pywin32)",
                    "다른 프로그램이 INCA COM-API를 사용 중이지 않은지 확인하세요",
                    "INCA를 재시작해보세요"
                ]
            )
            print(f"  오류 내용: {e}")
            return False
    
    def attach_to_experiment(self) -> bool:
        """
        현재 열려 있는 Experiment에 연결합니다.
        
        @return: 성공 여부
        """
        print_section_header("Experiment 연결")
        print("현재 열린 Experiment 검색 중...")
        
        try:
            self.experiment = self.inca.GetOpenedExperiment()
            
            if not self.experiment:
                self._handle_no_experiment_error()
                return False
            
            print_success("Experiment 발견")
            return self._connect_to_device()
            
        except Exception as e:
            self._handle_experiment_error(e)
            return False
    
    def _connect_to_device(self) -> bool:
        """
        장치에 연결합니다.
        
        @return: 성공 여부
        """
        print("\n연결된 장치 검색 중...")
        devices = self.experiment.GetAllDevices()
        
        if not devices:
            self._handle_no_device_error()
            return False
        
        self.device = devices[0]
        self.device_name = self._get_device_name()
        
        print_success(f"장치 연결 성공: {self.device_name}")
        
        if len(devices) > 1:
            print(f"  (참고: 총 {len(devices)}개 장치 발견, 첫 번째 장치 사용)")
        
        self._check_simulation_device()
        return True
    
    def _get_device_name(self) -> str:
        """장치 이름을 가져옵니다."""
        try:
            return self.device.GetName()
        except:
            return "Unknown Device"
    
    def _check_simulation_device(self) -> None:
        """시뮬레이션 장치인지 확인하고 경고합니다."""
        device_upper = self.device_name.upper()
        if "ETK" in device_upper or "TEST" in device_upper:
            print_warning(f"{self.device_name}는 시뮬레이션 장치입니다")
            print("  - 실제 ECU가 연결되지 않으면 측정값을 읽을 수 없습니다")
            print("  - 시뮬레이터를 실행하거나 실제 하드웨어를 연결하세요")
    
    def _handle_no_experiment_error(self) -> None:
        """Experiment 없음 오류를 처리합니다."""
        print_error(
            "열려 있는 Experiment가 없습니다!",
            [
                "INCA에서 Experiment를 먼저 열어주세요",
                "Database → Workspace → Experiment 순서로 확인하세요",
                f"프로젝트({self.project_name})가 올바르게 로드되었는지 확인하세요"
            ]
        )
    
    def _handle_no_device_error(self) -> None:
        """장치 없음 오류를 처리합니다."""
        print_error(
            "연결된 장치가 없습니다!",
            [
                "INCA에서 Hardware Configuration을 확인하세요",
                "Device Manager에서 하드웨어를 추가하세요",
                "ETK test device 또는 Virtual Device를 추가하세요",
                "Experiment에 하드웨어 구성을 할당하세요"
            ]
        )
    
    def _handle_experiment_error(self, error: Exception) -> None:
        """Experiment 연결 오류를 처리합니다."""
        print_error(
            "Experiment 연결 실패!",
            [
                "INCA에서 Experiment가 정상적으로 열려있는지 확인하세요",
                "COM-API 권한 문제일 수 있습니다. INCA를 재시작해보세요"
            ]
        )
        print(f"  오류 내용: {error}")
        import traceback
        traceback.print_exc()
    
    def start_measurement(self) -> bool:
        """
        측정을 시작합니다.
        
        @return: 성공 여부
        """
        print_section_header("측정 시작")
        print("측정 시작 명령 전송 중...")
        
        try:
            self.experiment.StartMeasurement()
            self.measurement_started = True
            print_success("측정 시작 성공")
            self.inca.WriteToMonitor("측정이 시작되었습니다.")
            return True
        except Exception as e:
            print_warning(f"측정 시작 실패 - {e}")
            print("  (이미 측정이 시작된 상태일 수 있습니다)")
            return False
    
    def stop_measurement(self) -> bool:
        """
        측정을 중지합니다.
        
        @return: 성공 여부
        """
        if not self.measurement_started:
            return True
        
        print_section_header("측정 중지")
        print("측정 중지 명령 전송 중...")
        
        try:
            self.experiment.StopMeasurement()
            self.measurement_started = False
            print_success("측정 중지 성공")
            self.inca.WriteToMonitor("측정이 중지되었습니다.")
            return True
        except Exception as e:
            print(f"✗ 측정 중지 실패: {e}")
            return False
    
    def disconnect(self) -> None:
        """INCA 연결을 종료합니다."""
        if not self.inca:
            return
        
        try:
            print_section_header("INCA 연결 종료")
            self.inca.WriteToMonitor("Python 자동화 스크립트 종료")
            print_success("INCA 모니터에 종료 메시지 전송")
            
            self.inca.DisconnectFromTool()
            print_success("INCA COM-API 연결 해제")
        except Exception as e:
            print_warning(f"연결 종료 중 오류: {e}")


# ============================================================================
# 캘리브레이션 적용기
# ============================================================================

class CalibrationApplicator:
    """캘리브레이션 변수 적용을 담당하는 클래스"""
    
    def __init__(self, controller: INCADemoController):
        """
        초기화합니다.
        
        @param controller: INCA 컨트롤러 인스턴스
        """
        self.controller = controller
    
    def apply_all(self, calib_dict: Dict[str, float]) -> Tuple[int, int]:
        """
        모든 캘리브레이션 변수를 적용합니다.
        
        @param calib_dict: {변수명: 값} 딕셔너리
        @return: (성공 개수, 실패 개수) 튜플
        """
        print_section_header("캘리브레이션 변수 적용")
        
        success_count = 0
        fail_count = 0
        
        for idx, (var_name, new_value) in enumerate(calib_dict.items(), 1):
            print(f"\n[{idx}/{len(calib_dict)}] {var_name}")
            print("─" * 60)
            
            if self._apply_single(var_name, new_value):
                success_count += 1
            else:
                fail_count += 1
        
        self._print_summary(success_count, fail_count, len(calib_dict))
        return success_count, fail_count
    
    def _apply_single(self, var_name: str, new_value: float) -> bool:
        """
        단일 캘리브레이션 변수를 적용합니다.
        
        @param var_name: 변수명
        @param new_value: 새 값
        @return: 성공 여부
        """
        print("  현재 값 읽기 중...")
        current_val = self._read_calibration(var_name)
        
        if current_val is not None:
            print(f"  현재 값: {current_val:.2f}")
        else:
            print_warning("현재 값을 읽을 수 없습니다")
        
        print(f"  새 값 설정: {new_value:.2f}")
        print("  값 쓰기 시도 중...")
        
        if not self._write_calibration(var_name, new_value):
            return False
        
        print_success("값 쓰기 완료")
        return self._verify_calibration(var_name, new_value)
    
    def _verify_calibration(self, var_name: str, expected_value: float) -> bool:
        """
        캘리브레이션 값을 검증합니다.
        
        @param var_name: 변수명
        @param expected_value: 예상 값
        @return: 성공 여부
        """
        print("  검증 중...")
        time.sleep(0.3)
        
        for retry in range(3):
            verify_val = self._read_calibration(var_name)
            
            if verify_val is not None and abs(verify_val - expected_value) < 0.01:
                print(f"  ✓ 확인 완료: {verify_val:.2f} (변경 성공!)")
                return True
            
            if retry < 2:
                print(f"  ⏳ 재시도 중... ({retry + 1}/3)")
                time.sleep(0.5)
        
        print_warning(f"확인된 값: {verify_val:.2f} (변경이 반영되지 않음)")
        print("  → INCA GUI에서 수동으로 'Download to ECU'를 클릭하세요")
        return False
    
    def _read_calibration(self, var_name: str) -> Optional[float]:
        """캘리브레이션 값을 읽습니다."""
        try:
            calib_obj = self.controller.experiment.GetCalibrationValueInDevice(
                var_name, self.controller.device
            )
            return calib_obj.GetDoublePhysValue()
        except Exception:
            return None
    
    def _write_calibration(self, var_name: str, value: float) -> bool:
        """캘리브레이션 값을 씁니다."""
        try:
            calib_obj = self.controller.experiment.GetCalibrationValueInDevice(
                var_name, self.controller.device
            )
            calib_obj.SetDoublePhysValue(value)
            time.sleep(0.1)
            self._sync_memory()
            return True
        except Exception as e:
            print(f"  ✗ {var_name} 쓰기 실패: {e}")
            return False
    
    def _sync_memory(self) -> bool:
        """메모리 페이지를 동기화합니다."""
        print("\n  메모리 페이지 동기화 시도 중...")
        
        sync_methods = [
            ("Synchronize", lambda: self.controller.experiment.Synchronize()),
            ("DownloadWorkingPage", lambda: self.controller.experiment.DownloadWorkingPage()),
            ("SyncWorkingPageToEcu", lambda: self.controller.experiment.SyncWorkingPageToEcu())
        ]
        
        for method_name, method in sync_methods:
            try:
                method()
                print(f"  ✓ {method_name} 완료")
                return True
            except:
                continue
        
        print_warning("메모리 페이지 동기화 메서드를 찾을 수 없습니다")
        print("  → INCA GUI에서 수동으로 'Download to ECU' 버튼을 클릭하세요")
        return False
    
    def _print_summary(self, success: int, fail: int, total: int) -> None:
        """적용 결과 요약을 출력합니다."""
        print(f"\n{'=' * 80}")
        print("캘리브레이션 적용 결과")
        print(f"{'=' * 80}")
        print(f"  성공: {success}개")
        print(f"  실패: {fail}개")
        print(f"  총계: {total}개")
        
        if fail > 0:
            print_warning("일부 변수 변경이 반영되지 않았습니다")
            print("  → INCA GUI에서 'Download to ECU' 버튼을 수동으로 클릭하세요")


# ============================================================================
# 측정 데이터 수집기
# ============================================================================

class MeasurementCollector:
    """측정 데이터 수집을 담당하는 클래스"""
    
    def __init__(self, controller: INCADemoController):
        """
        초기화합니다.
        
        @param controller: INCA 컨트롤러 인스턴스
        """
        self.controller = controller
    
    def collect_and_save(self, duration: float, interval: float, csv_filename: str) -> None:
        """
        측정 데이터를 수집하고 CSV에 저장합니다.
        
        @param duration: 측정 시간 (초)
        @param interval: 샘플링 간격 (초)
        @param csv_filename: 저장할 CSV 파일명
        """
        print_section_header("실시간 측정 및 CSV 저장")
        
        num_samples = int(duration / interval)
        self._print_settings(duration, interval, num_samples, csv_filename)
        
        var_names_list = list(self.controller.measurement_vars.values())
        connection_ok = self._check_connections(var_names_list)
        
        csv_file = self._open_csv_file(csv_filename, var_names_list)
        
        try:
            self._collect_samples(csv_file, var_names_list, num_samples, interval)
        finally:
            csv_file.close()
        
        print(f"\n✓ 측정 완료")
        print(f"✓ 결과가 CSV 파일에 저장되었습니다: {csv_filename}")
    
    def _print_settings(self, duration: float, interval: float, 
                       num_samples: int, csv_filename: str) -> None:
        """측정 설정을 출력합니다."""
        print("측정 설정:")
        print(f"  - 측정 시간: {duration}초")
        print(f"  - 샘플링 간격: {interval}초")
        print(f"  - 예상 샘플 수: {num_samples}개")
        print(f"  - 출력 파일: {csv_filename}")
        print(f"  - 시작 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    def _check_connections(self, var_names_list: List[str]) -> int:
        """
        측정 연결을 확인합니다.
        
        @param var_names_list: 측정 변수 리스트
        @return: 연결 성공 개수
        """
        print(f"\n측정 연결 확인 중...")
        print(f"  장치: {self.controller.device_name}")
        print(f"  변수 개수: {len(var_names_list)}개")
        
        connection_ok = 0
        for var_name in var_names_list:
            value = self._read_measurement(var_name, verbose=True)
            if value is not None:
                connection_ok += 1
        
        self._print_connection_status(connection_ok, len(var_names_list))
        return connection_ok
    
    def _print_connection_status(self, ok_count: int, total_count: int) -> None:
        """연결 상태를 출력합니다."""
        if ok_count == 0:
            print_warning("모든 측정 변수 읽기 실패!")
            print("  가능한 원인:")
            print("    - ECU 시뮬레이터가 실행되지 않음")
            print("    - 실제 하드웨어가 연결되지 않음")
            print("    - 변수명이 A2L 파일과 일치하지 않음")
            print("  → 측정은 계속되지만 모든 값이 'N/A'로 표시될 수 있습니다")
        elif ok_count < total_count:
            print_warning(f"일부 변수만 읽기 가능 ({ok_count}/{total_count})")
        else:
            print_success("모든 측정 변수 연결 확인 완료")
    
    def _open_csv_file(self, csv_filename: str, var_names_list: List[str]):
        """CSV 파일을 엽니다."""
        print("\nCSV 파일 생성 중...")
        try:
            csv_file = open(csv_filename, 'w', newline='', encoding='utf-8-sig')
            print_success("CSV 파일 생성 성공")
            
            # 헤더 작성
            header = ['시간(초)', '타임스탬프'] + var_names_list
            csv_writer = csv.writer(csv_file)
            csv_writer.writerow(header)
            
            return csv_file
        except Exception as e:
            print_error(f"CSV 파일을 생성할 수 없습니다!", [f"오류 내용: {e}"])
            raise
    
    def _collect_samples(self, csv_file, var_names_list: List[str], 
                        num_samples: int, interval: float) -> None:
        """샘플을 수집합니다."""
        print_section_header("측정 시작")
        
        # 헤더 출력
        header_line = f"{'시간(초)':>8} "
        header_line += ' '.join([f"{v:>15}" for v in var_names_list])
        print(header_line)
        print("-" * 80)
        
        csv_writer = csv.writer(csv_file)
        
        for i in range(num_samples):
            elapsed_time = (i + 1) * interval
            timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
            
            values, display_values = self._read_all_measurements(var_names_list)
            
            # CSV 기록
            row = [f"{elapsed_time:.1f}", timestamp] + values
            csv_writer.writerow(row)
            
            # 화면 출력
            display_line = f"{elapsed_time:>7.1f}s "
            display_line += ' '.join(display_values)
            print(display_line)
            
            time.sleep(interval)
        
        print("-" * 80)
        print(f"종료 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    
    def _read_all_measurements(self, var_names_list: List[str]) -> Tuple[List, List[str]]:
        """
        모든 측정값을 읽습니다.
        
        @param var_names_list: 측정 변수 리스트
        @return: (CSV용 값 리스트, 화면 표시용 문자열 리스트) 튜플
        """
        values = []
        display_values = []
        
        for var_name in var_names_list:
            value = self._read_measurement(var_name, verbose=False)
            
            if value is not None:
                values.append(value)
                display_values.append(f"{value:>15.2f}")
            else:
                values.append('N/A')
                display_values.append(f"{'N/A':>15}")
        
        return values, display_values
    
    def _read_measurement(self, var_name: str, verbose: bool = False) -> Optional[float]:
        """
        측정값을 읽습니다.
        
        @param var_name: 측정 변수명
        @param verbose: 디버그 출력 여부
        @return: 측정값 또는 None
        """
        try:
            measure_obj = self.controller.experiment.GetMeasurementValueInDevice(
                var_name, self.controller.device
            )
            value = measure_obj.GetDoublePhysValue()
            
            if verbose:
                print(f"  디버그: {var_name} = {value}")
            
            return value
        except Exception as e:
            if verbose:
                print(f"  경고: {var_name} 읽기 실패 - {e}")
            return None


# ============================================================================
# 명령줄 인자 파서
# ============================================================================

def parse_arguments() -> argparse.Namespace:
    """
    명령줄 인자를 파싱합니다.
    
    @return: 파싱된 인자 객체
    """
    parser = argparse.ArgumentParser(
        description='INCA 자동화 스크립트 - Excel 기반 캘리브레이션 및 측정',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
사용 예시:
  python %(prog)s -c calib.xlsx -m "Input_1,Input_2,Output" -d 10 -i 0.2 -o result.csv
  python %(prog)s -c calib.xlsx -m "Input_1,Output" -d 5 -i 0.1 -o fast.csv
  python %(prog)s --calib settings.xlsx --measure "B_RED,B_GREEN" --duration 60 --interval 1 --output test.csv

Excel 파일 형식 (첫 행은 헤더):
  | 변수명            | 값     |
  |------------------|--------|
  | DEMO_CONSTANT_1  | 100    |
  | DEMO_CONSTANT_2  | 75.5   |
        """
    )
    
    parser.add_argument('-c', '--calib', required=True, metavar='FILE',
                        help='캘리브레이션 변수가 저장된 Excel 파일 경로')
    parser.add_argument('-m', '--measure', required=True, metavar='VARS',
                        help='측정할 변수명 (콤마로 구분)')
    parser.add_argument('-d', '--duration', required=True, type=float, metavar='SECONDS',
                        help='측정 시간 (초)')
    parser.add_argument('-i', '--interval', required=True, type=float, metavar='SECONDS',
                        help='샘플링 간격 (초)')
    parser.add_argument('-o', '--output', required=True, metavar='FILE',
                        help='결과를 저장할 CSV 파일명')
    parser.add_argument('-p', '--project', default='Demo3', metavar='NAME',
                        help='INCA 프로젝트 이름 (기본값: Demo3)')
    parser.add_argument('-v', '--version', action='version', version='%(prog)s 2.0')
    
    return parser.parse_args()


# ============================================================================
# 메인 실행 함수
# ============================================================================

def print_script_header(args: argparse.Namespace) -> None:
    """
    스크립트 헤더를 출력합니다.
    
    @param args: 명령줄 인자
    """
    print("=" * 80)
    print("INCA DEMO 프로젝트 자동 제어 스크립트 (리팩토링 버전)")
    print("=" * 80)
    print("\n[실행 설정]")
    print(f"  프로젝트 이름: {args.project}")
    print(f"  캘리브레이션 파일: {args.calib}")
    print(f"  측정 변수: {args.measure}")
    print(f"  측정 시간: {args.duration}초")
    print(f"  샘플링 간격: {args.interval}초")
    print(f"  예상 샘플 수: {int(args.duration / args.interval)}개")
    print(f"  출력 CSV 파일: {args.output}")
    print(f"  실행 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")


def wait_for_stabilization() -> None:
    """데이터 안정화를 위해 대기합니다."""
    print("\n데이터 안정화 대기 중...")
    for i in range(3, 0, -1):
        print(f"  {i}초...")
        time.sleep(1)
    print_success("준비 완료")


def print_completion_summary(output_filename: str) -> None:
    """
    완료 요약을 출력합니다.
    
    @param output_filename: 출력 파일명
    """
    print_section_header("✓ 모든 작업 완료!")
    print("\n[결과 정보]")
    print(f"  출력 파일: {output_filename}")
    print(f"  완료 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("\n스크립트 실행이 성공적으로 완료되었습니다.")


def main():
    """
    메인 실행 함수입니다.
    """
    args = parse_arguments()
    print_script_header(args)
    
    controller = INCADemoController(project_name=args.project)
    
    try:
        # 1. Excel 로드
        calib_data = ExcelCalibrationLoader.load(args.calib)
        if not calib_data:
            print_section_header("✗ 스크립트 종료: Excel 파일 로드 실패")
            return
        
        # 2. 측정 변수 설정
        if not controller.set_measurement_vars(args.measure):
            print_section_header("✗ 스크립트 종료: 측정 변수 설정 실패")
            return
        
        # 3. 출력 파일 확인
        output_filename = FileValidator.get_available_filename(args.output)
        if not output_filename:
            print_section_header("✗ 스크립트 종료: 출력 파일 생성 불가")
            return
        
        # 4. INCA 연결
        if not controller.connect_to_inca():
            print_section_header("✗ 스크립트 종료: INCA 연결 실패")
            return
        
        # 5. Experiment 연결
        if not controller.attach_to_experiment():
            print_section_header("✗ 스크립트 종료: Experiment 연결 실패")
            return
        
        # 6. 측정 시작
        measurement_started = controller.start_measurement()
        wait_for_stabilization()
        
        # 7. 캘리브레이션 적용
        applicator = CalibrationApplicator(controller)
        applicator.apply_all(calib_data)
        
        # 8. 측정 및 저장
        collector = MeasurementCollector(controller)
        collector.collect_and_save(args.duration, args.interval, output_filename)
        
        # 9. 측정 중지
        if measurement_started:
            controller.stop_measurement()
        
        print_completion_summary(output_filename)
        
    except KeyboardInterrupt:
        print_section_header("⚠ 사용자에 의해 중단되었습니다")
        if controller.measurement_started:
            print("\n측정 중지 중...")
            controller.stop_measurement()
    
    except Exception as e:
        print_section_header("✗ 예상치 못한 오류 발생!")
        print(f"오류 내용: {e}")
        print("\n상세 정보:")
        import traceback
        traceback.print_exc()
        
        if controller.measurement_started:
            print("\n측정 중지 시도 중...")
            controller.stop_measurement()
    
    finally:
        controller.disconnect()


if __name__ == "__main__":
    main()
