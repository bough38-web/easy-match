from __future__ import annotations
import platform
import importlib.util
from dataclasses import dataclass

@dataclass
class DiagResult:
    ok: bool
    title: str
    detail: str

def check_python() -> DiagResult:
    import sys
    return DiagResult(True, "Python", f"{sys.version}")

def check_xlwings_import() -> DiagResult:
    spec = importlib.util.find_spec("xlwings")
    if spec is None:
        return DiagResult(False, "xlwings 미설치", "열려있는 엑셀 모드를 쓰려면 xlwings 설치가 필요합니다.")
    return DiagResult(True, "xlwings", "설치됨")

def check_excel_hint() -> DiagResult:
    osname = platform.system()
    if osname == "Windows":
        return DiagResult(True, "Excel", "Windows에서는 Microsoft Excel 설치가 필요합니다(열려있는 엑셀 모드).")
    if osname == "Darwin":
        return DiagResult(True, "Excel", "macOS에서는 Microsoft Excel 설치 및 Automation(Apple Events) 권한 허용이 필요할 수 있습니다.")
    return DiagResult(True, "Excel", f"{osname}에서는 파일 모드 사용을 권장합니다.")

def collect_summary() -> list[DiagResult]:
    return [check_python(), check_xlwings_import(), check_excel_hint()]

def format_summary(results: list[DiagResult]) -> str:
    lines = []
    for r in results:
        icon = "[OK]" if r.ok else "[WARN]"
        lines.append(f"{icon} {r.title}: {r.detail}")
    return "\n".join(lines)
