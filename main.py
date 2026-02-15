from __future__ import annotations

import os
import sys
import traceback

# -----------------------------
# 터미널 로그 유틸
# -----------------------------
def eprint(*args):
    print(*args, file=sys.stderr, flush=True)

def _msgbox(title: str, message: str, kind: str = "info"):
    """
    macOS에서 messagebox가 parent 없이 호출될 때 튕기는 경우가 있어
    임시 Tk root를 만들어 안전하게 띄웁니다.
    """
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        if kind == "error":
            messagebox.showerror(title, message, parent=root)
        elif kind == "warning":
            messagebox.showwarning(title, message, parent=root)
        else:
            messagebox.showinfo(title, message, parent=root)
        root.destroy()
    except Exception:
        # GUI도 안 뜨는 상황이면 stderr로라도 남김
        eprint(f"[{title}] {message}")

# -----------------------------
# 라이선스 결과 파싱(2개/3개/딕셔너리 등 가변 대응)
# -----------------------------
def _normalize_license_result(res):
    """
    license_manager 쪽 함수가 반환 형식을 바꿔도 절대 안 죽게 하는 래퍼.
    반환: (ok: bool, info: dict|str)
    """
    if isinstance(res, tuple):
        if len(res) >= 2:
            return bool(res[0]), res[1]
        if len(res) == 1:
            return bool(res[0]), {}
        return False, {}
    if isinstance(res, dict):
        # {"ok": True, "info": {...}} 같은 케이스
        ok = bool(res.get("ok", False))
        info = res.get("info", res)
        return ok, info
    if isinstance(res, bool):
        return res, {}
    # 그 외
    return False, res

# -----------------------------
# 엔트리
# -----------------------------
def main():
    eprint("[ExcelMatcher] starting...")

    # 1) 라이선스 관련 모듈 로드
    try:
        import license_manager
    except Exception as e:
        eprint("license_manager import failed:", e)
        traceback.print_exc()
        _msgbox("오류", f"license_manager 로드 실패:\n{e}", "error")
        return 2

    # 2) 라이선스 확인/등록 플로우
    #    프로젝트마다 함수명이 다를 수 있어서 가능한 이름들을 순차 시도합니다.
    try:
        ok = False
        info = {}

        # (A) validate_license()가 있다면 먼저 검사
        if hasattr(license_manager, "validate_license"):
            res = license_manager.validate_license()
            ok, info = _normalize_license_result(res)

        # (B) 라이선스가 없거나 실패면 등록/생성 UI 실행 (가능한 함수명 자동 탐색)
        if not ok:
            eprint("[ExcelMatcher] license not valid -> running license flow...")

            flow_fn = None
            for fn_name in [
                "ensure_license", "ensure_license_flow", "register_license",
                "run_license_flow", "create_or_register_license",
            ]:
                if hasattr(license_manager, fn_name):
                    flow_fn = getattr(license_manager, fn_name)
                    break

            if flow_fn is None:
                # 최소한 "라이선스 파일 없음"을 사용자에게 알림
                _msgbox("라이선스 오류", "라이선스가 없거나 유효하지 않습니다.\n(license_manager에 등록/생성 함수가 없습니다)", "error")
                return 3

            res = flow_fn()
            ok, info = _normalize_license_result(res)

        if not ok:
            _msgbox("라이선스 오류", f"라이선스 확인 실패:\n{info}", "error")
            return 4

        # info를 App이 기대하는 형태로 정리
        license_info = info if isinstance(info, dict) else {"type": "unknown", "expiry": "-", "raw": str(info)}
        if "type" not in license_info:
            license_info["type"] = "personal"
        if "expiry" not in license_info:
            license_info["expiry"] = "-"

        eprint(f"[ExcelMatcher] license OK: {license_info}")

    except Exception as e:
        eprint("license flow error:", e)
        traceback.print_exc()
        _msgbox("오류", f"라이선스 처리 중 오류:\n{e}", "error")
        return 5

    # 3) UI 실행 (여기서 mainloop가 떠 있어야 정상)
    try:
        from ui import App
        app = App(license_info=license_info)
        eprint("[ExcelMatcher] UI launching... (mainloop)")
        app.mainloop()
        eprint("[ExcelMatcher] UI closed.")
        return 0
    except Exception as e:
        eprint("UI failed:", e)
        traceback.print_exc()
        _msgbox("오류", f"UI 실행 실패:\n{e}", "error")
        return 6


if __name__ == "__main__":
    raise SystemExit(main())