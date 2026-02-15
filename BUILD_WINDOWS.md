# Windows EXE 빌드 - 빠른 시작

## Windows 컴퓨터에서 실행하세요

### 1단계: Python 설치 확인
```bash
python --version
```
Python 3.8 이상이 필요합니다.

### 2단계: 필수 라이브러리 설치
```bash
pip install -r requirements.txt
pip install pyinstaller
```

### 3단계: EXE 빌드
```bash
build_windows.bat
```

### 결과
`dist/EasyMatch_v1.0.exe` 파일이 생성됩니다!

---

## ⚠️ 중요: macOS에서는 불가능

현재 macOS를 사용 중이시므로 Windows EXE를 직접 빌드할 수 없습니다.

**해결 방법:**
1. Windows 컴퓨터에서 위 스크립트 실행
2. Windows 가상머신 사용 (Parallels, VMware)
3. GitHub Actions 사용 (자동 빌드)

자세한 내용은 `windows_build_guide.md`를 참조하세요.
