# Easy Match - 엑셀 데이터 매칭 프로그램

[![Build Multi-Platform Executables](https://github.com/bough38-web/easy-match/actions/workflows/build.yml/badge.svg)](https://github.com/bough38-web/easy-match/actions/workflows/build.yml)
![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS-blue)
![License](https://img.shields.io/badge/license-Commercial-orange)

**엑셀과 CSV를 하나로, 클릭 한 번으로 끝나는 데이터 매칭**

## 📥 다운로드

### 자동 빌드 (GitHub Actions)

코드가 업데이트되면 자동으로 실행 파일이 생성됩니다:

1. **GitHub 저장소** → **Actions** 탭
2. 최신 워크플로우 클릭
3. **Artifacts** 섹션에서 다운로드:
   - `EasyMatch-Windows` - Windows EXE 파일
   - `EasyMatch-macOS` - macOS 앱 (ZIP)

### 릴리즈 다운로드

[Releases](../../releases) 페이지에서 최신 버전을 다운로드하세요.

## ✨ 주요 기능

- 🔄 **자동 매칭**: 키 컬럼 기반 데이터 자동 매칭
- 📊 **다중 포맷 지원**: Excel (xlsx, xls), CSV 파일 지원
- 🎨 **직관적 UI**: 드래그 앤 드롭으로 간편한 파일 선택
- ⚡ **빠른 처리**: 대용량 데이터도 빠르게 처리
- 💾 **프리셋 저장**: 자주 사용하는 설정 저장 및 불러오기
- 🔍 **퍼지 매칭**: 유사한 텍스트 자동 매칭 (선택사항)
- 🎯 **컬럼 선택**: 필요한 컬럼만 선택해서 출력

## 🚀 사용 방법

1. **파일 선택**: 기준 파일과 매칭할 파일 선택
2. **키 컬럼 설정**: 매칭 기준이 될 컬럼 선택
3. **출력 컬럼 선택**: 결과에 포함할 컬럼 선택
4. **매칭 실행**: 버튼 클릭으로 즉시 실행
5. **결과 확인**: `outputs` 폴더에 결과 파일 생성

## 💰 가격

- **개인용 (Personal)**:
  - 1년 이용권: **33,000원**
  - 평생 이용권: **132,000원** (한 번 구매로 영구 사용)
  - 최대 처리 행 수: **1,000,000행** (기존 5만행에서 대폭 상향)

- **기업용 (Enterprise)**:
  - 영구 이용권: **180,000원** (PC 1대 기준)
  - 최대 처리 행 수: **무제한**
  - 우선 기술 지원 및 커스터마이징 문의 가능

## 📧 문의

- **커스터마이징**: bough38@gmail.com
- **후원 계좌**: 대구은행 508-14-202118-7 (이현주)

## 🛠️ 개발자용

### 요구사항

- Python 3.8 이상
- 필수 라이브러리: `pip install -r requirements.txt`

### 실행

```bash
python main.py
```

### 빌드

#### Windows
```bash
build_windows.bat
```

#### macOS
```bash
./build_mac.sh
```

#### GitHub Actions (자동)
코드를 푸시하면 자동으로 빌드됩니다. 자세한 내용은 `GITHUB_ACTIONS_SETUP.md`를 참조하세요.

## 📝 라이선스

Commercial License - 상업용 소프트웨어

---

Made by 세은아빠 ❤️
