# GitHub Actions 빠른 시작

## 1️⃣ GitHub 저장소 생성

1. GitHub에 로그인
2. 새 저장소 생성 (예: `easy-match`)
3. Public 또는 Private 선택

## 2️⃣ 코드 업로드

```bash
cd /Users/heebonpark/Downloads/매칭프로그램/ExcelMatcher_MultiPlatform_4.8.1

# Git 초기화
git init
git add .
git commit -m "Initial commit with GitHub Actions"

# 원격 저장소 연결 (YOUR_USERNAME을 본인 계정으로 변경)
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/easy-match.git
git push -u origin main
```

## 3️⃣ 자동 빌드 시작!

코드가 푸시되면 자동으로 빌드가 시작됩니다.

### 빌드 확인
1. GitHub 저장소 → **Actions** 탭
2. "Build Multi-Platform Executables" 워크플로우 확인
3. 완료되면 **Artifacts**에서 다운로드

## 4️⃣ 결과물 다운로드

- **Windows**: `EasyMatch-Windows` → `EasyMatch_v1.0.exe`
- **macOS**: `EasyMatch-macOS` → `EasyMatch-macOS.zip`

## 🎉 완료!

이제 코드를 수정하고 푸시할 때마다 자동으로 빌드됩니다!

---

자세한 내용은 `github_actions_guide.md`를 참조하세요.
