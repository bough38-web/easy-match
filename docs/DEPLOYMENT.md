# 배포 가이드 (Windows Only)

## 로컬 빌드
```powershell
.\build\build_win.ps1
.\build\release_pack_win.ps1 -Version "4.7.1-win"
```

## GitHub Actions 자동 릴리즈(권장)
1) 태그 생성 후 푸시:
```powershell
git tag v4.7.1-win
git push origin v4.7.1-win
```

2) Actions가 Windows 빌드 → ZIP 생성 → GitHub Release 업로드까지 자동 수행합니다.

## 코드서명(권장)
- Windows: `build/sign_win.ps1`
