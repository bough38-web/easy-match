# Commercial configuration
APP_EDITION = "Commercial"
COMMERCIAL_VERSION = "1.0.0"
from config import LICENSE_FILE
# Admin password to open license panel
ADMIN_PASSWORD = "admin"  # 관리자 비밀번호 (배포 시 변경 권장)

# --- Footer Info ---
CREATOR_NAME = "세은아빠"
SALES_INFO = "판매: 개인 / 기업"
BANK_INFO = "입금계좌/예금주 : 카카오뱅크 3333-03-9648364 박희본"
DISCORD_WEBHOOK_URL = "https://discord.com/api/webhooks/1473080276446871837/LOA5SAX4iTbwQbW5Z-Ec7gXNF5b5kp5XfAiW_3wda5HX8r-YEWkkPIEKbXngab58lOHz"  # 디스코드 웹후크 URL을 여기에 입력하세요
PRICE_INFO = "개인: 1년(3.3만)/평생(13.2만) | 기업: 영구(18만)"
CONTACT_INFO = "커스터마이징 문의: bough38@gmail.com"
TRIAL_EMAIL = "bough38@gmail.com"
TRIAL_SUBJECT = "[ExcelMatcher] 한달 무료 체험 신청합니다"

# --- Donation/Support Message ---
SUPPORT_MESSAGE = "이 프로그램이 도움이 되셨다면 후원으로 응원해주세요!\n개발자의 커피값이 됩니다."
DONATION_CTA = "후원하기 (계좌 확인)"  # Call-to-action for donation

# Personal plan limits (Increased for developer/commercial testing)
PERSONAL_MAX_ROWS = 1000000

# --- Security / Remote Block (Kill Switch) ---
# 분실/도난/불법복제 기기를 원격으로 차단하는 URL입니다.
# GitHub Gist 등에 blacklist.json (예: ["9B3B04F2D13FFED1"])을 올리고 Raw URL을 입력하세요.
BLACKLIST_URL = "https://raw.githubusercontent.com/username/repo/main/blacklist.json" # <-- 여기에 실제 URL 입력 필수
