import socket
try:
    import requests
except ImportError:
    requests = None
import threading
import json
import logging
import platform

def report_usage_status(webhook_url):
    """
    Reports the usage status (Computer Name, IP-based Location) to Discord via Webhook.
    Runs in a separate thread to prevent UI lag.
    """
    if not webhook_url:
        logging.info("Discord Webhook URL not provided. Skipping status report.")
        return

    if requests is None:
        logging.error("Requests module not found. Skipping status report.")
        return

    def _task():
        try:
            # 1. Get Hostname & Platform
            hostname = socket.gethostname()
            os_info = f"{platform.system()} {platform.release()}"
            
            # 2. Get Public IP & Location Info (using ip-api.com - free for non-commercial/limited use)
            # Note: We use a try-except to handle potential network issues or API blocks
            location_str = "Unknown"
            ip_addr = "Unknown"
            try:
                response = requests.get("http://ip-api.com/json/", timeout=5)
                if response.status_code == 200:
                    data = response.json()
                    ip_addr = data.get("query", "Unknown")
                    city = data.get("city", "")
                    country = data.get("country", "")
                    org = data.get("org", "Unknown")
                    location_str = f"{city}, {country} ({org})"
            except Exception as e:
                logging.error(f"Failed to fetch location info: {e}")

            # 3. Format Discord Message (Embed)
            payload = {
                "embeds": [{
                    "title": "🚀 ExcelMatcher Execution Detected",
                    "color": 0x3498db,  # Blue
                    "fields": [
                        {"name": "Computer Name", "value": hostname, "inline": True},
                        {"name": "OS", "value": os_info, "inline": True},
                        {"name": "Public IP", "value": ip_addr, "inline": False},
                        {"name": "Location", "value": location_str, "inline": False},
                    ],
                    "footer": {"text": "Real-time Monitoring System"},
                    "timestamp": None # Discord will auto-timestamp if omitted or provided
                }]
            }

            # 4. Send to Discord
            requests.post(webhook_url, json=payload, timeout=10)
            logging.info(f"Usage status reported to Discord for {hostname}")

        except Exception as e:
            logging.error(f"Error in Discord status report: {e}")

    # Run in background
    thread = threading.Thread(target=_task, daemon=True)
    thread.start()
