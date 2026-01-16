import gspread
import logging
from pathlib import Path
from datetime import datetime
import schedule
import time

# =====================
# LOG 設定（每天一個檔）
# =====================
BASE_DIR = Path(__file__).resolve().parent
LOG_DIR = BASE_DIR / "Logger"
LOG_DIR.mkdir(exist_ok=True)

def setup_logger():
    today = datetime.now().strftime("%Y-%m-%d")
    log_file = LOG_DIR / f"Log_{today}.txt"

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_file, encoding="utf-8"),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)
logger = setup_logger()
logger.info("程式啟動")

# =====================
# 工作內容
# =====================
def job():
    logger.info("05:00 job 開始")
    try:
        share_link = "https://docs.google.com/spreadsheets/d/1k2yjKUA4m886_Vj24V_jnpt4ClTkQIjoabBBQKb4Jko/edit"
        gc = gspread.service_account(filename='token.json')
        worksheet = gc.open_by_url(share_link).sheet1
        for row, value in enumerate(worksheet.col_values(3)[4:], start=5):
            logger.info(f"Row {row}, Col 3 = {value}")
        logger.info("05:00 job 結束")
    except Exception:
        logger.exception("05:00 job 發生錯誤")

# =====================
# schedule 設定
# =====================
schedule.every().day.at("05:00").do(job)
# =====================
# 主迴圈
# =====================
while True:
    schedule.run_pending()
    time.sleep(30)  # 30 秒檢查一次即可
