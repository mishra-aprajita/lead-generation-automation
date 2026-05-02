"""
scheduler.py — Bonus Feature: Scheduled Automation Trigger
Runs the lead generation pipeline on a defined schedule using the `schedule` library.
Usage: python scheduler.py
"""

import schedule
import time
import logging
from datetime import datetime
from scraper import run_pipeline

logging.basicConfig(
    filename="scheduler.log",
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)


def job():
    print(f"\n⏰ Scheduled run triggered at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    logging.info("Scheduled pipeline run started.")
    try:
        df = run_pipeline()
        logging.info(f"Pipeline completed. {len(df)} leads collected.")
    except Exception as e:
        logging.error(f"Pipeline failed: {e}")
        print(f"❌ Error: {e}")


if __name__ == "__main__":
    print("🕐 Scheduler started. Pipeline will run daily at 08:00 AM.")
    print("   Press Ctrl+C to stop.\n")

    schedule.every().day.at("08:00").do(job)
    # Uncomment below to test immediately:
    # schedule.every(1).minutes.do(job)

    while True:
        schedule.run_pending()
        time.sleep(30)
