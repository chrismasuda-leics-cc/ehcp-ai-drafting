from app.settings import TEMP_DIR, OUTPUT_DIR
import os

os.makedirs(TEMP_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)
