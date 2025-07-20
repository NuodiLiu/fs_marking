# core/pipeline.py

import os
import time
import traceback

from core.document_loader import load_word_document
from core.rule_engine import evaluate
from core.utils.utils import validate_zid

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)

SUBMISSION_DIR = os.path.join(ROOT_DIR, "data", "submissions")
LOG_DIR = os.path.join(ROOT_DIR, "logs")

def find_word_file(student_folder):
    time.sleep(0.5)
    word_files = []
    for root, dirs, files in os.walk(student_folder):
        for file in files:
            if file.startswith(".") or file.startswith("~"):
                continue
            if file.lower().endswith((".doc", ".docx")):
                word_files.append(os.path.join(root, file))

    if len(word_files) == 0:
        raise FileNotFoundError("❌ No Word document found.")
    elif len(word_files) > 1:
        raise ValueError("❌ Multiple Word files found.")
    return word_files[0]
    

def run_batch(config, writer=None):
    if not os.path.exists(LOG_DIR):
        os.makedirs(LOG_DIR)

    try:
        for zid_folder in os.listdir(SUBMISSION_DIR):
            # submission/z1234567
            student_path = os.path.join(SUBMISSION_DIR, zid_folder)
            if not os.path.isdir(student_path):
                continue

            # validate folder name is zid
            if not validate_zid(zid_folder):
                print(f"⚠️ Skipped invalid folder name: {zid_folder}")
                continue

            print(f"🔍 Checking {zid_folder}...")

            try:
                word_path = find_word_file(student_path)
                doc, word_app = load_word_document(word_path)

                result = evaluate(doc, config.RULES)

                # word COM close
                doc.Close(False)
                word_app.Quit()
                # python ref delete - based on reference count
                del doc
                del word_app

                if writer is not None:
                    writer.write(zid_folder, result)
                print(f"✅ Finished {zid_folder}: {result['total']} marks\n\n")

            except Exception as e:
                error_msg = f"❌ Failed {zid_folder}: {e}"
                print(error_msg)
                with open(os.path.join(LOG_DIR, f"{zid_folder}_error.log"), "w", encoding="utf-8") as f:
                    f.write(error_msg + "\n")
                    f.write(traceback.format_exc())
    
    finally:
        # Ensure Excel writer is properly closed and saved
        if writer is not None and hasattr(writer, 'close'):
            try:
                writer.close()
            except Exception as e:
                print(f"❌ Error closing writer: {e}")
