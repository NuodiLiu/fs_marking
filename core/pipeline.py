# core/pipeline.py

import os
import traceback
from core.document_loader import load_word_document
from core.rule_engine import evaluate
from core.utils.utils import validate_zid
import time

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(BASE_DIR)

SUBMISSION_DIR = os.path.join(ROOT_DIR, "data", "submissions")
LOG_DIR = os.path.join(ROOT_DIR, "logs")

def find_word_file(student_folder):
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

    total_start_time = time.time()
    timings = []

    for zid_folder in os.listdir(SUBMISSION_DIR):
        student_path = os.path.join(SUBMISSION_DIR, zid_folder)
        if not os.path.isdir(student_path):
            continue

        if not validate_zid(zid_folder):
            print(f"⚠️ Skipped invalid folder name: {zid_folder}")
            continue

        print(f"🔍 Checking {zid_folder}...")
        start_time = time.time()

        try:
            word_path = find_word_file(student_path)
            doc, word_app = load_word_document(word_path)

            result = evaluate(doc, config.RULES)

            doc.Close(False)
            word_app.Quit()
            del doc
            del word_app

            if writer is not None:
                writer.write(zid_folder, result)

            elapsed = time.time() - start_time
            timings.append(elapsed)
            # print(f"✅ Finished {zid_folder}: {result['total']} marks ({elapsed:.2f}s)\n\n")

        except Exception as e:
            elapsed = time.time() - start_time
            timings.append(elapsed)
            error_msg = f"❌ Failed {zid_folder} ({elapsed:.2f}s): {e}"
            print(error_msg)
            with open(os.path.join(LOG_DIR, f"{zid_folder}_error.log"), "w", encoding="utf-8") as f:
                f.write(error_msg + "\n")
                f.write(traceback.format_exc())

    total_time = time.time() - total_start_time
    count = len(timings)
    avg_time = sum(timings) / count if count > 0 else 0
    print(f"\n⏱️ Finished batch in {total_time:.2f}s")
    print(f"📊 Average time per document: {avg_time:.2f}s")