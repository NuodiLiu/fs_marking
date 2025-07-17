# core/document_loader.py

import win32com.client # type: ignore
import os

def load_word_document(path: str):
    if not os.path.exists(path):
        raise FileNotFoundError(f"File does not exist: {path}")

    word_app = win32com.client.Dispatch("Word.Application")
    word_app.Visible = False  # do not pop up the actual word window

    try:
        while word_app.Documents.Count > 0:
            doc = word_app.Documents(1)
            if doc.Saved == 0:  # means not saved
                doc.Save()
            doc.Close(SaveChanges=False)
            
        doc = word_app.Documents.Open(path)
        return doc, word_app
    except Exception as e:
        word_app.Quit()
        raise RuntimeError(f"Cannot open word document: {e}")
