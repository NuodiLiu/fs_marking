# core/result_writer.py

class ResultWriter:
    def __init__(self, writers):
        self.writers = writers  # List of writer plugins

    def write(self, zid: str, result: dict):
        for writer in self.writers:
            try:
                writer.write(zid, result)
            except Exception as e:
                print(f"❌ Error writing to {type(writer).__name__}: {e}")

    def save(self):
        """Save all writers that have a save method"""
        for writer in self.writers:
            if hasattr(writer, 'save'):
                try:
                    writer.save()
                except Exception as e:
                    print(f"❌ Error saving {type(writer).__name__}: {e}")

    def close(self):
        """Close all writers that have a close method"""
        for writer in self.writers:
            if hasattr(writer, 'close'):
                try:
                    writer.close()
                except Exception as e:
                    print(f"❌ Error closing {type(writer).__name__}: {e}")
