# core/result_writer.py

class ResultWriter:
    def __init__(self, writers):
        self.writers = writers  # List of writer plugins

    def write(self, zid: str, result: dict):
        for writer in self.writers:
            writer.write(zid, result)
