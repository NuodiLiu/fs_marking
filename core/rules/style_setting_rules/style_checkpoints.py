# core/rules/style_checkpoints.py

class StyleCheckPoint:
    def __init__(self, style_name: str, description: str, check_fn):
        self.style_name = style_name
        self.description = description
        self.check_fn = check_fn  # function(font, para_format) -> bool

    def run(self, style_obj):
        try:
            font = style_obj.Font
            para_format = style_obj.ParagraphFormat
            return self.check_fn(font, para_format)
        except Exception:
            return False
