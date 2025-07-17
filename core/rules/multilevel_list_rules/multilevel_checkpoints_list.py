from .multilevel_checkpoints import MultilevelCheckPoint
from core.utils.list_utils import (
    is_multilevel,
    is_list_format_matching,
    check_level_indent_exact,
)
from win32com.client import constants

def get_all_multilevel_checkpoints():
    return [
        MultilevelCheckPoint("Is multilevel list", is_multilevel),

        MultilevelCheckPoint("Level 1 uses A/B/C format", 
            lambda ps: is_list_format_matching(ps, r"^[A-Z]$", level=1)
        ),

        MultilevelCheckPoint("Level 1 align at 0cm, indent at 1cm", 
            lambda ps: check_level_indent_exact(ps, level=1, expected_align=0, expected_indent=1)
        ),

        MultilevelCheckPoint("Level 2 bullet is not a circle", 
            lambda ps: is_list_format_matching(ps, r"^(?!\u2022$).+", level=2) # not a circle
        ),

        MultilevelCheckPoint("Level 2 align at 1cm, indent at 2cm", 
            lambda ps: check_level_indent_exact(ps, level=2, expected_align=1, expected_indent=2)
        ),
    ]
