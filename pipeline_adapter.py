"""Compatibility adapter for the optional clone-formatting pipeline.

The public GUI imports ``reformat_docx_clone`` unconditionally, but this
repository snapshot does not include the newer clone pipeline entry point.
Keep the GUI usable by falling back to the standard document reformatter.
"""

from __future__ import annotations

from format_conversion import reformat_docx


def reformat_docx_clone(input_path, output_path, gui_config=None):
    """Fallback clone-formatting entry point used by ``gui_main``."""
    return reformat_docx(input_path, output_path, config=gui_config)
