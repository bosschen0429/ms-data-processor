import sys
import types
import unittest
from unittest.mock import MagicMock, patch

# ---------------------------------------------------------------------------
# Stub tkinter so tests run headlessly (no display required)
# ---------------------------------------------------------------------------
tk_stub = types.ModuleType('tkinter')
for name in ['Tk', 'Frame', 'Label', 'Text', 'Button', 'Entry', 'StringVar', 'END']:
    setattr(tk_stub, name, MagicMock)
tk_stub.filedialog = MagicMock()
tk_stub.messagebox = MagicMock()
ttk_stub = types.ModuleType('tkinter.ttk')
ttk_stub.Style = MagicMock()
ttk_stub.Button = MagicMock()
sys.modules['tkinter'] = tk_stub
sys.modules['tkinter.ttk'] = ttk_stub

import importlib
import pathlib

spec = importlib.util.spec_from_file_location(
    "ms_processor",
    pathlib.Path(__file__).parent.parent / "ms_processor.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
MSProcessorGUI = mod.MSProcessorGUI


# ---------------------------------------------------------------------------
# Task 3: update_status must not call root.update()
# ---------------------------------------------------------------------------
class TestUpdateStatus(unittest.TestCase):
    def _make_gui(self):
        gui = object.__new__(MSProcessorGUI)
        gui.root = MagicMock()
        gui.status_text = MagicMock()
        return gui

    def test_update_status_does_not_call_root_update(self):
        """update_status must not call root.update() — causes reentrancy in threaded mode."""
        gui = self._make_gui()
        gui.update_status("hello")
        gui.root.update.assert_not_called()


if __name__ == '__main__':
    unittest.main()
