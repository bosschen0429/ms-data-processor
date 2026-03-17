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


# ---------------------------------------------------------------------------
# Task 4: select_files label behaviour
# ---------------------------------------------------------------------------
class TestSelectFiles(unittest.TestCase):
    def _make_gui(self):
        gui = object.__new__(MSProcessorGUI)
        gui.root = MagicMock()
        gui.COLORS = {'text': '#212121', 'text_secondary': '#757575'}
        gui.file_label = MagicMock()
        gui.input_files = []
        return gui

    def test_single_file_shows_filename(self):
        gui = self._make_gui()
        with patch('tkinter.filedialog.askopenfilenames',
                   return_value=('/data/sample_A.xlsx',)):
            gui.select_files()
        gui.file_label.config.assert_called_once_with(
            text='sample_A.xlsx', fg='#212121'
        )
        self.assertEqual(gui.input_files, ['/data/sample_A.xlsx'])

    def test_multiple_files_shows_count(self):
        gui = self._make_gui()
        paths = ('/data/sample_A.xlsx', '/data/sample_B.xlsx', '/data/sample_C.xlsx')
        with patch('tkinter.filedialog.askopenfilenames', return_value=paths):
            gui.select_files()
        call_kwargs = gui.file_label.config.call_args[1]
        self.assertIn('sample_A.xlsx', call_kwargs['text'])
        self.assertIn('2 more files', call_kwargs['text'])
        self.assertEqual(len(gui.input_files), 3)

    def test_cancel_does_not_change_state(self):
        gui = self._make_gui()
        with patch('tkinter.filedialog.askopenfilenames', return_value=()):
            gui.select_files()
        gui.file_label.config.assert_not_called()
        self.assertEqual(gui.input_files, [])


if __name__ == '__main__':
    unittest.main()
