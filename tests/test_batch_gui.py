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
sys.modules['ms_processor'] = mod   # required so patch('ms_processor.X') resolves correctly
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


# ---------------------------------------------------------------------------
# Task 6: open_output_folder
# ---------------------------------------------------------------------------
class TestOpenOutputFolder(unittest.TestCase):
    def _make_gui(self, exists=True):
        gui = object.__new__(MSProcessorGUI)
        gui.root = MagicMock()
        p = MagicMock()
        p.exists.return_value = exists
        p.__str__ = MagicMock(return_value='/output/folder')
        gui.output_dir = p
        return gui

    def test_opens_folder_on_windows(self):
        gui = self._make_gui(exists=True)
        with patch('sys.platform', 'win32'), \
             patch('os.startfile') as mock_start:
            gui.open_output_folder()
        mock_start.assert_called_once_with('/output/folder')

    def test_shows_error_if_folder_missing(self):
        gui = self._make_gui(exists=False)
        with patch('tkinter.messagebox.showerror') as mock_err, \
             patch('os.startfile') as mock_start:
            gui.open_output_folder()
        mock_err.assert_called_once()
        mock_start.assert_not_called()


# ---------------------------------------------------------------------------
# Task 7: _batch_worker
# ---------------------------------------------------------------------------
from pathlib import Path as _Path


class TestBatchWorker(unittest.TestCase):
    def _make_gui(self):
        gui = object.__new__(MSProcessorGUI)
        gui.root = MagicMock()
        # Use a real Path-like so / operator works correctly
        gui.output_dir = _Path('/out')
        gui.update_status = MagicMock()
        gui.process_btn = None
        gui.processing = False
        return gui

    def test_success_appends_to_results(self):
        gui = self._make_gui()
        fake_stats = {
            'original_count': 100, 'unique_count': 50, 'output_count': 30,
            'red_preserved_count': 0, 'data_source': 'MZmine', 'sample_count': 3
        }
        fake_df = MagicMock()
        results_captured = []

        gui.root.after = lambda delay, fn, *args: fn(*args)

        with patch('ms_processor.MSDataProcessor') as MockProc:
            instance = MockProc.return_value
            instance.process.return_value = (fake_df, fake_stats)
            instance.save_results.return_value = None
            gui._batch_worker(['/data/file_A.xlsx'], 20, 1.0, 10)

        # _on_batch_complete was called synchronously via root.after mock
        # We verify update_status was called (summary rendered)
        self.assertTrue(gui.update_status.called)

    def test_error_is_captured_not_raised(self):
        """Processing error must be captured in results, not propagate as exception."""
        gui = self._make_gui()
        results_captured = []

        def fake_after(delay, fn, *args):
            if getattr(fn, '__name__', None) == '_on_batch_complete':
                results_captured.extend(args[0])
            else:
                fn(*args)

        gui.root.after = fake_after

        with patch('ms_processor.MSDataProcessor') as MockProc:
            MockProc.return_value.process.side_effect = ValueError("bad columns")
            gui._batch_worker(['/data/bad.xlsx'], 20, 1.0, 10)

        self.assertEqual(len(results_captured), 1)
        status, path, err_msg = results_captured[0]
        self.assertEqual(status, 'error')
        self.assertIn('bad columns', err_msg)

    def test_new_processor_created_per_file(self):
        """Each file must get a fresh MSDataProcessor instance."""
        gui = self._make_gui()
        fake_stats = {
            'original_count': 10, 'unique_count': 5, 'output_count': 5,
            'red_preserved_count': 0, 'data_source': 'X', 'sample_count': 1
        }
        gui.root.after = lambda delay, fn, *args: fn(*args)

        with patch('ms_processor.MSDataProcessor') as MockProc:
            MockProc.return_value.process.return_value = (MagicMock(), fake_stats)
            MockProc.return_value.save_results.return_value = None
            gui._batch_worker(
                ['/data/a.xlsx', '/data/b.xlsx', '/data/c.xlsx'],
                20, 1.0, None
            )
        self.assertEqual(MockProc.call_count, 3)


if __name__ == '__main__':
    unittest.main()
