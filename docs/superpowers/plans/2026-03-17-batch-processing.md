# Batch Processing & Open Output Folder — Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Allow users to select and process multiple MS data files in one batch run, and add an "Open Output Folder" button to the GUI.

**Architecture:** All changes are confined to `MSProcessorGUI` in `ms_processor.py`. `MSDataProcessor` (core logic) is untouched. Batch processing runs in a background `threading.Thread`; all tkinter widget updates are marshalled back to the main thread via `root.after(0, ...)`.

**Tech Stack:** Python 3, tkinter, threading (stdlib), pandas, openpyxl

---

## File Map

| File | Action | Responsibility |
|------|--------|----------------|
| `ms_processor.py` | Modify | All GUI changes: imports, COLORS, state, widgets, batch logic |
| `tests/test_batch_gui.py` | Create | Unit tests for batch logic (mocking tkinter & MSDataProcessor) |

---

### Task 1: Add imports and secondary colour

**Files:**
- Modify: `ms_processor.py` (lines 1–12, `COLORS` dict ~line 508)

- [ ] **Step 1: Add top-level imports**

In `ms_processor.py`, replace the existing import block at the top:
```python
# BEFORE (line 8 area):
import sys
import os
from copy import copy
from datetime import datetime

# AFTER — add threading and subprocess; remove any inline import subprocess later:
import sys
import os
import threading
import subprocess
from copy import copy
from datetime import datetime
```

- [ ] **Step 2: Add secondary colour to COLORS dict**

In `MSProcessorGUI.COLORS` (~line 508), add two entries:
```python
COLORS = {
    # ... existing entries ...
    'secondary': '#607D8B',       # Blue-grey for Open Folder button
    'secondary_dark': '#455A64',  # Darker blue-grey for hover
}
```

- [ ] **Step 3: Remove inline subprocess import**

Search for `import subprocess` inside `process_data()` (~line 958) and delete that line (it is now at the top level).

- [ ] **Step 4: Commit**
```bash
git add ms_processor.py
git commit -m "feat: add threading/subprocess imports and secondary colour"
```

---

### Task 2: Update `__init__` state variables

**Files:**
- Modify: `ms_processor.py` `MSProcessorGUI.__init__` (~line 543)

- [ ] **Step 1: Replace `input_file` with `input_files` and add `processing`**

Find:
```python
self.processor = None
self.input_file = None
self.param_entries = []  # Initialize before create_widgets
```

Replace with:
```python
self.processor = None
self.input_files = []       # List of selected file paths (replaces input_file)
self.processing = False     # Guard: prevents re-clicking Start Processing mid-batch
self.process_btn = None     # Set in create_widgets; kept here for reference
self.param_entries = []     # Initialize before create_widgets
```

- [ ] **Step 2: Commit**
```bash
git add ms_processor.py
git commit -m "feat: update GUI state variables for batch mode"
```

---

### Task 3: Fix `update_status()` — remove `root.update()`

**Files:**
- Modify: `ms_processor.py` `update_status()` (~line 878)

- [ ] **Step 1: Write the failing test**

Create `tests/test_batch_gui.py`:
```python
import sys, types, unittest
from unittest.mock import MagicMock, patch

# Stub tkinter so tests run headlessly
tk_stub = types.ModuleType('tkinter')
for name in ['Tk','Frame','Label','Text','Button','Entry','StringVar','END']:
    setattr(tk_stub, name, MagicMock)
tk_stub.filedialog = MagicMock()
tk_stub.messagebox = MagicMock()
ttk_stub = types.ModuleType('tkinter.ttk')
ttk_stub.Style = MagicMock()
ttk_stub.Button = MagicMock()
sys.modules['tkinter'] = tk_stub
sys.modules['tkinter.ttk'] = ttk_stub

import importlib, pathlib, types as _types
spec = importlib.util.spec_from_file_location(
    "ms_processor",
    pathlib.Path(__file__).parent.parent / "ms_processor.py"
)
mod = importlib.util.module_from_spec(spec)
spec.loader.exec_module(mod)
MSProcessorGUI = mod.MSProcessorGUI


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
```

- [ ] **Step 2: Run test — expect FAIL**
```bash
cd "C:/Users/user/Desktop/MS Data process package/ms-data-processor"
python -m pytest tests/test_batch_gui.py::TestUpdateStatus::test_update_status_does_not_call_root_update -v
```
Expected: FAIL (`root.update()` is still there)

- [ ] **Step 3: Remove `self.root.update()` from `update_status()`**

Find in `update_status()`:
```python
    def update_status(self, message):
        """Update status display"""
        self.status_text.config(state="normal")
        self.status_text.insert("end", message + "\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()   # ← DELETE THIS LINE
```

Remove `self.root.update()`.

- [ ] **Step 4: Run test — expect PASS**
```bash
python -m pytest tests/test_batch_gui.py::TestUpdateStatus -v
```
Expected: PASS

- [ ] **Step 5: Commit**
```bash
git add ms_processor.py tests/test_batch_gui.py
git commit -m "fix: remove root.update() from update_status to prevent reentrancy"
```

---

### Task 4: Replace `select_file` with `select_files`

**Files:**
- Modify: `ms_processor.py` `select_file()` (~line 858)

- [ ] **Step 1: Write failing tests**

Add to `tests/test_batch_gui.py`:
```python
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
```

- [ ] **Step 2: Run tests — expect FAIL**
```bash
python -m pytest tests/test_batch_gui.py::TestSelectFiles -v
```
Expected: FAIL (`select_files` not defined)

- [ ] **Step 3: Replace `select_file()` with `select_files()`**

Replace the entire `select_file` method:
```python
def select_files(self):
    """Select one or more input files"""
    file_paths = filedialog.askopenfilenames(
        title="Select Data Files",
        filetypes=[
            ("All Supported Formats", "*.xlsx *.xls *.csv *.tsv *.txt"),
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("TSV files", "*.tsv *.txt"),
            ("All files", "*.*")
        ]
    )
    if file_paths:
        self.input_files = list(file_paths)
        names = [Path(p).name for p in self.input_files]
        if len(names) == 1:
            label_text = names[0]
        else:
            label_text = f"{names[0]} and {len(names) - 1} more files"
        self.file_label.config(text=label_text, fg=self.COLORS['text'])
```

- [ ] **Step 4: Run tests — expect PASS**
```bash
python -m pytest tests/test_batch_gui.py::TestSelectFiles -v
```
Expected: all 3 PASS

- [ ] **Step 5: Commit**
```bash
git add ms_processor.py tests/test_batch_gui.py
git commit -m "feat: replace select_file with multi-file select_files"
```

---

### Task 5: Update `create_widgets()`

**Files:**
- Modify: `ms_processor.py` `create_widgets()` (~line 642)

- [ ] **Step 1: Update section label (singular → plural)**

Find:
```python
        tk.Label(
            file_inner,
            text="1. Select Input File",
```
Replace `"1. Select Input File"` with `"1. Select Input Files"`.

- [ ] **Step 2: Update Browse button commands — both branches**

In the macOS branch (~line 704):
```python
# BEFORE:
            select_btn = ttk.Button(
                file_row,
                text="Browse Files",
                command=self.select_file      # ← change to select_files
            )
# AFTER:
            select_btn = ttk.Button(
                file_row,
                text="Browse Files",
                command=self.select_files
            )
```

In the Windows branch (~line 710):
```python
# BEFORE:
            select_btn = tk.Button(
                file_row,
                text="Browse Files",
                command=self.select_file,     # ← change to select_files
```
```python
# AFTER:
            select_btn = tk.Button(
                file_row,
                text="Browse Files",
                command=self.select_files,
```

- [ ] **Step 3: Save `self.process_btn` and add Open Folder button**

Find the process button creation block (~line 759). After the button is created (after the `else` block closes), save the reference and add the second button:

```python
        # EXISTING: process_btn creation (macOS + Windows branches) stays as-is
        # ADD: save reference and pack both buttons in a row

        self.process_btn = process_btn   # ← add this line before pack()

        # Replace the single process_btn.pack() with a button row:
        btn_row = tk.Frame(main_container, bg=self.COLORS['bg'])
        btn_row.pack(pady=(0, 15))

        process_btn.pack(side="left", padx=(0, 10), in_=btn_row)

        folder_btn = self._create_button(
            btn_row,
            text="📂 Open Output Folder",
            command=self.open_output_folder,
            color_key='secondary'
        )
        folder_btn.pack(side="left")
```

> Remove the original `process_btn.pack(pady=(0, 15))` line — it is replaced by the `btn_row` above.

- [ ] **Step 4: Smoke-test manually**

Run the app and confirm:
- Title now says "1. Select Input Files"
- Clicking Browse opens multi-select dialog
- Both buttons appear side-by-side
- Open Output Folder button is blue-grey

```bash
python ms_processor.py
```

- [ ] **Step 5: Commit**
```bash
git add ms_processor.py
git commit -m "feat: update create_widgets for multi-file and Open Folder button"
```

---

### Task 6: Add `open_output_folder()`

**Files:**
- Modify: `ms_processor.py`

- [ ] **Step 1: Write failing test**

Add to `tests/test_batch_gui.py`:
```python
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
        with patch('tkinter.messagebox.showerror') as mock_err:
            gui.open_output_folder()
        mock_err.assert_called_once()
        # os.startfile must NOT be called
```

- [ ] **Step 2: Run tests — expect FAIL**
```bash
python -m pytest tests/test_batch_gui.py::TestOpenOutputFolder -v
```
Expected: FAIL (`open_output_folder` not defined)

- [ ] **Step 3: Implement `open_output_folder()`**

Add as a new method in `MSProcessorGUI`:
```python
def open_output_folder(self):
    """Open the output folder in the system file explorer"""
    if not self.output_dir.exists():
        messagebox.showerror(
            "Error",
            f"Output folder not found:\n{self.output_dir}"
        )
        return
    if sys.platform == 'win32':
        os.startfile(str(self.output_dir))
    elif sys.platform == 'darwin':
        subprocess.run(['open', str(self.output_dir)])
    else:
        subprocess.run(['xdg-open', str(self.output_dir)])
```

- [ ] **Step 4: Run tests — expect PASS**
```bash
python -m pytest tests/test_batch_gui.py::TestOpenOutputFolder -v
```
Expected: PASS

- [ ] **Step 5: Commit**
```bash
git add ms_processor.py tests/test_batch_gui.py
git commit -m "feat: add open_output_folder with missing-dir guard"
```

---

### Task 7: Implement `_batch_worker` and `_on_batch_complete`

**Files:**
- Modify: `ms_processor.py`

- [ ] **Step 1: Write failing tests**

Add to `tests/test_batch_gui.py`:
```python
from pathlib import Path as _Path

class TestBatchWorker(unittest.TestCase):
    def _make_gui(self):
        gui = object.__new__(MSProcessorGUI)
        gui.root = MagicMock()
        gui.output_dir = MagicMock()
        gui.output_dir.__truediv__ = lambda self, other: _Path('/out') / other
        gui.update_status = MagicMock()
        return gui

    def test_success_appends_to_results(self):
        gui = self._make_gui()
        fake_stats = {
            'original_count': 100, 'unique_count': 50, 'output_count': 30,
            'red_preserved_count': 0, 'data_source': 'MZmine', 'sample_count': 3
        }
        fake_df = MagicMock()
        results_captured = []

        def capture_complete(results):
            results_captured.extend(results)

        gui.root.after = lambda delay, fn, *args: fn(*args)

        with patch('ms_processor.MSDataProcessor') as MockProc:
            instance = MockProc.return_value
            instance.process.return_value = (fake_df, fake_stats)
            instance.save_results.return_value = None
            gui._batch_worker(['/data/file_A.xlsx'], 20, 1.0, 10)

        self.assertEqual(len(results_captured), 1)
        status, path, stats, out = results_captured[0]
        self.assertEqual(status, 'success')
        self.assertIn('file_A', str(path))

    def test_error_is_captured_not_raised(self):
        """Processing error must be captured in results, not propagate as exception."""
        gui = self._make_gui()
        results_captured = []

        def fake_after(delay, fn, *args):
            # Intercept _on_batch_complete to capture results
            if fn.__name__ == '_on_batch_complete':
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
```

- [ ] **Step 2: Run tests — expect FAIL**
```bash
python -m pytest tests/test_batch_gui.py::TestBatchWorker -v
```
Expected: FAIL (`_batch_worker` not defined)

- [ ] **Step 3: Implement `_batch_worker()`**

Add to `MSProcessorGUI`:
```python
def _batch_worker(self, files, mz_tol, rt_tol, top_n):
    """Run in background thread. Processes each file sequentially."""
    results = []
    total = len(files)
    for idx, file_path in enumerate(files, 1):
        name = Path(file_path).name
        self.root.after(
            0, self.update_status,
            f"\n[{idx}/{total}] 處理中: {name}"
        )
        try:
            processor = MSDataProcessor(
                mz_tolerance_ppm=mz_tol,
                rt_tolerance=rt_tol
            )
            df_result, stats = processor.process(file_path, top_n)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = (
                self.output_dir
                / f"processed_{Path(file_path).stem}_{timestamp}{Path(file_path).suffix}"
            )
            processor.save_results(df_result, str(output_path))
            results.append(('success', file_path, stats, output_path))
            red = stats.get('red_preserved_count', 0)
            msg = (
                f"  ✔ {name} → "
                f"{stats['original_count']} → {stats['output_count']} signals"
                + (f" (含 {red} 紅色保留列)" if red > 0 else "")
            )
            self.root.after(0, self.update_status, msg)
        except Exception as e:
            results.append(('error', file_path, str(e)))
            self.root.after(0, self.update_status, f"  ✗ {name} → 失敗: {e}")

    self.root.after(0, self._on_batch_complete, results)
```

- [ ] **Step 4: Implement `_on_batch_complete()`**

Add to `MSProcessorGUI`:
```python
def _on_batch_complete(self, results):
    """Called on main thread when batch finishes. Updates UI and re-enables button."""
    self.processing = False
    if self.process_btn:
        self.process_btn.config(state="normal")

    total = len(results)
    success = [r for r in results if r[0] == 'success']
    errors  = [r for r in results if r[0] == 'error']

    self.update_status("\n" + "=" * 50)
    self.update_status(f"批量處理完成！共 {total} 個檔案")
    for r in success:
        _, path, stats, _ = r
        red = stats.get('red_preserved_count', 0)
        line = (
            f"✔ {Path(path).name} → "
            f"{stats['original_count']} → {stats['output_count']} signals"
            + (f" (含 {red} 紅色保留列)" if red > 0 else "")
        )
        self.update_status(line)
    for r in errors:
        _, path, err = r
        self.update_status(f"✗ {Path(path).name} → 失敗: {err}")
    self.update_status(f"輸出資料夾：{self.output_dir}")
    self.update_status("=" * 50)
```

- [ ] **Step 5: Run tests — expect PASS**
```bash
python -m pytest tests/test_batch_gui.py::TestBatchWorker -v
```
Expected: all 3 PASS

- [ ] **Step 6: Commit**
```bash
git add ms_processor.py tests/test_batch_gui.py
git commit -m "feat: add _batch_worker and _on_batch_complete for threaded batch processing"
```

---

### Task 8: Rewrite `process_data()` entry point

**Files:**
- Modify: `ms_processor.py` `process_data()` (~line 886)

- [ ] **Step 1: Replace `process_data()` body**

Replace the entire method body (keep the method signature):
```python
def process_data(self):
    """Validate inputs and start batch processing in a background thread."""
    if self.processing:
        return

    if not self.input_files:
        messagebox.showerror("Error", "Please select one or more input files first!")
        return

    try:
        mz_tol  = float(self.mz_tolerance_var.get())
        rt_tol  = float(self.rt_tolerance_var.get())
        top_n_v = int(self.top_n_var.get())
        top_n   = top_n_v if top_n_v > 0 else None
    except ValueError:
        messagebox.showerror("Error", "Invalid parameter value. Please enter numbers only.")
        return

    # Clear status area
    self.status_text.config(state="normal")
    self.status_text.delete(1.0, "end")
    self.status_text.config(state="disabled")

    self.update_status(f"輸出資料夾：{self.output_dir}")
    self.update_status(f"共 {len(self.input_files)} 個檔案，開始批量處理...\n")

    self.processing = True
    if self.process_btn:
        self.process_btn.config(state="disabled")

    t = threading.Thread(
        target=self._batch_worker,
        args=(list(self.input_files), mz_tol, rt_tol, top_n),
        daemon=True
    )
    t.start()
```

- [ ] **Step 2: Run the full test suite**
```bash
python -m pytest tests/ -v
```
Expected: all tests PASS

- [ ] **Step 3: Smoke-test end-to-end manually**

Run the app, select 2–3 files, click Start Processing. Verify:
- Status area shows per-file progress
- Start Processing is disabled during run, re-enabled after
- Final summary shows ✔/✗ per file with counts
- Open Output Folder opens the correct folder

```bash
python ms_processor.py
```

- [ ] **Step 4: Commit**
```bash
git add ms_processor.py
git commit -m "feat: rewrite process_data as threaded batch entry point"
```

---

### Task 9: Final cleanup & full test run

**Files:**
- Modify: `ms_processor.py`

- [ ] **Step 1: Verify no remaining references to `self.input_file` (singular)**
```bash
# Unix/Git Bash:
grep -n "self\.input_file[^s]" ms_processor.py
# Windows CMD fallback:
# python -c "import re,pathlib; txt=pathlib.Path('ms_processor.py').read_text(); [print(i+1,l) for i,l in enumerate(txt.splitlines()) if re.search(r'self\.input_file[^s]',l)]"
```
Expected: no matches.

- [ ] **Step 2: Verify no remaining inline `import subprocess`**
```bash
# Unix/Git Bash:
grep -n "import subprocess" ms_processor.py
# Windows CMD fallback:
# python -c "import pathlib; [print(i+1,l) for i,l in enumerate(pathlib.Path('ms_processor.py').read_text().splitlines()) if 'import subprocess' in l]"
```
Expected: exactly 1 match at top of file.

- [ ] **Step 3: Run complete test suite**
```bash
python -m pytest tests/ -v
```
Expected: all PASS.

- [ ] **Step 4: Final commit**
```bash
git add ms_processor.py tests/
git commit -m "feat: batch processing and Open Output Folder complete"
```

---

## Test Commands Reference

| Command | Purpose |
|---------|---------|
| `python -m pytest tests/ -v` | Full suite |
| `python -m pytest tests/test_batch_gui.py -v` | GUI tests only |
| `python ms_processor.py` | Manual smoke test |
