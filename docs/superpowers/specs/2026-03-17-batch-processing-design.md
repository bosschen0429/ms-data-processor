# Batch Processing & Open Output Folder — Design Spec

**Date:** 2026-03-17
**Branch:** Multiple_data
**Status:** Approved (v3 — post spec-review round 2 fixes)

---

## 1. Overview

Add two features to the MS Data Deduplication Tool:

1. **Batch file processing** — users can select multiple files at once; each file is processed independently with the same parameters and saved to the preset output folder.
2. **Open Output Folder button** — a button placed alongside "Start Processing" that opens the output folder in the system file explorer.

---

## 2. Scope

### In scope
- Multi-file selection via `askopenfilenames()`
- Sequential batch processing in a background thread
- Per-file error isolation (skip failures, continue processing remaining files)
- Unified parameters (m/z tolerance, RT tolerance, Top N) applied to all files
- "Open Output Folder" button (Windows: `os.startfile`, macOS: `subprocess.run(['open', ...])`)
- Batch summary report in the status area

### Out of scope
- Parallel / concurrent file processing
- Per-file parameter configuration
- Progress bar (status text is sufficient)

---

## 3. Architecture

`MSDataProcessor` is **not modified**. All changes are in `MSProcessorGUI`.

```
MSProcessorGUI
├── select_files()          ← replaces select_file()
├── process_data()          ← validates params, spawns background thread
├── _batch_worker()         ← NEW: runs in background thread
│   ├── for each file:
│   │   ├── MSDataProcessor().process(file, top_n)   ← new instance per file
│   │   ├── processor.save_results(df, output_path)
│   │   └── root.after(0, update_status, msg)        ← thread-safe UI update
│   └── root.after(0, _on_batch_complete, results)
├── _on_batch_complete()    ← NEW: called on main thread when batch finishes
└── open_output_folder()    ← NEW: opens output_dir in file explorer
```

> **Important:** A new `MSDataProcessor` instance **must be created inside the loop** for each file. Reusing one instance across files is incorrect because `MSDataProcessor` stores per-file state on `self` (e.g., `rt_col`, `mz_col`, `intensity_cols`, `source_excel_path`, `temp_mz_rt_cols`).

---

## 4. GUI Changes

### 4.1 State variables

Remove `self.input_file = None` from `__init__`. Replace with:

| Variable | Type | Initial value | Description |
|----------|------|---------------|-------------|
| `self.input_files` | `list[str]` | `[]` | Selected file paths |
| `self.processing` | `bool` | `False` | Guard flag; prevents re-clicking Start Processing during a run |
| `self.process_btn` | `tk.Button` / `ttk.Button` | — | Reference to the Start Processing button, saved in `create_widgets` as `self.process_btn = ...` so it can be disabled/re-enabled from `_on_batch_complete` |

### 4.2 File selection label

- Default: `"No file selected"`
- 1 file: `"sample_A.xlsx"`
- 2+ files: `"sample_A.xlsx and 2 more files"` (first filename + count of remaining)

### 4.3 Button layout

```
[ ▶ Start Processing ]  [ 📂 Open Output Folder ]
```

Both buttons are always visible. "Start Processing" is disabled during batch processing and re-enabled on completion. "Open Output Folder" is always enabled.

**Colors (Windows `tk.Button` only — macOS `ttk.Button` uses system styling):**
- Start Processing: existing `'success'` = `#4CAF50`
- Open Output Folder: new `'secondary'` = `#607D8B`, hover `'secondary_dark'` = `#455A64`

Add to `COLORS` dict:
```python
'secondary': '#607D8B',
'secondary_dark': '#455A64',
```

Save the button reference: `self.process_btn = process_btn` immediately after creating the Start Processing button.

### 4.4 `update_status()` — remove `root.update()`

The existing `update_status()` contains `self.root.update()`. This call **must be removed** as part of this change. When called via `root.after(0, ...)` from the background thread, `root.update()` causes reentrancy (processes pending events recursively), which is a well-known source of hard-to-reproduce tkinter bugs on Windows. The method should only insert text; tkinter's event loop handles screen refreshes.

---

## 5. Batch Processing Logic

### 5.1 Imports

Add to top-level imports (remove any inline import of `subprocess`):
```python
import threading
import subprocess
```

`from datetime import datetime` is already present in the file — no change needed.

### 5.2 `select_files()`

Replaces `select_file()`. Update Browse button command in `create_widgets` from `self.select_file` → `self.select_files` in **both** branches (macOS `ttk.Button` and Windows `tk.Button`).

Also update the section label in `create_widgets` from `"1. Select Input File"` → `"1. Select Input Files"` (singular → plural).

```python
def select_files(self):
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

### 5.3 `process_data()` (entry point)

1. Guard: if `self.processing`, return immediately
2. Validate `self.input_files` is not empty → `messagebox.showerror` if empty
3. Parse and validate parameters (mz_tol, rt_tol, top_n) — show error dialog on invalid value
4. Clear status area
5. Set `self.processing = True`, disable Start Processing button
6. Spawn background thread:
```python
t = threading.Thread(
    target=self._batch_worker,
    args=(list(self.input_files), mz_tol, rt_tol, top_n),
    daemon=True
)
t.start()
```
7. Remove the existing `messagebox.showinfo("Success", ...)` call — the batch summary in the status area is the sole completion notification.

### 5.4 `_batch_worker(files, mz_tol, rt_tol, top_n)`

Runs in background thread.

```python
def _batch_worker(self, files, mz_tol, rt_tol, top_n):
    results = []
    total = len(files)
    for idx, file_path in enumerate(files, 1):
        name = Path(file_path).name
        self.root.after(0, self.update_status, f"\n[{idx}/{total}] 處理中: {name}")
        try:
            processor = MSDataProcessor(mz_tolerance_ppm=mz_tol, rt_tolerance=rt_tol)
            df_result, stats = processor.process(file_path, top_n)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            suffix = Path(file_path).suffix
            stem = Path(file_path).stem
            output_path = self.output_dir / f"processed_{stem}_{timestamp}{suffix}"
            processor.save_results(df_result, str(output_path))
            results.append(('success', file_path, stats, output_path))
            red = stats.get('red_preserved_count', 0)
            msg = (f"  ✔ {name} → "
                   f"{stats['original_count']} → {stats['output_count']} signals"
                   + (f" (含 {red} 紅色保留列)" if red > 0 else ""))
            self.root.after(0, self.update_status, msg)
        except Exception as e:
            results.append(('error', file_path, str(e)))
            self.root.after(0, self.update_status, f"  ✗ {name} → 失敗: {e}")
    self.root.after(0, self._on_batch_complete, results)
```

> Summary line format uses `stats['original_count']` and `stats['output_count']`. Do **not** use `unique_count` — that excludes the effect of Top N.

### 5.5 `_on_batch_complete(results)`

Called on main thread. Re-enables the Start Processing button and displays summary:

```
==================================================
批量處理完成！共 3 個檔案
✔ sample_A.xlsx → 312 → 89 signals (含 2 紅色保留列)
✔ sample_B.xlsx → 204 → 61 signals
✗ sample_C.xlsx → 失敗: Cannot identify required columns
輸出資料夾：C:\path\to\output_Replicates_eliminating_tool
==================================================
```

Summary footer uses `str(self.output_dir)` (full path), not `.name`, so users see the correct location even if the fallback directory was used.

Sets `self.processing = False` and re-enables Start Processing button.

### 5.6 `open_output_folder()`

```python
def open_output_folder(self):
    if not self.output_dir.exists():
        messagebox.showerror("Error", f"Output folder not found:\n{self.output_dir}")
        return
    if sys.platform == 'win32':
        os.startfile(str(self.output_dir))
    elif sys.platform == 'darwin':
        subprocess.run(['open', str(self.output_dir)])
    else:
        subprocess.run(['xdg-open', str(self.output_dir)])
```

---

## 6. Error Handling

| Scenario | Behaviour |
|----------|-----------|
| No files selected | Show error dialog, abort |
| Invalid parameter value | Show error dialog, abort before starting batch |
| Single file processing error | Log to status area, mark as `✗`, continue with next file |
| All files fail | Batch completes with all-failure summary; no crash |
| Output directory not writable | Existing fallback logic in `__init__` handles this |
| Output directory missing at open time | Show error dialog in `open_output_folder()` |
| User clicks Browse mid-batch | Intentional: Browse is **not** disabled during processing. `_batch_worker` receives a snapshot copy of `input_files` at start, so the running job is unaffected. The file label will update on-screen, which is acceptable. |

---

## 7. Thread Safety

All tkinter widget updates from `_batch_worker` must use `self.root.after(0, callback)`. Direct widget calls from the background thread are forbidden and will cause undefined behaviour.

`update_status()` must **not** contain `self.root.update()` — remove it as part of this change.

---

## 8. Files Modified

| File | Change |
|------|--------|
| `ms_processor.py` | Add top-level `import threading`, `import subprocess` (remove inline import); update `COLORS` dict; remove `self.root.update()` from `update_status`; rename `select_file` → `select_files`; update Browse button command; rewrite `process_data`; add `_batch_worker`, `_on_batch_complete`, `open_output_folder`; add Open Folder button in `create_widgets` |

---

## 9. Out-of-scope / Future

- Drag-and-drop file selection
- Cancel button to abort mid-batch
- Per-file Top N override
