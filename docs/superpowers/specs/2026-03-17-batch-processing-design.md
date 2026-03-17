# Batch Processing & Open Output Folder — Design Spec

**Date:** 2026-03-17
**Branch:** Multiple_data
**Status:** Approved

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
│   │   ├── MSDataProcessor().process(file, top_n)
│   │   ├── processor.save_results(df, output_path)
│   │   └── root.after(0, update_status, msg)  ← thread-safe UI update
│   └── root.after(0, _on_batch_complete, results)
├── _on_batch_complete()    ← NEW: called on main thread when batch finishes
└── open_output_folder()    ← NEW: opens output_dir in file explorer
```

---

## 4. GUI Changes

### 4.1 State variables

| Variable | Type | Description |
|----------|------|-------------|
| `self.input_files` | `list[str]` | Replaces `self.input_file`. Stores selected file paths. |
| `self.processing` | `bool` | Guard flag; prevents re-clicking Start Processing during a run. |

### 4.2 File selection label

- Default: `"No file selected"`
- After selection: `"3 個檔案已選取: sample_A.xlsx, sample_B.xlsx, sample_C.xlsx"`
- Long file lists are truncated with `...` to fit the label width

### 4.3 Button layout

```
[ ▶ Start Processing ]  [ 📂 Open Output Folder ]
```

Both buttons are always visible. "Start Processing" is disabled during batch processing and re-enabled on completion. "Open Output Folder" is always enabled.

**Colors:**
- Start Processing: `#4CAF50` (existing green)
- Open Output Folder: `#607D8B` (blue-grey)

---

## 5. Batch Processing Logic

### 5.1 `select_files()`

```python
def select_files(self):
    file_paths = filedialog.askopenfilenames(
        title="Select Data Files",
        filetypes=[...]  # same as existing
    )
    if file_paths:
        self.input_files = list(file_paths)
        # Update label: "N 個檔案已選取: name1, name2, ..."
```

### 5.2 `process_data()` (entry point)

1. Guard: if `self.processing`, return immediately
2. Validate `self.input_files` is not empty
3. Parse and validate parameters (mz_tol, rt_tol, top_n)
4. Clear status area
5. Set `self.processing = True`, disable Start Processing button
6. Spawn `threading.Thread(target=self._batch_worker, args=(files, mz_tol, rt_tol, top_n), daemon=True).start()`

### 5.3 `_batch_worker(files, mz_tol, rt_tol, top_n)`

Runs in background thread. For each file:

```
results = []
for file_path in files:
    try:
        processor = MSDataProcessor(mz_tolerance_ppm=mz_tol, rt_tolerance=rt_tol)
        df_result, stats = processor.process(file_path, top_n)
        output_path = self.output_dir / f"processed_{stem}_{timestamp}{suffix}"
        processor.save_results(df_result, str(output_path))
        results.append(('success', file_path, stats, output_path))
    except Exception as e:
        results.append(('error', file_path, str(e)))

    # thread-safe status update
    self.root.after(0, self.update_status, status_message)

self.root.after(0, self._on_batch_complete, results)
```

### 5.4 `_on_batch_complete(results)`

Called on the main thread. Displays batch summary and re-enables the Start Processing button:

```
==================================================
批量處理完成！共 3 個檔案
✔ sample_A.xlsx → 312 → 89 signals
✔ sample_B.xlsx → 204 → 61 signals
✗ sample_C.xlsx → 失敗: Cannot identify required columns
輸出資料夾：output_Replicates_eliminating_tool
==================================================
```

Sets `self.processing = False` and re-enables the Start Processing button.

### 5.5 `open_output_folder()`

```python
def open_output_folder(self):
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

---

## 7. Thread Safety

All tkinter widget updates from `_batch_worker` must use `self.root.after(0, callback)`. Direct widget calls from the background thread are forbidden and will cause undefined behaviour.

---

## 8. Files Modified

| File | Change |
|------|--------|
| `ms_processor.py` | GUI class only: `select_files`, `process_data`, `_batch_worker`, `_on_batch_complete`, `open_output_folder`, button layout in `create_widgets` |

---

## 9. Out-of-scope / Future

- Drag-and-drop file selection
- Cancel button to abort mid-batch
- Per-file Top N override
