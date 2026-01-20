import pandas as pd
import numpy as np
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import os
from datetime import datetime

class MSDataProcessor:
    """Mass Spectrometry Data Processor"""
    
    def __init__(self, mz_tolerance_ppm=20, rt_tolerance=1):
        """
        Initialize processor
        
        Parameters:
        -----------
        mz_tolerance_ppm : float
            m/z tolerance (ppm)
        rt_tolerance : float
            RT tolerance
        """
        self.mz_tolerance = mz_tolerance_ppm / 1_000_000
        self.rt_tolerance = rt_tolerance
        
    def load_data(self, file_path):
        """
        Load data and automatically identify columns (supports Excel, CSV, TSV)

        Parameters:
        -----------
        file_path : str
            File path

        Returns:
        --------
        pd.DataFrame
            DataFrame with all columns
        """
        file_path = str(file_path)

        # Read file based on extension
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.tsv') or file_path.endswith('.txt'):
            df = pd.read_csv(file_path, sep='\t')
        elif file_path.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"Unsupported file format. Supported: .xlsx, .xls, .csv, .tsv, .txt")

        # 自動識別欄位（只要包含關鍵詞即可，大小寫不敏感）
        rt_col = self._find_column(df.columns, ['rt', 'retention'])
        mz_col = self._find_column(df.columns, ['m/z', 'mz', 'mass'])
        intensity_col = self._find_column(df.columns, ['area', 'intensity', 'abundance', 'height'])
        id_col = self._find_column(df.columns, ['id'])

        # 判斷資料來源（僅供顯示）
        has_mzmine = any('mzmine' in str(col).lower() for col in df.columns)
        self.data_source = "MZmine" if has_mzmine else "FeatureHunter"

        if rt_col and mz_col and intensity_col:
            self.rt_col = rt_col
            self.mz_col = mz_col
            self.intensity_col = intensity_col

            # 如果有 ID 欄位且資料來源是 MZmine，過濾 NA 資料
            if id_col and has_mzmine:
                df = df[
                    df[id_col].notna() &
                    (df[id_col].astype(str).str.strip().str.upper() != 'NA') &
                    df[rt_col].notna() &
                    df[mz_col].notna() &
                    df[intensity_col].notna()
                ]
        else:
            available_cols = "\nAvailable columns: " + ", ".join(df.columns.tolist())
            raise ValueError(f"Cannot identify required columns.\nPlease check your file headers.{available_cols}")

        self.all_columns = list(df.columns)

        # Remove invalid data (m/z and intensity > 0)
        df = df[(df[self.mz_col] > 0) & (df[self.intensity_col] > 0)]
        df = df.dropna(subset=[self.rt_col, self.mz_col, self.intensity_col])

        return df.reset_index(drop=True)
    
    def _find_column(self, columns, keywords):
        """
        Find matching column name - 只要欄位名包含任一關鍵詞即可（大小寫不敏感）

        Parameters:
        -----------
        columns : list
            All column names
        keywords : list
            Keywords where at least ONE must be present in the column name

        Returns:
        --------
        str or None
            Found column name
        """
        for col in columns:
            col_lower = str(col).lower().strip()
            if any(kw.lower() in col_lower for kw in keywords):
                return col
        return None
    
    def find_unique_signals(self, df):
        """
        Find unique signals (remove duplicates), keep all other columns
        使用分箱策略優化效能，避免 O(n²) 複雜度

        Parameters:
        -----------
        df : pd.DataFrame
            Original data

        Returns:
        --------
        pd.DataFrame
            De-duplicated data
        """
        if len(df) == 0:
            return df

        from collections import defaultdict

        # 提取數值陣列以加速存取
        rt_values = df[self.rt_col].values
        mz_values = df[self.mz_col].values
        intensity_values = df[self.intensity_col].values

        # 計算每個資料點的 RT 箱子索引
        bin_indices = (rt_values / self.rt_tolerance).astype(int)

        # 建立箱子到索引的映射
        bins = defaultdict(list)
        for idx in range(len(df)):
            bins[bin_indices[idx]].append(idx)

        # 追蹤哪些索引要保留
        keep_mask = [True] * len(df)

        # 對每個資料點，只比較同箱和相鄰箱的資料
        for idx in range(len(df)):
            if not keep_mask[idx]:
                continue

            rt = rt_values[idx]
            mz = mz_values[idx]
            intensity = intensity_values[idx]
            bin_idx = bin_indices[idx]

            # 檢查同箱和相鄰箱（確保不漏掉邊界情況）
            for check_bin in [bin_idx - 1, bin_idx, bin_idx + 1]:
                if check_bin not in bins:
                    continue

                for other_idx in bins[check_bin]:
                    if other_idx <= idx or not keep_mask[other_idx]:
                        continue

                    other_rt = rt_values[other_idx]
                    other_mz = mz_values[other_idx]
                    other_intensity = intensity_values[other_idx]

                    # RT tolerance check
                    if abs(other_rt - rt) <= self.rt_tolerance:
                        # m/z tolerance check (use larger m/z as denominator)
                        reference_mz = max(other_mz, mz)
                        if reference_mz > 0:
                            mz_diff_ratio = abs(other_mz - mz) / reference_mz
                            if mz_diff_ratio <= self.mz_tolerance:
                                # Found duplicate, keep higher intensity
                                if other_intensity > intensity:
                                    keep_mask[idx] = False
                                    break
                                else:
                                    keep_mask[other_idx] = False

            if not keep_mask[idx]:
                continue

        # 篩選保留的資料
        kept_indices = [i for i, keep in enumerate(keep_mask) if keep]
        return df.iloc[kept_indices].reset_index(drop=True)
    
    def process(self, file_path, top_n=None):
        """
        Complete processing workflow
        
        Parameters:
        -----------
        file_path : str
            Input file path
        top_n : int, optional
            Output top N signals, None means output all
            
        Returns:
        --------
        tuple
            (processed DataFrame, statistics dictionary)
        """
        # Load data
        df_original = self.load_data(file_path)
        original_count = len(df_original)
        
        # Remove duplicates
        df_unique = self.find_unique_signals(df_original)
        unique_count = len(df_unique)
        
        # Sort by intensity
        df_sorted = df_unique.sort_values(self.intensity_col, ascending=False).reset_index(drop=True)
        
        # Take top N
        if top_n and top_n > 0:
            df_result = df_sorted.head(top_n)
        else:
            df_result = df_sorted
        
        # Statistics
        stats = {
            'original_count': original_count,
            'unique_count': unique_count,
            'output_count': len(df_result),
            'data_source': self.data_source
        }
        
        return df_result, stats
    
    def save_results(self, df, output_path):
        """
        Save results (supports Excel, CSV, TSV)
        
        Parameters:
        -----------
        df : pd.DataFrame
            Data to save
        output_path : str
            Output file path
        """
        output_path = str(output_path)
        
        if output_path.endswith('.csv'):
            df.to_csv(output_path, index=False)
        elif output_path.endswith('.tsv') or output_path.endswith('.txt'):
            df.to_csv(output_path, sep='\t', index=False)
        elif output_path.endswith(('.xlsx', '.xls')):
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Top Results', index=False)
                
                # Format intensity column as scientific notation
                workbook = writer.book
                worksheet = writer.sheets['Top Results']
                
                # Find intensity column position
                intensity_col_idx = list(df.columns).index(self.intensity_col) + 1
                
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=intensity_col_idx)
                    cell.number_format = '0.00E+00'
        else:
            raise ValueError(f"Unsupported output format. Supported: .xlsx, .xls, .csv, .tsv, .txt")


class MSProcessorGUI:
    """Graphical User Interface with flat design"""
    
    # Color scheme - high contrast flat design
    COLORS = {
        'bg': '#F5F5F5',           # Light gray background
        'card': '#FFFFFF',          # White cards
        'primary': '#2196F3',       # Blue primary
        'primary_dark': '#1976D2',  # Darker blue
        'success': '#4CAF50',       # Green
        'success_dark': '#388E3C',  # Darker green
        'text': '#212121',          # Dark text
        'text_secondary': '#757575', # Gray text
        'border': '#E0E0E0',        # Light border
        'shadow': '#00000010'       # Subtle shadow
    }
    
    def __init__(self, root):
        self.root = root
        self.root.title("MS Data Deduplication Tool")
        self.root.geometry("700x550")
        
        # Detect platform and apply appropriate styling
        self.is_macos = sys.platform == 'darwin'
        
        if self.is_macos:
            # macOS specific configuration
            try:
                # Use ttk.Style for better macOS compatibility
                self.style = ttk.Style()
                self.style.theme_use('aqua')
            except:
                pass
        
        self.root.configure(bg=self.COLORS['bg'])
        
        # Make window non-resizable for consistent layout
        self.root.resizable(False, False)
        
        self.processor = None
        self.input_file = None
        self.param_entries = []  # Initialize before create_widgets
        
        # Get the directory where the executable is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            if self.is_macos:
                # For macOS .app bundle, get the directory containing the .app
                # sys.executable points to: YourApp.app/Contents/MacOS/YourApp
                # We want the directory containing YourApp.app
                executable_path = Path(sys.executable)
                # Go up: MacOS -> Contents -> YourApp.app -> parent directory
                app_bundle = executable_path.parent.parent.parent
                self.base_dir = app_bundle.parent
            else:
                # For Windows executable
                self.base_dir = Path(sys.executable).parent
        else:
            # Running as script
            self.base_dir = Path(__file__).parent
        
        # Create output directory with error handling
        try:
            self.output_dir = self.base_dir / "output"
            self.output_dir.mkdir(exist_ok=True)
            # Test write permission
            test_file = self.output_dir / ".write_test"
            test_file.touch()
            test_file.unlink()
        except (PermissionError, OSError) as e:
            # If we can't write to the app directory, use user's Documents folder
            if self.is_macos:
                home = Path.home()
                self.output_dir = home / "Documents" / "MS_Data_Output"
            else:
                # Windows fallback to Documents
                home = Path.home()
                self.output_dir = home / "Documents" / "MS_Data_Output"
            
            try:
                self.output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as mkdir_error:
                # Last resort: use temporary directory
                import tempfile
                self.output_dir = Path(tempfile.gettempdir()) / "MS_Data_Output"
                self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.create_widgets()
    
    def create_card(self, parent, **kwargs):
        """Create a card-style frame"""
        frame = tk.Frame(
            parent,
            bg=self.COLORS['card'],
            highlightbackground=self.COLORS['border'],
            highlightthickness=1,
            **kwargs
        )
        return frame
    
    def _create_button(self, parent, text, command, color_key='primary', **kwargs):
        """Create a styled button with platform-specific handling"""
        if self.is_macos:
            # Use ttk.Button for macOS with custom styling
            button = ttk.Button(
                parent,
                text=text,
                command=command,
                **kwargs
            )
            # Apply hover effect using bind
            def on_enter(e):
                # macOS ttk buttons handle their own hover states
                pass
            def on_leave(e):
                pass
            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
        else:
            # Use tk.Button with full styling for Windows/Linux
            button = tk.Button(
                parent,
                text=text,
                command=command,
                bg=self.COLORS[color_key],
                fg="white",
                font=("Segoe UI", kwargs.get('font_size', 10), "bold"),
                relief="flat",
                cursor="hand2",
                **{k: v for k, v in kwargs.items() if k not in ['font_size']}
            )
            dark_color = color_key + '_dark'
            if dark_color in self.COLORS:
                button.bind("<Enter>", lambda e: button.config(bg=self.COLORS[dark_color]))
                button.bind("<Leave>", lambda e: button.config(bg=self.COLORS[color_key]))
        
        return button
    
    def create_widgets(self):
        # Main container with padding
        main_container = tk.Frame(self.root, bg=self.COLORS['bg'])
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title section
        title_frame = tk.Frame(main_container, bg=self.COLORS['bg'])
        title_frame.pack(fill="x", pady=(0, 20))
        
        title_font = ("SF Pro Display", 24, "bold") if self.is_macos else ("Segoe UI", 24, "bold")
        subtitle_font = ("SF Pro Text", 10) if self.is_macos else ("Segoe UI", 10)
        
        title_label = tk.Label(
            title_frame,
            text="MS Data Deduplication Tool",
            font=title_font,
            bg=self.COLORS['bg'],
            fg=self.COLORS['text']
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            title_frame,
            text="Remove duplicate signals and export top results",
            font=subtitle_font,
            bg=self.COLORS['bg'],
            fg=self.COLORS['text_secondary']
        )
        subtitle_label.pack()
        
        # File selection card
        file_card = self.create_card(main_container)
        file_card.pack(fill="x", pady=(0, 15))
        
        file_inner = tk.Frame(file_card, bg=self.COLORS['card'])
        file_inner.pack(fill="x", padx=20, pady=15)
        
        tk.Label(
            file_inner,
            text="1. Select Input File",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLORS['card'],
            fg=self.COLORS['text']
        ).pack(anchor="w", pady=(0, 10))
        
        file_row = tk.Frame(file_inner, bg=self.COLORS['card'])
        file_row.pack(fill="x")
        
        label_font = ("SF Pro Text", 9) if self.is_macos else ("Segoe UI", 9)
        
        self.file_label = tk.Label(
            file_row,
            text="No file selected",
            font=label_font,
            bg=self.COLORS['card'],
            fg=self.COLORS['text_secondary'],
            anchor="w"
        )
        self.file_label.pack(side="left", fill="x", expand=True)
        
        # Create browse button
        if self.is_macos:
            select_btn = ttk.Button(
                file_row,
                text="Browse Files",
                command=self.select_file
            )
        else:
            select_btn = tk.Button(
                file_row,
                text="Browse Files",
                command=self.select_file,
                bg=self.COLORS['primary'],
                fg="white",
                font=("Segoe UI", 10, "bold"),
                relief="flat",
                cursor="hand2",
                padx=20,
                pady=8
            )
            select_btn.bind("<Enter>", lambda e: select_btn.config(bg=self.COLORS['primary_dark']))
            select_btn.bind("<Leave>", lambda e: select_btn.config(bg=self.COLORS['primary']))
        
        select_btn.pack(side="right")
        
        # Parameters card
        param_card = self.create_card(main_container)
        param_card.pack(fill="x", pady=(0, 15))
        
        param_inner = tk.Frame(param_card, bg=self.COLORS['card'])
        param_inner.pack(fill="x", padx=20, pady=15)
        
        tk.Label(
            param_inner,
            text="2. Configure Parameters",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLORS['card'],
            fg=self.COLORS['text']
        ).pack(anchor="w", pady=(0, 15))
        
        # Parameter grid
        param_grid = tk.Frame(param_inner, bg=self.COLORS['card'])
        param_grid.pack(fill="x")
        
        # m/z tolerance
        self.mz_tolerance_var = self._create_param_row(param_grid, "m/z Tolerance (ppm):", "20", 
                               "Acceptable mass difference")
        
        # RT tolerance
        self.rt_tolerance_var = self._create_param_row(param_grid, "RT Tolerance:", "1", 
                               "Acceptable retention time difference")
        
        # Top N
        self.top_n_var = self._create_param_row(param_grid, "Output Top N Signals:", "10", 
                               "Enter 0 for all signals")
        
        # Process button
        if self.is_macos:
            process_btn = ttk.Button(
                main_container,
                text="Start Processing",
                command=self.process_data
            )
        else:
            process_btn = tk.Button(
                main_container,
                text="Start Processing",
                command=self.process_data,
                bg=self.COLORS['success'],
                fg="white",
                font=("Segoe UI", 12, "bold"),
                relief="flat",
                cursor="hand2",
                padx=30,
                pady=15
            )
            process_btn.bind("<Enter>", lambda e: process_btn.config(bg=self.COLORS['success_dark']))
            process_btn.bind("<Leave>", lambda e: process_btn.config(bg=self.COLORS['success']))
        
        process_btn.pack(pady=(0, 15))
        
        # Status card
        status_card = self.create_card(main_container)
        status_card.pack(fill="both", expand=True)
        
        status_inner = tk.Frame(status_card, bg=self.COLORS['card'])
        status_inner.pack(fill="both", expand=True, padx=20, pady=15)
        
        tk.Label(
            status_inner,
            text="Processing Status",
            font=("Segoe UI", 11, "bold"),
            bg=self.COLORS['card'],
            fg=self.COLORS['text']
        ).pack(anchor="w", pady=(0, 10))
        
        text_font = ("Menlo", 9) if self.is_macos else ("Consolas", 9)
        
        self.status_text = tk.Text(
            status_inner,
            height=8,
            font=text_font,
            bg="#FAFAFA",
            fg=self.COLORS['text'],
            relief="flat",
            borderwidth=0,
            state="disabled"
        )
        self.status_text.pack(fill="both", expand=True)
    
    def _create_param_row(self, parent, label_text, default_value, hint_text):
        """Create a parameter input row"""
        row_frame = tk.Frame(parent, bg=self.COLORS['card'])
        row_frame.pack(fill="x", pady=8)
        
        param_font = ("SF Pro Text", 10, "bold") if self.is_macos else ("Segoe UI", 10, "bold")
        hint_font = ("SF Pro Text", 9) if self.is_macos else ("Segoe UI", 9)
        entry_font = ("SF Pro Text", 10) if self.is_macos else ("Segoe UI", 10)
        
        # Left side: Label
        tk.Label(
            row_frame,
            text=label_text,
            font=param_font,
            bg=self.COLORS['card'],
            fg=self.COLORS['text'],
            anchor="w",
            width=20
        ).pack(side="left")
        
        # Middle: Entry
        entry_var = tk.StringVar(value=default_value)
        entry = tk.Entry(
            row_frame,
            textvariable=entry_var,
            font=entry_font,
            bg="white",
            fg=self.COLORS['text'],
            relief="solid",
            borderwidth=1,
            width=12
        )
        entry.pack(side="left", padx=(0, 15))
        
        # Right side: Hint text
        tk.Label(
            row_frame,
            text=hint_text,
            font=hint_font,
            bg=self.COLORS['card'],
            fg=self.COLORS['text_secondary'],
            anchor="w"
        ).pack(side="left", fill="x", expand=True)
        
        return entry_var  # Return the StringVar directly
    
    def select_file(self):
        """Select input file"""
        file_path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[
                ("All Supported Formats", "*.xlsx *.xls *.csv *.tsv *.txt"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("TSV files", "*.tsv *.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.input_file = file_path
            self.file_label.config(
                text=Path(file_path).name,
                fg=self.COLORS['text']
            )
    
    def update_status(self, message):
        """Update status display"""
        self.status_text.config(state="normal")
        self.status_text.insert("end", message + "\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()
    
    def process_data(self):
        """Process data"""
        if not self.input_file:
            messagebox.showerror("Error", "Please select an input file first!")
            return
        
        try:
            # Clear status
            self.status_text.config(state="normal")
            self.status_text.delete(1.0, "end")
            self.status_text.config(state="disabled")
            
            # Read parameters
            mz_tol = float(self.mz_tolerance_var.get())
            rt_tol = float(self.rt_tolerance_var.get())
            top_n = int(self.top_n_var.get())
            if top_n == 0:
                top_n = None
            
            self.update_status("Starting processing...")
            self.update_status(f"Output directory: {self.output_dir}")
            
            # Create processor
            processor = MSDataProcessor(mz_tolerance_ppm=mz_tol, rt_tolerance=rt_tol)
            
            # Process data
            self.update_status("Loading data...")
            df_result, stats = processor.process(self.input_file, top_n)
            
            # Display identified columns
            self.update_status(f"\nData Source: {stats['data_source']}")
            self.update_status(f"Identified Columns:")
            self.update_status(f"  RT: {processor.rt_col}")
            self.update_status(f"  m/z: {processor.mz_col}")
            self.update_status(f"  Intensity: {processor.intensity_col}")
            self.update_status(f"Other columns preserved: {len(processor.all_columns) - 3}")
            
            # Generate output filename with timestamp
            input_path = Path(self.input_file)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"processed_{input_path.stem}_{timestamp}{input_path.suffix}"
            output_path = self.output_dir / output_filename
            
            # Ensure output directory exists and is writable
            try:
                self.output_dir.mkdir(parents=True, exist_ok=True)
            except Exception as mkdir_error:
                self.update_status(f"\nWarning: Could not create output directory: {mkdir_error}")
                # Fallback to Desktop
                desktop = Path.home() / "Desktop" / "MS_Data_Output"
                desktop.mkdir(parents=True, exist_ok=True)
                output_path = desktop / output_filename
                self.update_status(f"Using alternative location: {desktop}")
            
            # Save results
            self.update_status("\nSaving results...")
            processor.save_results(df_result, str(output_path))
            
            # Display statistics
            self.update_status("\n" + "="*50)
            self.update_status("Processing Complete!")
            self.update_status(f"Original data: {stats['original_count']} signals")
            self.update_status(f"After deduplication: {stats['unique_count']} signals")
            self.update_status(f"Output count: {stats['output_count']} signals")
            self.update_status(f"\nResults saved to:\n{output_path}")
            
            # Show file in Finder/Explorer
            if self.is_macos:
                import subprocess
                subprocess.run(["open", "-R", str(output_path)])
            
            messagebox.showinfo("Success", f"Processing complete!\n\nResults saved to:\n{output_path}")
            
        except Exception as e:
            import traceback
            error_details = traceback.format_exc()
            messagebox.showerror("Error", f"An error occurred during processing:\n{str(e)}")
            self.update_status(f"\nError: {str(e)}")
            self.update_status(f"\nDetails:\n{error_details}")


def main():
    """Main program"""
    root = tk.Tk()
    app = MSProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
