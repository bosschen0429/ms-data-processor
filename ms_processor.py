import pandas as pd
import numpy as np
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import sys
import os

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
        
        # Identify MZmine columns (priority)
        mzmine_rt_col = self._find_column(df.columns, ['mzmine rt', 'mzmine rt (min)'])
        mzmine_mz_col = self._find_column(df.columns, ['mzmine m/z', 'mzmine mz'])
        mzmine_area_col = self._find_column(df.columns, ['peak area', 'area'])
        mzmine_id_col = self._find_column(df.columns, ['mzmine id'])
        
        # Identify FeatureHunter columns (fallback)
        fh_rt_col = self._find_column(df.columns, [
            'rt', 'retention time', 'retention_time', 'retentiontime',
            'rt (min)', 'rt(min)', 'retention time (min)'
        ])
        fh_mz_col = self._find_column(df.columns, [
            'precursor ion m/z', 'precursor m/z', 'precursormz',
            'm/z', 'mz', 'm_z', 'mass'
        ])
        fh_intensity_col = self._find_column(df.columns, [
            'precursor ion intensity', 'precursor intensity', 'precursorintensity',
            'intensity', 'int', 'abundance', 'height'
        ])
        
        # Filter out rows where MZmine data is NA
        if mzmine_id_col and mzmine_rt_col and mzmine_mz_col and mzmine_area_col:
            # Remove rows where any MZmine column is NA
            df = df[
                df[mzmine_id_col].notna() & 
                (df[mzmine_id_col].astype(str).str.strip().str.upper() != 'NA') &
                df[mzmine_rt_col].notna() & 
                df[mzmine_mz_col].notna() & 
                df[mzmine_area_col].notna()
            ]
            
            # Use MZmine columns
            self.rt_col = mzmine_rt_col
            self.mz_col = mzmine_mz_col
            self.intensity_col = mzmine_area_col
            self.data_source = "MZmine"
        elif fh_rt_col and fh_mz_col and fh_intensity_col:
            # Fallback to FeatureHunter columns
            self.rt_col = fh_rt_col
            self.mz_col = fh_mz_col
            self.intensity_col = fh_intensity_col
            self.data_source = "FeatureHunter"
        else:
            available_cols = "\nAvailable columns: " + ", ".join(df.columns.tolist())
            raise ValueError(f"Cannot identify required columns.\nPlease check your file headers.{available_cols}")
        
        self.all_columns = list(df.columns)
        
        # Remove invalid data (m/z and intensity > 0)
        df = df[(df[self.mz_col] > 0) & (df[self.intensity_col] > 0)]
        df = df.dropna(subset=[self.rt_col, self.mz_col, self.intensity_col])
        
        return df.reset_index(drop=True)
    
    def _find_column(self, columns, possible_names):
        """
        Find matching column name
        
        Parameters:
        -----------
        columns : list
            All column names
        possible_names : list
            List of possible column names
            
        Returns:
        --------
        str or None
            Found column name
        """
        for col in columns:
            col_lower = str(col).lower().strip()
            for name in possible_names:
                if name in col_lower:
                    return col
        return None
    
    def find_unique_signals(self, df):
        """
        Find unique signals (remove duplicates), keep all other columns
        
        Parameters:
        -----------
        df : pd.DataFrame
            Original data
            
        Returns:
        --------
        pd.DataFrame
            De-duplicated data
        """
        unique_signals = []
        unique_indices = []
        
        for idx, row in df.iterrows():
            rt = row[self.rt_col]
            mz = row[self.mz_col]
            intensity = row[self.intensity_col]
            
            # Check for duplicates
            is_unique = True
            for i, existing_idx in enumerate(unique_indices):
                existing_row = df.loc[existing_idx]
                existing_rt = existing_row[self.rt_col]
                existing_mz = existing_row[self.mz_col]
                existing_intensity = existing_row[self.intensity_col]
                
                # RT tolerance check
                if abs(existing_rt - rt) <= self.rt_tolerance:
                    # m/z tolerance check (use larger m/z as denominator)
                    reference_mz = max(existing_mz, mz)
                    if reference_mz > 0:
                        mz_diff_ratio = abs(existing_mz - mz) / reference_mz
                        if mz_diff_ratio <= self.mz_tolerance:
                            # Found duplicate, keep higher intensity
                            if intensity > existing_intensity:
                                unique_indices[i] = idx
                            is_unique = False
                            break
            
            if is_unique:
                unique_indices.append(idx)
        
        return df.loc[unique_indices].reset_index(drop=True)
    
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
        self.root.configure(bg=self.COLORS['bg'])
        
        # Make window non-resizable for consistent layout
        self.root.resizable(False, False)
        
        self.processor = None
        self.input_file = None
        self.param_entries = []  # Initialize before create_widgets
        
        # Get the directory where the executable is located
        if getattr(sys, 'frozen', False):
            # Running as compiled executable
            self.base_dir = Path(sys.executable).parent
        else:
            # Running as script
            self.base_dir = Path(__file__).parent
        
        # Create output directory
        self.output_dir = self.base_dir / "output"
        self.output_dir.mkdir(exist_ok=True)
        
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
    
    def create_widgets(self):
        # Main container with padding
        main_container = tk.Frame(self.root, bg=self.COLORS['bg'])
        main_container.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Title section
        title_frame = tk.Frame(main_container, bg=self.COLORS['bg'])
        title_frame.pack(fill="x", pady=(0, 20))
        
        title_label = tk.Label(
            title_frame,
            text="MS Data Deduplication Tool",
            font=("Segoe UI", 24, "bold"),
            bg=self.COLORS['bg'],
            fg=self.COLORS['text']
        )
        title_label.pack()
        
        subtitle_label = tk.Label(
            title_frame,
            text="Remove duplicate signals and export top results",
            font=("Segoe UI", 10),
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
        
        self.file_label = tk.Label(
            file_row,
            text="No file selected",
            font=("Segoe UI", 9),
            bg=self.COLORS['card'],
            fg=self.COLORS['text_secondary'],
            anchor="w"
        )
        self.file_label.pack(side="left", fill="x", expand=True)
        
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
        select_btn.pack(side="right")
        select_btn.bind("<Enter>", lambda e: select_btn.config(bg=self.COLORS['primary_dark']))
        select_btn.bind("<Leave>", lambda e: select_btn.config(bg=self.COLORS['primary']))
        
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
        process_btn.pack(pady=(0, 15))
        process_btn.bind("<Enter>", lambda e: process_btn.config(bg=self.COLORS['success_dark']))
        process_btn.bind("<Leave>", lambda e: process_btn.config(bg=self.COLORS['success']))
        
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
        
        self.status_text = tk.Text(
            status_inner,
            height=8,
            font=("Consolas", 9),
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
        
        # Left side: Label
        tk.Label(
            row_frame,
            text=label_text,
            font=("Segoe UI", 10, "bold"),
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
            font=("Segoe UI", 10),
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
            font=("Segoe UI", 9),
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
            
            # Generate output filename in output directory
            input_path = Path(self.input_file)
            output_path = self.output_dir / f"{input_path.stem}_processed{input_path.suffix}"
            
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
            
            messagebox.showinfo("Success", f"Processing complete!\n\nResults saved to:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during processing:\n{str(e)}")
            self.update_status(f"\nError: {str(e)}")


def main():
    """Main program"""
    root = tk.Tk()
    app = MSProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
