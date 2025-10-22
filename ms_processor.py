import pandas as pd
import numpy as np
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

class MSDataProcessor:
    """質譜數據處理類別"""
    
    def __init__(self, mz_tolerance_ppm=20, rt_tolerance=1):
        """
        初始化處理器
        
        Parameters:
        -----------
        mz_tolerance_ppm : float
            m/z 容差值 (ppm)
        rt_tolerance : float
            RT 容差值
        """
        self.mz_tolerance = mz_tolerance_ppm / 1_000_000
        self.rt_tolerance = rt_tolerance
        
    def load_data(self, file_path):
        """
        載入數據並自動識別欄位 (支援 Excel, CSV, TSV)
        
        Parameters:
        -----------
        file_path : str
            檔案路徑
            
        Returns:
        --------
        pd.DataFrame
            包含所有欄位的數據框
        """
        file_path = str(file_path)
        
        # 根據副檔名選擇讀取方式
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        elif file_path.endswith('.tsv') or file_path.endswith('.txt'):
            df = pd.read_csv(file_path, sep='\t')
        elif file_path.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file_path)
        else:
            raise ValueError(f"不支援的檔案格式。支援: .xlsx, .xls, .csv, .tsv, .txt")
        
        # 自動識別 RT, m/z, Intensity 欄位
        rt_col = self._find_column(df.columns, [
            'rt', 'retention time', 'retention_time', 'retentiontime',
            'rt (min)', 'rt(min)', 'retention time (min)'
        ])
        mz_col = self._find_column(df.columns, [
            'm/z', 'mz', 'm_z', 'mass', 
            'precursor ion m/z', 'precursor m/z', 'precursormz'
        ])
        intensity_col = self._find_column(df.columns, [
            'intensity', 'int', 'abundance', 'height',
            'precursor ion intensity', 'precursor intensity', 'precursorintensity'
        ])
        
        if not rt_col or not mz_col or not intensity_col:
            missing = []
            if not rt_col: missing.append("RT")
            if not mz_col: missing.append("m/z")
            if not intensity_col: missing.append("Intensity")
            
            available_cols = "\n可用的欄位: " + ", ".join(df.columns.tolist())
            raise ValueError(f"無法識別欄位: {', '.join(missing)}\n請確認標頭包含這些欄位名稱{available_cols}")
        
        # 標記主要欄位
        self.rt_col = rt_col
        self.mz_col = mz_col
        self.intensity_col = intensity_col
        self.all_columns = list(df.columns)
        
        # 移除無效數據 (只檢查 m/z 和 intensity > 0, RT 允許為 0)
        df = df[(df[mz_col] > 0) & (df[intensity_col] > 0)]
        df = df.dropna(subset=[rt_col, mz_col, intensity_col])
        
        return df.reset_index(drop=True)
    
    def _find_column(self, columns, possible_names):
        """
        尋找符合的欄位名稱
        
        Parameters:
        -----------
        columns : list
            所有欄位名稱
        possible_names : list
            可能的欄位名稱列表
            
        Returns:
        --------
        str or None
            找到的欄位名稱
        """
        for col in columns:
            col_lower = str(col).lower().strip()
            for name in possible_names:
                if name in col_lower:
                    return col
        return None
    
    def find_unique_signals(self, df):
        """
        找出唯一訊號 (去除重複),保留所有其他欄位
        
        Parameters:
        -----------
        df : pd.DataFrame
            原始數據
            
        Returns:
        --------
        pd.DataFrame
            去重複後的數據
        """
        unique_signals = []
        unique_indices = []
        
        for idx, row in df.iterrows():
            rt = row[self.rt_col]
            mz = row[self.mz_col]
            intensity = row[self.intensity_col]
            
            # 檢查是否與現有訊號重複
            is_unique = True
            for i, existing_idx in enumerate(unique_indices):
                existing_row = df.loc[existing_idx]
                existing_rt = existing_row[self.rt_col]
                existing_mz = existing_row[self.mz_col]
                existing_intensity = existing_row[self.intensity_col]
                
                # RT 容差檢查
                if abs(existing_rt - rt) <= self.rt_tolerance:
                    # m/z 容差檢查 (使用較大的 m/z 作為分母)
                    reference_mz = max(existing_mz, mz)
                    if reference_mz > 0:
                        mz_diff_ratio = abs(existing_mz - mz) / reference_mz
                        if mz_diff_ratio <= self.mz_tolerance:
                            # 找到重複訊號,保留強度較大的
                            if intensity > existing_intensity:
                                unique_indices[i] = idx
                            is_unique = False
                            break
            
            if is_unique:
                unique_indices.append(idx)
        
        return df.loc[unique_indices].reset_index(drop=True)
    
    def process(self, file_path, top_n=None):
        """
        完整處理流程
        
        Parameters:
        -----------
        file_path : str
            輸入檔案路徑
        top_n : int, optional
            輸出前 N 個訊號,None 表示全部輸出
            
        Returns:
        --------
        tuple
            (處理後的數據框, 統計資訊字典)
        """
        # 載入數據
        df_original = self.load_data(file_path)
        original_count = len(df_original)
        
        # 去除重複
        df_unique = self.find_unique_signals(df_original)
        unique_count = len(df_unique)
        
        # 按強度排序
        df_sorted = df_unique.sort_values(self.intensity_col, ascending=False).reset_index(drop=True)
        
        # 取前 N 個
        if top_n and top_n > 0:
            df_result = df_sorted.head(top_n)
        else:
            df_result = df_sorted
        
        # 統計資訊
        stats = {
            'original_count': original_count,
            'unique_count': unique_count,
            'output_count': len(df_result)
        }
        
        return df_result, stats
    
    def save_results(self, df, output_path):
        """
        儲存結果 (支援 Excel, CSV, TSV)
        
        Parameters:
        -----------
        df : pd.DataFrame
            要儲存的數據
        output_path : str
            輸出檔案路徑
        """
        output_path = str(output_path)
        
        if output_path.endswith('.csv'):
            df.to_csv(output_path, index=False)
        elif output_path.endswith('.tsv') or output_path.endswith('.txt'):
            df.to_csv(output_path, sep='\t', index=False)
        elif output_path.endswith(('.xlsx', '.xls')):
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Top Results', index=False)
                
                # 格式化 Intensity 欄位為科學記號
                workbook = writer.book
                worksheet = writer.sheets['Top Results']
                
                # 找到 Intensity 欄位的位置
                intensity_col_idx = list(df.columns).index(self.intensity_col) + 1
                
                for row in range(2, len(df) + 2):
                    cell = worksheet.cell(row=row, column=intensity_col_idx)
                    cell.number_format = '0.00E+00'
        else:
            raise ValueError(f"不支援的輸出格式。支援: .xlsx, .xls, .csv, .tsv, .txt")


class MSProcessorGUI:
    """圖形化介面"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("質譜數據去重複處理工具")
        self.root.geometry("600x450")
        
        self.processor = None
        self.input_file = None
        
        self.create_widgets()
    
    def create_widgets(self):
        # 標題
        title_label = tk.Label(
            self.root, 
            text="質譜數據去重複處理工具",
            font=("Arial", 16, "bold")
        )
        title_label.pack(pady=10)
        
        # 檔案選擇框架
        file_frame = tk.LabelFrame(self.root, text="檔案選擇", padx=10, pady=10)
        file_frame.pack(padx=20, pady=10, fill="x")
        
        self.file_label = tk.Label(file_frame, text="未選擇檔案", fg="gray")
        self.file_label.pack(side="left", padx=5)
        
        tk.Button(
            file_frame, 
            text="選擇檔案 (Excel/CSV/TSV)", 
            command=self.select_file
        ).pack(side="right", padx=5)
        
        # 參數設定框架
        param_frame = tk.LabelFrame(self.root, text="參數設定", padx=10, pady=10)
        param_frame.pack(padx=20, pady=10, fill="x")
        
        # m/z 容差
        tk.Label(param_frame, text="m/z 容差 (ppm):").grid(row=0, column=0, sticky="w", pady=5)
        self.mz_tolerance_var = tk.StringVar(value="20")
        tk.Entry(param_frame, textvariable=self.mz_tolerance_var, width=15).grid(row=0, column=1, pady=5)
        
        # RT 容差
        tk.Label(param_frame, text="RT 容差:").grid(row=1, column=0, sticky="w", pady=5)
        self.rt_tolerance_var = tk.StringVar(value="1")
        tk.Entry(param_frame, textvariable=self.rt_tolerance_var, width=15).grid(row=1, column=1, pady=5)
        
        # 輸出數量
        tk.Label(param_frame, text="輸出前 N 個訊號:").grid(row=2, column=0, sticky="w", pady=5)
        self.top_n_var = tk.StringVar(value="10")
        tk.Entry(param_frame, textvariable=self.top_n_var, width=15).grid(row=2, column=1, pady=5)
        tk.Label(param_frame, text="(輸入 0 表示全部)", fg="gray").grid(row=2, column=2, sticky="w", padx=5)
        
        # 執行按鈕
        tk.Button(
            self.root, 
            text="開始處理", 
            command=self.process_data,
            bg="#4CAF50",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=10
        ).pack(pady=20)
        
        # 狀態顯示
        self.status_text = tk.Text(self.root, height=8, width=70, state="disabled")
        self.status_text.pack(padx=20, pady=10)
    
    def select_file(self):
        """選擇輸入檔案"""
        file_path = filedialog.askopenfilename(
            title="選擇資料檔案",
            filetypes=[
                ("所有支援格式", "*.xlsx *.xls *.csv *.tsv *.txt"),
                ("Excel files", "*.xlsx *.xls"),
                ("CSV files", "*.csv"),
                ("TSV files", "*.tsv *.txt"),
                ("All files", "*.*")
            ]
        )
        
        if file_path:
            self.input_file = file_path
            self.file_label.config(text=Path(file_path).name, fg="black")
    
    def update_status(self, message):
        """更新狀態顯示"""
        self.status_text.config(state="normal")
        self.status_text.insert("end", message + "\n")
        self.status_text.see("end")
        self.status_text.config(state="disabled")
        self.root.update()
    
    def process_data(self):
        """處理數據"""
        if not self.input_file:
            messagebox.showerror("錯誤", "請先選擇輸入檔案!")
            return
        
        try:
            # 清空狀態
            self.status_text.config(state="normal")
            self.status_text.delete(1.0, "end")
            self.status_text.config(state="disabled")
            
            # 讀取參數
            mz_tol = float(self.mz_tolerance_var.get())
            rt_tol = float(self.rt_tolerance_var.get())
            top_n = int(self.top_n_var.get())
            if top_n == 0:
                top_n = None
            
            self.update_status("開始處理...")
            
            # 建立處理器
            processor = MSDataProcessor(mz_tolerance_ppm=mz_tol, rt_tolerance=rt_tol)
            
            # 處理數據
            self.update_status("讀取數據中...")
            df_result, stats = processor.process(self.input_file, top_n)
            
            # 顯示識別的欄位
            self.update_status(f"已識別欄位:")
            self.update_status(f"  RT: {processor.rt_col}")
            self.update_status(f"  m/z: {processor.mz_col}")
            self.update_status(f"  Intensity: {processor.intensity_col}")
            self.update_status(f"保留的其他欄位: {len(processor.all_columns) - 3} 個")
            
            # 生成輸出檔名 (保持相同格式)
            input_path = Path(self.input_file)
            output_path = input_path.parent / f"{input_path.stem}_processed{input_path.suffix}"
            
            # 儲存結果
            self.update_status("儲存結果中...")
            processor.save_results(df_result, str(output_path))
            
            # 顯示統計
            self.update_status("\n處理完成!")
            self.update_status(f"原始數據: {stats['original_count']} 筆")
            self.update_status(f"去重複後: {stats['unique_count']} 筆")
            self.update_status(f"輸出數量: {stats['output_count']} 筆")
            self.update_status(f"\n結果已儲存至:\n{output_path}")
            
            messagebox.showinfo("完成", f"處理完成!\n\n結果已儲存至:\n{output_path}")
            
        except Exception as e:
            messagebox.showerror("錯誤", f"處理時發生錯誤:\n{str(e)}")
            self.update_status(f"\n錯誤: {str(e)}")


def main():
    """主程式"""
    root = tk.Tk()
    app = MSProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()