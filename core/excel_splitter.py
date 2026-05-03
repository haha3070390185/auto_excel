import os
import pandas as pd
from typing import List, Optional, Callable, Dict


class ExcelSplitter:
    def __init__(self):
        self.progress_callback = None
        self.error_callback = None

    def set_progress_callback(self, callback: Callable[[int, str], None]):
        self.progress_callback = callback

    def set_error_callback(self, callback: Callable[[str], None]):
        self.error_callback = callback

    def _report_progress(self, progress: int, message: str):
        if self.progress_callback:
            self.progress_callback(progress, message)

    def _report_error(self, message: str):
        if self.error_callback:
            self.error_callback(message)

    def split_by_column(
        self,
        file_path: str,
        output_dir: str,
        split_column: str,
        sheet_name: Optional[str] = None,
        file_prefix: str = "",
        include_headers: bool = True
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            self._report_progress(10, "正在读取数据...")
            
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            if split_column not in df.columns:
                self._report_error(f"列 '{split_column}' 不存在于数据中")
                return False
            
            self._report_progress(30, "正在分析数据分组...")
            
            unique_values = df[split_column].dropna().unique()
            total_groups = len(unique_values)
            
            if total_groups == 0:
                self._report_error("没有可拆分的数据")
                return False
            
            self._report_progress(40, f"发现 {total_groups} 个分组，开始拆分...")
            
            for idx, value in enumerate(unique_values):
                self._report_progress(
                    int(40 + (idx / total_groups) * 50),
                    f"正在处理分组: {value}"
                )
                
                filtered_df = df[df[split_column] == value]
                
                safe_value = self._sanitize_filename(str(value))
                if file_prefix:
                    output_filename = f"{file_prefix}_{safe_value}.xlsx"
                else:
                    output_filename = f"{safe_value}.xlsx"
                
                output_path = os.path.join(output_dir, output_filename)
                
                filtered_df.to_excel(
                    output_path,
                    index=False,
                    engine='openpyxl',
                    header=include_headers
                )
            
            self._report_progress(100, f"拆分完成! 共生成 {total_groups} 个文件")
            return True
        
        except Exception as e:
            self._report_error(f"拆分过程中发生错误: {str(e)}")
            return False

    def split_by_row_count(
        self,
        file_path: str,
        output_dir: str,
        rows_per_file: int,
        sheet_name: Optional[str] = None,
        file_prefix: str = "part",
        include_headers: bool = True
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        if rows_per_file <= 0:
            self._report_error("每个文件的行数必须大于0")
            return False
        
        try:
            self._report_progress(10, "正在读取数据...")
            
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
            else:
                df = pd.read_excel(file_path)
            
            total_rows = len(df)
            if total_rows == 0:
                self._report_error("文件中没有数据")
                return False
            
            self._report_progress(30, f"共 {total_rows} 行数据，准备拆分...")
            
            import math
            total_files = math.ceil(total_rows / rows_per_file)
            
            self._report_progress(40, f"将拆分为 {total_files} 个文件...")
            
            for file_idx in range(total_files):
                self._report_progress(
                    int(40 + (file_idx / total_files) * 50),
                    f"正在生成第 {file_idx + 1} 个文件"
                )
                
                start_row = file_idx * rows_per_file
                end_row = min((file_idx + 1) * rows_per_file, total_rows)
                
                if file_idx == 0 or not include_headers:
                    chunk_df = df.iloc[start_row:end_row]
                else:
                    header_row = df.iloc[0:1]
                    data_rows = df.iloc[start_row:end_row]
                    chunk_df = pd.concat([header_row, data_rows], ignore_index=True)
                
                output_filename = f"{file_prefix}_{file_idx + 1}.xlsx"
                output_path = os.path.join(output_dir, output_filename)
                
                chunk_df.to_excel(
                    output_path,
                    index=False,
                    engine='openpyxl',
                    header=include_headers
                )
            
            self._report_progress(100, f"拆分完成! 共生成 {total_files} 个文件")
            return True
        
        except Exception as e:
            self._report_error(f"拆分过程中发生错误: {str(e)}")
            return False

    def split_sheets_to_files(
        self,
        file_path: str,
        output_dir: str,
        file_prefix: str = ""
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            self._report_progress(10, "正在读取工作簿...")
            
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names
            total_sheets = len(sheet_names)
            
            if total_sheets == 0:
                self._report_error("工作簿中没有工作表")
                return False
            
            self._report_progress(20, f"发现 {total_sheets} 个工作表...")
            
            for idx, sheet_name in enumerate(sheet_names):
                self._report_progress(
                    int(20 + (idx / total_sheets) * 70),
                    f"正在处理工作表: {sheet_name}"
                )
                
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                safe_name = self._sanitize_filename(sheet_name)
                if file_prefix:
                    output_filename = f"{file_prefix}_{safe_name}.xlsx"
                else:
                    output_filename = f"{safe_name}.xlsx"
                
                output_path = os.path.join(output_dir, output_filename)
                
                df.to_excel(output_path, index=False, engine='openpyxl')
            
            self._report_progress(100, f"拆分完成! 共生成 {total_sheets} 个文件")
            return True
        
        except Exception as e:
            self._report_error(f"拆分过程中发生错误: {str(e)}")
            return False

    def get_columns(
        self,
        file_path: str,
        sheet_name: Optional[str] = None
    ) -> List[str]:
        try:
            if sheet_name:
                df = pd.read_excel(file_path, sheet_name=sheet_name, nrows=0)
            else:
                df = pd.read_excel(file_path, nrows=0)
            return list(df.columns)
        except Exception as e:
            self._report_error(f"获取列名时出错: {str(e)}")
            return []

    def get_sheet_names(self, file_path: str) -> List[str]:
        try:
            excel_file = pd.ExcelFile(file_path)
            return excel_file.sheet_names
        except Exception as e:
            self._report_error(f"获取工作表名称时出错: {str(e)}")
            return []

    def _sanitize_filename(self, filename: str) -> str:
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename.strip()
