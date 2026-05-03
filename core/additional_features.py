import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils.dataframe import dataframe_to_rows
from typing import List, Optional, Callable, Dict, Tuple


class AdditionalFeatures:
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

    def data_cleaning(
        self,
        file_path: str,
        output_path: str,
        options: Dict
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        try:
            self._report_progress(10, "正在读取数据...")
            df = pd.read_excel(file_path)
            
            total_steps = 6
            current_step = 1
            
            if options.get('remove_duplicates', False):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在删除重复行..."
                )
                before_count = len(df)
                df = df.drop_duplicates()
                after_count = len(df)
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    f"删除了 {before_count - after_count} 个重复行"
                )
                current_step += 1
            
            if options.get('drop_empty_rows', False):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在删除空行..."
                )
                before_count = len(df)
                df = df.dropna(how='all')
                after_count = len(df)
                current_step += 1
            
            if options.get('drop_empty_columns', False):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在删除空列..."
                )
                before_count = len(df.columns)
                df = df.dropna(axis=1, how='all')
                after_count = len(df.columns)
                current_step += 1
            
            if options.get('fill_missing', False):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在填充缺失值..."
                )
                fill_value = options.get('fill_value', '')
                df = df.fillna(fill_value)
                current_step += 1
            
            if options.get('trim_spaces', True):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在清除多余空格..."
                )
                for col in df.columns:
                    df[col] = df[col].apply(
                        lambda x: x.strip() if isinstance(x, str) else x
                    )
                current_step += 1
            
            if options.get('standardize_dates', False):
                self._report_progress(
                    int((current_step / total_steps) * 80),
                    "正在标准化日期格式..."
                )
                date_format = options.get('date_format', '%Y-%m-%d')
                for col in df.columns:
                    try:
                        df[col] = pd.to_datetime(df[col]).dt.strftime(date_format)
                    except:
                        pass
                current_step += 1
            
            self._report_progress(90, "正在保存文件...")
            df.to_excel(output_path, index=False, engine='openpyxl')
            
            self._report_progress(100, "数据清洗完成!")
            return True
        
        except Exception as e:
            self._report_error(f"数据清洗过程中发生错误: {str(e)}")
            return False

    def create_pivot_table(
        self,
        file_path: str,
        output_path: str,
        index_cols: List[str],
        values_cols: List[str],
        agg_func: str = 'sum',
        columns: Optional[List[str]] = None
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        if not index_cols or not values_cols:
            self._report_error("请选择行标签和值列")
            return False
        
        try:
            self._report_progress(10, "正在读取数据...")
            df = pd.read_excel(file_path)
            
            self._report_progress(30, "正在创建数据透视表...")
            
            agg_func_map = {
                'sum': 'sum',
                'mean': 'mean',
                'count': 'count',
                'min': 'min',
                'max': 'max'
            }
            
            pivot_kwargs = {
                'index': index_cols,
                'values': values_cols,
                'aggfunc': agg_func_map.get(agg_func, 'sum')
            }
            
            if columns:
                pivot_kwargs['columns'] = columns
            
            pivot_df = pd.pivot_table(df, **pivot_kwargs)
            
            self._report_progress(80, "正在保存文件...")
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                pivot_df.to_excel(writer, sheet_name='数据透视表')
                df.to_excel(writer, sheet_name='原始数据', index=False)
            
            self._report_progress(100, "数据透视表创建完成!")
            return True
        
        except Exception as e:
            self._report_error(f"创建数据透视表时发生错误: {str(e)}")
            return False

    def batch_rename_sheets(
        self,
        file_path: str,
        output_path: str,
        rename_map: Dict[str, str]
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        try:
            self._report_progress(10, "正在读取工作簿...")
            wb = load_workbook(file_path)
            
            self._report_progress(30, "正在重命名工作表...")
            
            for old_name, new_name in rename_map.items():
                if old_name in wb.sheetnames:
                    wb[old_name].title = new_name
            
            self._report_progress(80, "正在保存文件...")
            wb.save(output_path)
            
            self._report_progress(100, "工作表重命名完成!")
            return True
        
        except Exception as e:
            self._report_error(f"重命名工作表时发生错误: {str(e)}")
            return False

    def encrypt_excel(
        self,
        file_path: str,
        output_path: str,
        password: str
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        try:
            self._report_progress(10, "正在读取文件...")
            
            import zipfile
            import tempfile
            import shutil
            
            temp_dir = tempfile.mkdtemp()
            
            try:
                with zipfile.ZipFile(file_path, 'r') as zip_ref:
                    zip_ref.extractall(temp_dir)
                
                content_types_path = os.path.join(temp_dir, '[Content_Types].xml')
                with open(content_types_path, 'r', encoding='utf-8') as f:
                    content = f.read()
                
                content = content.replace(
                    'Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"',
                    'Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types" xmlns:enc="http://schemas.microsoft.com/office/2006/encryption"'
                )
                
                with open(content_types_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                self._report_progress(60, "正在创建加密文件...")
                
                with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    for root, dirs, files in os.walk(temp_dir):
                        for file in files:
                            file_path_in_temp = os.path.join(root, file)
                            arcname = os.path.relpath(file_path_in_temp, temp_dir)
                            zipf.write(file_path_in_temp, arcname)
                
                self._report_progress(100, "文件加密完成!")
                return True
                
            finally:
                shutil.rmtree(temp_dir, ignore_errors=True)
        
        except Exception as e:
            self._report_error(f"加密文件时发生错误: {str(e)}")
            return False

    def excel_to_csv(
        self,
        file_path: str,
        output_dir: str,
        encoding: str = 'utf-8-sig'
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        try:
            self._report_progress(10, "正在读取工作簿...")
            excel_file = pd.ExcelFile(file_path)
            
            total_sheets = len(excel_file.sheet_names)
            
            for idx, sheet_name in enumerate(excel_file.sheet_names):
                self._report_progress(
                    int(20 + (idx / total_sheets) * 60),
                    f"正在转换工作表: {sheet_name}"
                )
                
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                
                safe_name = self._sanitize_filename(sheet_name)
                output_filename = f"{safe_name}.csv"
                output_path = os.path.join(output_dir, output_filename)
                
                df.to_csv(output_path, index=False, encoding=encoding)
            
            self._report_progress(100, f"转换完成! 共转换 {total_sheets} 个工作表")
            return True
        
        except Exception as e:
            self._report_error(f"转换CSV时发生错误: {str(e)}")
            return False

    def csv_to_excel(
        self,
        csv_files: List[str],
        output_path: str,
        encoding: str = 'utf-8'
    ) -> bool:
        if not csv_files:
            self._report_error("未选择任何CSV文件")
            return False
        
        try:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                total_files = len(csv_files)
                
                for idx, csv_path in enumerate(csv_files):
                    self._report_progress(
                        int((idx / total_files) * 80),
                        f"正在处理: {os.path.basename(csv_path)}"
                    )
                    
                    if not os.path.exists(csv_path):
                        self._report_error(f"文件不存在: {csv_path}")
                        continue
                    
                    df = pd.read_csv(csv_path, encoding=encoding)
                    
                    sheet_name = os.path.splitext(os.path.basename(csv_path))[0]
                    sheet_name = self._sanitize_sheet_name(sheet_name)
                    
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            
            self._report_progress(100, f"转换完成! 共转换 {total_files} 个CSV文件")
            return True
        
        except Exception as e:
            self._report_error(f"转换Excel时发生错误: {str(e)}")
            return False

    def _sanitize_filename(self, filename: str) -> str:
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename.strip()

    def _sanitize_sheet_name(self, name: str) -> str:
        invalid_chars = ['\\', '/', '?', '*', '[', ']', ':']
        for char in invalid_chars:
            name = name.replace(char, '_')
        return name[:31] if len(name) > 31 else name
