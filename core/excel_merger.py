import os
import pandas as pd
from typing import List, Optional, Callable


class ExcelMerger:
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

    def merge_files(
        self,
        file_paths: List[str],
        output_path: str,
        merge_type: str = "all_sheets",
        sheet_name: Optional[str] = None,
        add_source_column: bool = False,
        skip_rows: int = 0,
        header_row: int = 0
    ) -> bool:
        if not file_paths:
            self._report_error("未选择任何文件")
            return False

        all_data = []
        total_files = len(file_paths)
        
        try:
            for idx, file_path in enumerate(file_paths):
                self._report_progress(
                    int((idx / total_files) * 80),
                    f"正在处理文件: {os.path.basename(file_path)}"
                )
                
                if not os.path.exists(file_path):
                    self._report_error(f"文件不存在: {file_path}")
                    continue
                
                try:
                    if merge_type == "first_sheet":
                        df = pd.read_excel(
                            file_path,
                            skiprows=skip_rows,
                            header=header_row
                        )
                        if add_source_column:
                            df["来源文件"] = os.path.basename(file_path)
                        all_data.append(df)
                    
                    elif merge_type == "all_sheets":
                        excel_file = pd.ExcelFile(file_path)
                        for sn in excel_file.sheet_names:
                            df = pd.read_excel(
                                excel_file,
                                sheet_name=sn,
                                skiprows=skip_rows,
                                header=header_row
                            )
                            if add_source_column:
                                df["来源文件"] = os.path.basename(file_path)
                                df["来源工作表"] = sn
                            all_data.append(df)
                    
                    elif merge_type == "specific_sheet" and sheet_name:
                        df = pd.read_excel(
                            file_path,
                            sheet_name=sheet_name,
                            skiprows=skip_rows,
                            header=header_row
                        )
                        if add_source_column:
                            df["来源文件"] = os.path.basename(file_path)
                        all_data.append(df)
                
                except Exception as e:
                    self._report_error(f"读取文件 {file_path} 时出错: {str(e)}")
                    continue
            
            if not all_data:
                self._report_error("没有成功读取任何数据")
                return False
            
            self._report_progress(85, "正在合并数据...")
            merged_df = pd.concat(all_data, ignore_index=True)
            
            self._report_progress(90, "正在保存文件...")
            merged_df.to_excel(output_path, index=False, engine='openpyxl')
            
            self._report_progress(100, f"合并完成! 共合并 {len(all_data)} 个数据集")
            return True
        
        except Exception as e:
            self._report_error(f"合并过程中发生错误: {str(e)}")
            return False

    def merge_sheets_in_file(
        self,
        file_path: str,
        output_path: str,
        add_sheet_column: bool = False
    ) -> bool:
        if not os.path.exists(file_path):
            self._report_error(f"文件不存在: {file_path}")
            return False
        
        try:
            self._report_progress(10, "正在读取工作簿...")
            excel_file = pd.ExcelFile(file_path)
            
            all_data = []
            total_sheets = len(excel_file.sheet_names)
            
            for idx, sheet_name in enumerate(excel_file.sheet_names):
                self._report_progress(
                    int(10 + (idx / total_sheets) * 70),
                    f"正在处理工作表: {sheet_name}"
                )
                
                df = pd.read_excel(excel_file, sheet_name=sheet_name)
                if add_sheet_column:
                    df["来源工作表"] = sheet_name
                all_data.append(df)
            
            self._report_progress(85, "正在合并数据...")
            merged_df = pd.concat(all_data, ignore_index=True)
            
            self._report_progress(90, "正在保存文件...")
            merged_df.to_excel(output_path, index=False, engine='openpyxl')
            
            self._report_progress(100, f"合并完成! 共合并 {total_sheets} 个工作表")
            return True
        
        except Exception as e:
            self._report_error(f"合并工作表时发生错误: {str(e)}")
            return False
