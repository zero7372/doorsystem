import pandas as pd
import xlsxwriter
import pandas as pd

class ExcelExporter:
    def __init__(self):
        # 表頭中英文映射字典
        self.column_mapping = {
            'date': '日期',
            'weekday': '星期',
            'emp_name': '姓名',
            'check_in': '上班時間',
            'check_out': '下班時間',
            'status': '狀態'
        }
        
    def export_to_excel(self, file_path, all_processed_data):
        """將數據導出到Excel文件，為每個員工建立一個工作表，按要求進行客製化設置"""
        # 建立一個ExcelWriter對象
        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            # 準備要導出的數據（移除不需要的列）
            processed_data_list = []
            for record in all_processed_data:
                # 只保留需要的列，並添加原始數據用於條件判斷
                filtered_record = {
                    'date': record['date'],
                    'weekday': record['weekday'],
                    'emp_name': record['emp_name'],
                    'check_in': record['check_in'],
                    'check_out': record['check_out'],
                    'status': record['status'],
                    'is_weekend': record.get('is_weekend', False),  # 保留用於判斷，但不導出
                    'original_status': record['status']  # 保留原始狀態用於判斷
                }
                processed_data_list.append(filtered_record)
            
            # 建立一個全局工作表
            self._create_custom_worksheet(writer, '全部記錄', processed_data_list)
            
            # 為每個員工建立一個工作表
            employee_names = sorted(list(set([record['emp_name'] for record in all_processed_data])))
            
            # 建立統計工作表
            stats_df = self._create_statistics_dataframe(all_processed_data)
            stats_df.to_excel(writer, sheet_name='統計資訊', index=False)
            
            # 為每個員工建立工作表
            for emp_name in employee_names:
                # 過濾該員工的數據
                emp_data = [record for record in processed_data_list if record['emp_name'] == emp_name]
                
                # 截取工作表名稱，Excel工作表名稱最長31個字符
                sheet_name = emp_name[:31]
                
                # 建立自定義工作表
                self._create_custom_worksheet(writer, sheet_name, emp_data)
            
            print(f"成功導出{len(employee_names)+2}個工作表到Excel文件")
    
    def _create_custom_worksheet(self, writer, sheet_name, data_list):
        """建立客製化的工作表，實現凍結窗格、條件格式化等功能"""
        # 建立一個臨時DataFrame用於獲取列名
        temp_df = pd.DataFrame(data_list)
        
        # 只選擇需要導出的列
        columns_to_export = ['date', 'weekday', 'emp_name', 'check_in', 'check_out', 'status']
        
        # 建立一個新的DataFrame，只包含需要的列，並轉換表頭為中文
        df = pd.DataFrame([{self.column_mapping[col]: row[col] for col in columns_to_export} for row in data_list])
        
        # 獲取xlsxwriter的workbook和worksheet對象
        workbook = writer.book
        worksheet = writer.sheets.get(sheet_name)
        
        # 创建格式
        # 週六週日行的黃色背景
        weekend_format = workbook.add_format({
            'bg_color': '#FFFF99',
            'border': 1,  # 新增邊框
            'border_color': '#000000'
        })
        # 遲到、早退、外出記錄的紅色文字
        special_status_format = workbook.add_format({
            'font_color': '#FF0000',
            'border': 1,  # 新增邊框
            'border_color': '#000000'
        })
        # 普通儲存格格式（帶邊框）
        default_format = workbook.add_format({
            'border': 1,
            'border_color': '#000000'
        })
        # 表頭儲存格格式（帶邊框和加粗）
        header_format = workbook.add_format({
            'border': 1,
            'border_color': '#000000',
            'bold': True,
            'bg_color': '#D3D3D3'  # 淺灰色背景，增強表頭識別度
        })
        
        # 如果工作表不存在，先建立它
        if worksheet is None:
            worksheet = workbook.add_worksheet(sheet_name)
            
            # 先設定凍結窗格（在寫入任何資料之前）
            worksheet.freeze_panes(1, 0)  # 參數為(行, 列)，表示凍結0行以上和1列以左的區域，即凍結第一列(A列)
                
            # 寫入表頭（使用表頭格式）
            header_row = 0
            for col_num, col_name in enumerate(df.columns):
                worksheet.write(header_row, col_num, col_name, header_format)
                
            # 寫入資料，並同時應用格式
            for row_num, (record, row_data) in enumerate(zip(data_list, df.itertuples(index=False, name=None)), start=1):
                # 檢查是否是週末行
                if record.get('is_weekend', False):
                    # 對週末行應用黃色背景（已包含邊框）
                    for col_num, cell_value in enumerate(row_data):
                        worksheet.write(row_num, col_num, cell_value, weekend_format)
                else:
                    # 檢查是否包含特殊狀態（遲到、早退、外出）
                    status = record.get('original_status', '')
                    if '遲到' in status or '早退' in status or '外出' in status:
                        # 對特殊狀態行應用紅色文字（已包含邊框）
                        for col_num, cell_value in enumerate(row_data):
                            worksheet.write(row_num, col_num, cell_value, special_status_format)
                    else:
                        # 普通行，使用帶邊框的預設格式
                        for col_num, cell_value in enumerate(row_data):
                            worksheet.write(row_num, col_num, cell_value, default_format)
        
        # 自動調整欄寬
        for col_num, col_name in enumerate(df.columns):
            # 計算欄的最大寬度
            max_width = len(col_name)  # 至少為列名長度
            for row_num in range(len(data_list)):
                cell_value = str(df.iloc[row_num, col_num])
                cell_width = len(cell_value)
                if cell_width > max_width:
                    max_width = cell_width
            
            # 設定欄寬（加一點餘量，確保有足夠的顯示空間）
            worksheet.set_column(col_num, col_num, max_width + 5)  # 增加更多餘量以確保內容完整顯示
    
    def _create_statistics_dataframe(self, all_processed_data):
        """建立統計資訊DataFrame"""
        # 按員工分組統計
        employee_stats = []
        employee_names = sorted(list(set([record['emp_name'] for record in all_processed_data])))
        
        for emp_name in employee_names:
            emp_data = [record for record in all_processed_data if record['emp_name'] == emp_name]
            
            # 統計不同狀態的數量
            status_counts = {}
            for record in emp_data:
                status = record['status']
                status_counts[status] = status_counts.get(status, 0) + 1
            
            # 計算正常、遲到、早退的數量
            normal_count = status_counts.get('正常', 0)
            late_count = status_counts.get('遲到', 0) + status_counts.get('迟到', 0) + (status_counts.get('假日、遲到', 0) if '假日、遲到' in status_counts else 0)
            early_leave_count = status_counts.get('早退', 0) + (status_counts.get('假日、早退', 0) if '假日、早退' in status_counts else 0)
            absent_count = status_counts.get('未進公司', 0)
            out_count = status_counts.get('外出', 0)
            holiday_count = status_counts.get('假日', 0)
            
            # 新增到統計結果
            employee_stats.append({
                '員工姓名': emp_name,
                '總記錄數': len(emp_data),
                '正常': normal_count,
                '遲到': late_count,
                '早退': early_leave_count,
                '未進公司': absent_count,
                '外出': out_count,
                '假日': holiday_count
            })
        
        # 建立DataFrame
        stats_df = pd.DataFrame(employee_stats)
        
        # 新增總計行
        total_row = {
            '員工姓名': '總計',
            '總記錄數': stats_df['總記錄數'].sum(),
            '正常': stats_df['正常'].sum(),
            '遲到': stats_df['遲到'].sum(),
            '早退': stats_df['早退'].sum(),
            '未進公司': stats_df['未進公司'].sum(),
            '外出': stats_df['外出'].sum(),
            '假日': stats_df['假日'].sum()
        }
        
        # 將總計行新增到DataFrame
        stats_df = pd.concat([stats_df, pd.DataFrame([total_row])], ignore_index=True)
        
        return stats_df

# 建立一個單例實例，方便其他模組直接使用
excel_exporter = ExcelExporter()