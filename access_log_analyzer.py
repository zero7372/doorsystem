import flet as ft
import pandas as pd
import numpy as np
from datetime import datetime, time
import os
import xlsxwriter
from excel_exporter import excel_exporter

class AccessLogAnalyzer:
    def __init__(self, page):
        self.page = page
        self.page.title = "門禁日誌分析器"
        self.page.theme_mode = ft.ThemeMode.DARK  # 設置為深色主題
        self.page.bgcolor = "#1e1e1e"  # 深灰色背景
        self.page.window_width = 1000
        self.page.window_height = 700
        
        # 用於存儲編號與姓名的映射關係
        self.id_name_map = {}
        # 开启debug模式
        self.debug_mode = False
        
        # 存储处理后的数据，用于筛选
        self.all_processed_data = []
        # 当前选中的员工姓名
        self.selected_name = None
        # 员工姓名列表
        self.employee_names = []
        
        # 用於顯示統計信息的Text組件
        self.stats_text = ft.Text("統計信息: 無數據", color="#cccccc", size=14)
        
        # 創建UI組件
        self.create_ui()
    
    def create_ui(self):
        # 檔案選擇器
        self.file_picker = ft.FilePicker(on_result=self.on_file_selected)
        self.page.overlay.append(self.file_picker)
        
        # 保存文件选择器（用于导出Excel）
        self.save_file_picker = ft.FilePicker(on_result=self.on_save_file_selected)
        self.page.overlay.append(self.save_file_picker)
        
        # 標題
        title = ft.Text("門禁日誌分析器", size=24, weight=ft.FontWeight.BOLD, color="#ffffff")
        
        # 選擇檔案按鈕
        select_file_btn = ft.ElevatedButton(
            text="選擇門禁日誌文件",
            on_click=lambda _: self.file_picker.pick_files(
                allowed_extensions=["csv"],
                file_type=ft.FilePickerFileType.CUSTOM,
                dialog_title="選擇門禁日誌CSV文件"
            ),
            bgcolor="#374151",  # 按鈕背景色
            color="#ffffff",    # 按鈕文字顏色
        )
        
        # 导出Excel按钮
        self.export_excel_btn = ft.ElevatedButton(
            text="导出Excel文件",
            on_click=self.on_export_excel,
            bgcolor="#374151",  # 按鈕背景色
            color="#ffffff",    # 按鈕文字顏色
            disabled=True  # 默认禁用，数据加载后启用
        )
        
        # 分析結果表格
        self.columns = [
            ft.DataColumn(ft.Text("日期", color="#ffffff")),
            ft.DataColumn(ft.Text("星期", color="#ffffff")),
            ft.DataColumn(ft.Text("姓名", color="#ffffff")),
            ft.DataColumn(ft.Text("上班時間", color="#ffffff")),
            ft.DataColumn(ft.Text("下班時間", color="#ffffff")),
            ft.DataColumn(ft.Text("狀態", color="#ffffff")),
        ]
        self.data_table = ft.DataTable(
            columns=self.columns,
            rows=[],
            heading_row_color="#2d3748",  # 表头背景色
            heading_row_height=40,
            data_row_min_height=36,
        )
        
        # 狀態顯示
        self.status = ft.Text("等待選擇文件...", color="#cccccc", size=14)
        
        # 名字篩選下拉菜單
        self.name_filter_label = ft.Text("篩選名字: ", color="#ffffff")
        self.name_filter = ft.Dropdown(
            options=[
                ft.dropdown.Option("全部顯示")
            ],
            value="全部顯示",
            on_change=self.on_name_selected,
            bgcolor="#374151",
            color="#ffffff",
            width=200
        )
        
        # 创建一个滚动视图来包裹表格
        scrollable_table = ft.ListView(
            controls=[self.data_table],
            expand=True,
            auto_scroll=False,
        )
        
        # 布局
        self.page.add(
            ft.Column(
                [
                    ft.Row([title], alignment=ft.MainAxisAlignment.CENTER),
                    ft.Row([select_file_btn, self.export_excel_btn], alignment=ft.MainAxisAlignment.CENTER, height=60, spacing=20),
                    ft.Row([self.name_filter_label, self.name_filter], alignment=ft.MainAxisAlignment.START, height=40, spacing=10),
                    scrollable_table,
                    ft.Row([self.stats_text], alignment=ft.MainAxisAlignment.START, height=30),
                    ft.Row([self.status], alignment=ft.MainAxisAlignment.START, height=30),
                ],
                expand=True,
                spacing=10,
            )
        )
    
    def on_name_selected(self, e):
        """處理名字篩選選擇事件"""
        selected_value = e.control.value
        
        if selected_value == "全部顯示":
            self.selected_name = None
            # 顯示所有數據
            self.display_results(self.all_processed_data)
            self.status.value = f"顯示全部 {len(self.all_processed_data)} 條記錄"
        else:
            self.selected_name = selected_value
            # 根據選中的名字篩選數據，並排除周末無記錄的數據
            # 保留非周末的所有記錄，以及周末但有記錄（不是"未進公司"狀態）的記錄
            filtered_data = [
                record for record in self.all_processed_data 
                if record['emp_name'] == selected_value 
                and (not record['is_weekend'] or ('status' in record and '未進公司' not in record['status']))
            ]
            self.display_results(filtered_data)
            self.status.value = f"顯示 {selected_value} 的 {len(filtered_data)} 條記錄"
        
        self.status.color = "#4ade80"  # 綠色
        self.page.update()
    
    def on_file_selected(self, e):
        if e.files:
            file_path = e.files[0].path
            try:
                # 顯示加載中狀態
                self.status.value = "正在載入和分析數據..."
                self.status.color = "#cccccc"
                self.page.update()
                
                if self.debug_mode:
                    print(f"\n===== DEBUG: 选择的文件路径: {file_path}")
                    print(f"DEBUG: 文件是否存在: {os.path.exists(file_path)}")
                    print(f"DEBUG: 文件大小: {os.path.getsize(file_path)} 字节")
                
                # 加载并处理数据
                data = self.load_data(file_path)
                processed_data = self.process_data(data)
                
                # 保存所有处理后的数据
                self.all_processed_data = processed_data
                
                # 更新名字筛选下拉菜单选项
                self.employee_names = sorted(list(set([record['emp_name'] for record in processed_data])))
                self.name_filter.options = [ft.dropdown.Option("全部显示")] + [ft.dropdown.Option(name) for name in self.employee_names]
                self.name_filter.value = "全部显示"
                self.selected_name = None
                
                # 启用导出Excel按钮
                self.export_excel_btn.disabled = False
                
                # 显示结果
                self.display_results(processed_data)
                
                # 更新狀態
                self.status.value = f"分析完成，共 {len(processed_data)} 條記錄"
                self.status.color = "#4ade80"  # 綠色
                self.page.update()
                
            except Exception as ex:
                self.status.value = f"分析出錯: {str(ex)}"
                self.status.color = "#ef4444"  # 紅色
                self.page.update()
                if self.debug_mode:
                    import traceback
                    print(f"\n===== DEBUG: 分析过程异常 ====")
                    traceback.print_exc()
    
    def load_data(self, file_path):
        # 读取CSV文件，处理列数不一致的问题
        print(f"开始读取文件: {file_path}")
        
        # 尝试不同的编码读取文件
        encodings = ['utf-8', 'gbk', 'latin1']
        df = None
        
        # 正确的列名（根据CSV文件表头）
        expected_columns = ['序號', '記錄時間', '編號', '姓名', '允許通行', '詳細資訊']
        
        for encoding in encodings:
            try:
                # 首先读取前几行来检查列数
                with open(file_path, 'r', encoding=encoding) as f:
                    header_line = f.readline().strip()
                    first_data_line = f.readline().strip()
                    
                    # 检查列数
                    header_cols = header_line.split(',')
                    data_cols = first_data_line.split(',')
                    
                    print(f"使用{encoding}编码读取的列信息: 表头{len(header_cols)}列, 数据{len(data_cols)}列")
                    
                    # 处理列数不一致的情况
                    if len(data_cols) > len(header_cols):
                        # 为额外的列创建临时名称
                        additional_cols = [f'临时列{i}' for i in range(len(data_cols) - len(header_cols))]
                        actual_columns = header_cols + additional_cols
                        print(f"处理列数不一致: 添加了{len(additional_cols)}个临时列")
                    else:
                        actual_columns = header_cols[:len(data_cols)]
                        print(f"处理列数不一致: 只使用前{len(actual_columns)}个表头列")
                
                # 重新读取整个文件，指定正确的列名
                df = pd.read_csv(
                    file_path, 
                    encoding=encoding, 
                    header=0,  # 使用第一行作为表头
                    names=actual_columns,  # 指定实际列名
                    on_bad_lines='skip'  # 跳过格式错误的行
                )
                
                print(f"成功使用{encoding}编码读取文件，形状: {df.shape}")
                break
            except Exception as e:
                print(f"使用{encoding}编码读取失败: {str(e)}")
        
        if df is None:
            raise Exception("无法读取CSV文件，尝试了多种编码")
        
        # 显示前几行数据用于调试
        print("文件前5行数据:")
        print(df.head())
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 列信息 ====")
            print(f"所有列名: {list(df.columns)}")
            print(f"列数据类型:\n{df.dtypes}")
        
        # 检查是否包含必需的列
        required_columns = ['記錄時間', '編號', '姓名']
        for col in required_columns:
            if col not in df.columns:
                raise Exception(f"CSV文件中未找到必需的列: {col}")
        
        # 创建编号与姓名的映射
        # 确保只使用有效的编号和姓名对
        valid_pairs = df[df['編號'].notna() & df['姓名'].notna() & (df['姓名'] != '是')]
        self.id_name_map = dict(zip(valid_pairs['編號'], valid_pairs['姓名']))
        print(f"创建了编号-姓名映射，共{len(self.id_name_map)}个条目")
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 编号-姓名映射前5条 ====")
            for i, (emp_id, name) in enumerate(list(self.id_name_map.items())[:5]):
                print(f"{emp_id}: {name}")
        
        # 日期时间解析 - 尝试多种格式
        print("开始解析日期时间...")
        datetime_formats = [
            '%Y-%m-%d %H:%M:%S',  # 标准格式
            '%Y/%m/%d %H:%M:%S',
            '%Y-%m-%d %H:%M',
            '%Y/%m/%d %H:%M',
            '%d-%m-%Y %H:%M:%S',
            '%d/%m/%Y %H:%M:%S',
        ]
        
        # 初始化datetime列为空
        df['datetime'] = pd.NaT
        
        for fmt in datetime_formats:
            try:
                # 保存当前的有效解析结果
                previous_valid = df['datetime'].notna().sum()
                
                # 使用pd.to_datetime并指定format
                temp_datetime = pd.to_datetime(df['記錄時間'], format=fmt, errors='coerce')
                
                # 更新datetime列，但只保留新的有效解析结果
                df.loc[temp_datetime.notna(), 'datetime'] = temp_datetime[temp_datetime.notna()]
                
                # 计算解析成功率
                success_rate = df['datetime'].notna().sum() / len(df)
                new_valid = df['datetime'].notna().sum() - previous_valid
                print(f"尝试格式'{fmt}'，新增有效: {new_valid}, 总成功率: {success_rate:.2%} ({df['datetime'].notna().sum()}/{len(df)})")
                
                # 如果所有记录都已成功解析，提前退出
                if df['datetime'].notna().all():
                    print("所有记录均已成功解析，停止尝试其他格式")
                    break
            except Exception as e:
                print(f"尝试格式'{fmt}'时出错: {str(e)}")
        
        # 如果还有未解析的记录，尝试自动解析
        if not df['datetime'].notna().all():
            print("仍有未解析的记录，尝试自动解析...")
            # 保存当前的有效解析结果
            previous_valid = df['datetime'].notna().sum()
            
            # 尝试自动解析剩余的记录
            temp_datetime = pd.to_datetime(df.loc[df['datetime'].isna(), '記錄時間'], errors='coerce')
            df.loc[temp_datetime.notna().index, 'datetime'] = temp_datetime
            
            new_valid = df['datetime'].notna().sum() - previous_valid
            success_rate = df['datetime'].notna().sum() / len(df)
            print(f"自动解析新增有效: {new_valid}, 最终成功率: {success_rate:.2%} ({df['datetime'].notna().sum()}/{len(df)})")
        
        # 统计解析结果
        valid_count = df['datetime'].notna().sum()
        total_count = len(df)
        print(f"最终日期时间解析结果: 有效 {valid_count}/{total_count}")
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 日期时间解析样本 ====")
            # 显示前5个解析成功的记录
            valid_samples = df[df['datetime'].notna()].head()
            if not valid_samples.empty:
                for i, row in valid_samples.iterrows():
                    print(f"原始值: {row['記錄時間']} -> 解析后: {row['datetime']}")
            
            # 显示前5个解析失败的记录（如果有）
            invalid_samples = df[df['datetime'].isna()].head()
            if not invalid_samples.empty:
                print(f"\n解析失败的样本:")
                for i, row in invalid_samples.iterrows():
                    print(f"原始值: {row['記錄時間']}")
        
        # 如果没有有效数据，抛出异常
        if valid_count == 0:
            # 显示前几个日期时间值用于调试
            if total_count > 0:
                sample_datetimes = df['記錄時間'].head().tolist()
                print(f"日期时间样本值: {sample_datetimes}")
            raise Exception("没有有效的日期时间数据")
        
        # 过滤掉无效的日期时间记录
        df = df.dropna(subset=['datetime'])
        
        # 添加日期列，用于分组
        df['date'] = df['datetime'].dt.date
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 过滤后的数据信息 ====")
            print(f"过滤后的数据形状: {df.shape}")
            print(f"日期范围: {df['date'].min()} 至 {df['date'].max()}")
            print(f"唯一日期数量: {df['date'].nunique()}")
            print(f"唯一编号数量: {df['編號'].nunique()}")
        
        return df
    
    def process_data(self, data):
        print("开始处理数据...")
        results = []
        
        # 获取日期范围
        min_date = data['date'].min()
        max_date = data['date'].max()
        
        # 创建日期范围
        date_range = pd.date_range(start=min_date, end=max_date, freq='D').date.tolist()
        
        # 按日期和编号分组
        grouped = data.groupby(['date', '編號'])
        
        # 标准上班和下班时间
        standard_check_in = time(9, 0)
        standard_check_out = time(18, 0)
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 分组处理详情 ====")
            print(f"总组数: {len(grouped)}")
        
        group_count = 0
        for (date, emp_id), group in grouped:
            group_count += 1
            
            if self.debug_mode and group_count <= 3:  # 只显示前3组的详情
                print(f"\n组 {group_count}: 日期={date}, 编号={emp_id}, 记录数={len(group)}")
                print(f"该组原始记录:\n{group[['datetime', '編號', '姓名']].to_string(index=False)}")
            
            # 按时间排序
            sorted_group = group.sort_values('datetime')
            
            # 获取最早的记录作为上班时间，最晚的记录作为下班时间
            check_in_time = sorted_group.iloc[0]['datetime'].time()
            check_out_time = sorted_group.iloc[-1]['datetime'].time()
            
            # 获取员工姓名
            emp_name = self.id_name_map.get(emp_id, str(emp_id))
            
            # 计算星期几
            weekday_map = {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'}
            weekday = weekday_map[date.weekday()]
            is_weekend = date.weekday() in [5, 6]  # 周六或周日
            
            # 判斷狀態：1筆記錄標記為外出
            if len(group) == 1:
                status_text = "外出"
            else:
                # 判斷是否遲到或早退
                status = []
                if check_in_time > standard_check_in:
                    status.append("遲到")
                if check_out_time < standard_check_out:
                    status.append("早退")
                
                # 如果是周末，添加假日標記
                if is_weekend:
                    status.append("假日")

                status_text = "、".join(status) if status else "正常"
            
            # 格式化日期显示
            formatted_date = date.strftime('%Y-%m-%d')
            
            if self.debug_mode and group_count <= 3:
                print(f"上班时间: {check_in_time} (标准: {standard_check_in}) -> {'迟到' if check_in_time > standard_check_in else '正常'}")
                print(f"下班时间: {check_out_time} (标准: {standard_check_out}) -> {'早退' if check_out_time < standard_check_out else '正常'}")
                print(f"最终状态: {status_text}")
            
            # 计算星期几
            weekday_map = {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'}
            weekday = weekday_map[date.weekday()]
            is_weekend = date.weekday() in [5, 6]  # 周六或周日
            
            if self.debug_mode and group_count <= 3:
                print(f"日期 {formatted_date} 是 {weekday}, {'周末' if is_weekend else '工作日'}")
            
            # 添加到结果
            results.append({
                'date': formatted_date,
                'weekday': weekday,
                'is_weekend': is_weekend,
                'emp_id': emp_id,
                'emp_name': emp_name,
                'check_in': check_in_time.strftime('%H:%M'),
                'check_out': check_out_time.strftime('%H:%M'),
                'status': status_text
            })
        
        # 收集所有员工和日期的组合
        all_employee_dates = set()
        for record in results:
            all_employee_dates.add((record['date'], record['emp_name']))
            
        # 收集所有唯一员工姓名
        all_employees = set(record['emp_name'] for record in results)
        
        # 为每个员工和每一天检查是否有记录，如果没有则添加"未进公司"记录
        missing_records = []
        for employee in all_employees:
            for date in date_range:
                date_str = date.strftime('%Y-%m-%d')
                if (date_str, employee) not in all_employee_dates:
                    # 计算星期几
                    weekday_map = {0: '周一', 1: '周二', 2: '周三', 3: '周四', 4: '周五', 5: '周六', 6: '周日'}
                    weekday = weekday_map[date.weekday()]
                    is_weekend = date.weekday() in [5, 6]  # 周六或周日
                    
                    missing_records.append({
                        'date': date_str,
                        'weekday': weekday,
                        'is_weekend': is_weekend,
                        'emp_id': '',
                        'emp_name': employee,
                        'check_in': '-',
                        'check_out': '-',
                        'status': '未進公司'
                    })
        
        # 合并结果
        results.extend(missing_records)
        
        # 按日期和姓名排序
        results.sort(key=lambda x: (x['date'], x['emp_name']))
        
        if self.debug_mode:
            print(f"\n===== DEBUG: 处理结果样本 ====")
            for i, record in enumerate(results[:5]):
                print(f"记录 {i+1}: {record}")
            
            # 统计未进公司的记录
            absent_count = sum(1 for record in results if record['status'] == '未进公司')
            print(f"未进公司记录数: {absent_count}")
        
        print(f"数据处理完成，共生成{len(results)}条记录")
        return results
        
    def on_export_excel(self, e):
        """处理导出Excel按钮点击事件"""
        if not self.all_processed_data:
            self.status.value = "没有数据可导出"
            self.status.color = "#ef4444"  # 红色
            self.page.update()
            return
        
        # 打开保存文件对话框
        self.save_file_picker.save_file(
            file_type=ft.FilePickerFileType.CUSTOM,
            allowed_extensions=["xlsx"],
            dialog_title="保存Excel文件"
        )
        
    def on_save_file_selected(self, e):
        """处理保存文件对话框的结果"""
        if e.path:
            # 确保文件路径包含.xlsx扩展名
            file_path = e.path
            if not file_path.lower().endswith('.xlsx'):
                file_path += '.xlsx'
            try:
                # 显示导出中状态
                self.status.value = "正在导出Excel文件..."
                self.status.color = "#cccccc"
                self.page.update()
                
                # 导出Excel（使用独立的excel_exporter模块）
                excel_exporter.export_to_excel(file_path, self.all_processed_data)
                
                # 更新状态
                self.status.value = f"Excel文件已成功导出到: {file_path}"
                self.status.color = "#4ade80"  # 绿色
                self.page.update()
                
            except Exception as ex:
                self.status.value = f"导出Excel出错: {str(ex)}"
                self.status.color = "#ef4444"  # 红色
                self.page.update()
                if self.debug_mode:
                    import traceback
                    print(f"\n===== DEBUG: 导出Excel异常 ====")
                    traceback.print_exc()
    
    def calculate_statistics(self, processed_data):
        """計算並顯示數據統計信息"""
        if not processed_data:
            self.stats_text.value = "統計信息: 無數據"
            return
        
        # 統計不同狀態的數量
        status_counts = {}
        for record in processed_data:
            status = record['status']
            status_counts[status] = status_counts.get(status, 0) + 1
        
        # 計算上班和下班時間相關統計
        check_in_times = []
        check_out_times = []
        late_count = 0
        early_leave_count = 0
        normal_count = 0
        
        for record in processed_data:
            if record['check_in'] != '-' and ':' in record['check_in']:
                check_in_time = datetime.strptime(record['check_in'], '%H:%M').time()
                check_in_times.append(check_in_time)
                # 檢查是否遲到（9:00以後）
                if check_in_time > time(9, 0):
                    late_count += 1
            
            if record['check_out'] != '-' and ':' in record['check_out']:
                check_out_time = datetime.strptime(record['check_out'], '%H:%M').time()
                check_out_times.append(check_out_time)
                # 檢查是否早退（18:00以前）
                if check_out_time < time(18, 0):
                    early_leave_count += 1
            
            if record['status'] == '正常':
                normal_count += 1
        
        # 生成統計文本
        stats = f"統計信息: 總記錄 {len(processed_data)}, 正常 {normal_count}, 遲到 {late_count}, 早退 {early_leave_count}"
        
        # 添加其他狀態統計
        other_statuses = [f"{status}: {count}" for status, count in status_counts.items() if status not in ['正常', '遲到', '早退']]
        if other_statuses:
            stats += f", {', '.join(other_statuses)}"
        
        self.stats_text.value = stats
        self.stats_text.color = "#4ade80"  # 綠色
    
    def display_results(self, processed_data):
        # 清空现有行
        self.data_table.rows.clear()
        
        # 添加新行
        for record in processed_data:
            # 强制确保status值正确显示
            status_display = record['status'] or "未知状态"
            
            # 检查是否需要整行红色字体
            if status_display == "外出" or status_display == "未进公司" or "迟到" in status_display or "遲到" in status_display:
                row_text_color = ft.Colors.RED
            else:
                row_text_color = ft.Colors.WHITE
                
            # 检查状态并设置不同的颜色
            if status_display == "正常":
                status_color = ft.Colors.GREEN_300
            elif "迟到" in status_display or "遲到" in status_display:
                status_color = ft.Colors.RED
            elif "早退" in status_display:
                status_color = ft.Colors.YELLOW
            elif status_display == "外出":
                status_color = ft.Colors.PURPLE
            elif status_display == "未进公司":
                status_color = ft.Colors.GREY
            else:
                status_color = ft.Colors.GREY
            
            # 打印调试信息
            if self.debug_mode and len(self.data_table.rows) < 5:
                print(f"添加记录到UI: 日期={record['date']}, 编号={record['emp_id']}, 状态={status_display}")
            
            # 设置周末行的背景色
            row_color = ft.Colors.with_opacity(0.3, ft.Colors.AMBER_700) if record.get('is_weekend', False) else None
            
            # 设置星期几的文本颜色
            weekday_color = ft.Colors.AMBER_300 if record.get('is_weekend', False) else ft.Colors.WHITE
            
            self.data_table.rows.append(
                ft.DataRow(
                    color=row_color,
                    cells=[
                        ft.DataCell(ft.Text(record['date'], color=row_text_color)),
                        ft.DataCell(ft.Text(record.get('weekday', ''), color=weekday_color)),
                        ft.DataCell(ft.Text(record['emp_name'], color=row_text_color)),
                        ft.DataCell(ft.Text(record['check_in'], color=row_text_color)),
                        ft.DataCell(ft.Text(record['check_out'], color=row_text_color)),
                        ft.DataCell(ft.Text(status_display, color=status_color)),
                    ]
                )
            )
        
        # 计算并显示统计信息
        self.calculate_statistics(processed_data)
        
        # 强制页面更新两次以确保UI正确渲染
        self.page.update()
        # 短暂延迟后再次更新
        self.page.update()
        
        if self.debug_mode:
            print(f"已将{len(processed_data)}条记录添加到UI表格")
            # 统计不同状态的数量
            status_counts = {}
            for record in processed_data:
                status = record['status']
                status_counts[status] = status_counts.get(status, 0) + 1
            print(f"状态统计: {status_counts}")

def main(page):
    analyzer = AccessLogAnalyzer(page)

import flet as ft

if __name__ == "__main__":
    ft.app(target=main)