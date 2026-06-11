import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import shutil
import pandas as pd
from datetime import datetime

from core import DataValidator, DataMerger, ReminderGenerator, DataExporter, ReportGenerator


class MainWindow:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("公共卫生管理自动化工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)

        self.current_category = tk.StringVar(value='慢病')
        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.file_list = []
        self.file_info_list = []
        self.current_data = None
        self.validation_result = None
        self.merged_data = None
        self.deduped_data = None
        self.overdue_result = None
        self.next_month_result = None
        self.export_result = None
        self.report_generated = False
        self.reminder_generated = False
        self.archive_generated = False
        self.archive_path = ''

        self._setup_ui()
        self._setup_styles()

    def _setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')

        style.configure('Title.TLabel', font=('微软雅黑', 16, 'bold'))
        style.configure('Section.TLabelframe.Label', font=('微软雅黑', 10, 'bold'))
        style.configure('Action.TButton', font=('微软雅黑', 10), padding=8)
        style.configure('Primary.TButton', font=('微软雅黑', 10, 'bold'), padding=10)

    def _setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="15")
        main_frame.pack(fill=tk.BOTH, expand=True)

        title = ttk.Label(main_frame, text="公共卫生管理自动化工具", style='Title.TLabel')
        title.pack(anchor=tk.W, pady=(0, 10))

        subtitle = ttk.Label(main_frame, text="重点人群随访资料整理工具", foreground='gray')
        subtitle.pack(anchor=tk.W, pady=(0, 15))

        self._setup_config_section(main_frame)
        self._setup_actions_section(main_frame)
        self._setup_log_section(main_frame)

    def _setup_config_section(self, parent):
        config_frame = ttk.LabelFrame(parent, text="基础配置", style='Section.TLabelframe', padding="10")
        config_frame.pack(fill=tk.X, pady=(0, 10))

        form_frame = ttk.Frame(config_frame)
        form_frame.pack(fill=tk.X)

        ttk.Label(form_frame, text="人群分类:").grid(row=0, column=0, sticky=tk.W, pady=5)
        categories = ['慢病', '孕产妇', '儿童', '老年人']
        category_combo = ttk.Combobox(form_frame, textvariable=self.current_category,
                                      values=categories, state='readonly', width=15)
        category_combo.grid(row=0, column=1, sticky=tk.W, padx=(5, 20), pady=5)
        category_combo.bind('<<ComboboxSelected>>', self._on_category_change)

        ttk.Label(form_frame, text="随访间隔:").grid(row=0, column=2, sticky=tk.W, pady=5)
        self.interval_label = ttk.Label(form_frame, text="90 天", foreground='blue')
        self.interval_label.grid(row=0, column=3, sticky=tk.W, padx=(5, 0), pady=5)

        ttk.Label(form_frame, text="数据文件夹:").grid(row=1, column=0, sticky=tk.W, pady=5)
        input_entry = ttk.Entry(form_frame, textvariable=self.input_folder, width=50)
        input_entry.grid(row=1, column=1, columnspan=3, sticky=tk.EW, padx=(5, 5), pady=5)
        ttk.Button(form_frame, text="浏览...", command=self._select_input_folder, width=10).grid(
            row=1, column=4, padx=(5, 0), pady=5)

        ttk.Label(form_frame, text="输出文件夹:").grid(row=2, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(form_frame, textvariable=self.output_folder, width=50)
        output_entry.grid(row=2, column=1, columnspan=3, sticky=tk.EW, padx=(5, 5), pady=5)
        ttk.Button(form_frame, text="浏览...", command=self._select_output_folder, width=10).grid(
            row=2, column=4, padx=(5, 0), pady=5)

        form_frame.columnconfigure(1, weight=1)

    def _setup_actions_section(self, parent):
        actions_frame = ttk.LabelFrame(parent, text="操作功能", style='Section.TLabelframe', padding="10")
        actions_frame.pack(fill=tk.X, pady=(0, 10))

        btns_frame = ttk.Frame(actions_frame)
        btns_frame.pack(fill=tk.X)

        actions = [
            ("1. 导入名单", self.action_import, '导入随访数据文件'),
            ("2. 校验字段", self.action_validate, '检查字段完整性'),
            ("3. 合并随访", self.action_merge, '合并重复随访记录'),
            ("4. 生成提醒", self.action_remind, '生成超期未访提醒'),
            ("5. 分组导出", self.action_export, '按村居分组导出'),
            ("6. 汇总报告", self.action_report, '生成汇总统计报告'),
        ]

        for i, (text, cmd, desc) in enumerate(actions):
            col = i % 3
            row = i // 3
            btn_frame = ttk.Frame(btns_frame)
            btn_frame.grid(row=row, column=col, padx=5, pady=5, sticky=tk.EW)

            btn = ttk.Button(btn_frame, text=text, command=cmd, style='Action.TButton')
            btn.pack(fill=tk.X)

            desc_label = ttk.Label(btn_frame, text=desc, foreground='gray', font=('微软雅黑', 8))
            desc_label.pack(fill=tk.X, pady=(3, 0))

        for col in range(3):
            btns_frame.columnconfigure(col, weight=1)

        all_btn_frame = ttk.Frame(actions_frame)
        all_btn_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(all_btn_frame, text="一键执行全部操作", command=self.action_all,
                   style='Primary.TButton').pack(fill=tk.X)

        archive_btn_frame = ttk.Frame(actions_frame)
        archive_btn_frame.pack(fill=tk.X, pady=(8, 0))

        ttk.Button(archive_btn_frame, text="生成月底归档包", command=self.action_archive,
                   style='Action.TButton').pack(fill=tk.X)

    def _setup_log_section(self, parent):
        bottom_frame = ttk.Frame(parent)
        bottom_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        log_frame = ttk.LabelFrame(bottom_frame, text="操作日志", style='Section.TLabelframe', padding="10")
        log_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.config(state=tk.DISABLED)

        overview_frame = ttk.LabelFrame(bottom_frame, text="当前分类概览", style='Section.TLabelframe', padding="10")
        overview_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))
        overview_frame.configure(width=260)

        self.overview_category_label = ttk.Label(overview_frame, text="慢病",
                                                   font=('微软雅黑', 12, 'bold'), foreground='blue')
        self.overview_category_label.pack(anchor=tk.W, pady=(0, 8))

        info_frame = ttk.Frame(overview_frame)
        info_frame.pack(fill=tk.X)

        self.overview_vars = {}

        overview_items = [
            ('已导入文件', '-', 'file_count'),
            ('总记录数', '-', 'total_records'),
            ('数据有效率', '-', 'valid_rate'),
            ('管理人数', '-', 'person_count'),
            ('超期未访', '-', 'overdue_count'),
            ('下月待随访', '-', 'next_month_count'),
            ('异常明细文件', '未生成', 'error_file'),
            ('汇总报告', '未生成', 'report_file'),
            ('下月提醒', '未生成', 'reminder_file'),
            ('归档包', '未生成', 'archive_file'),
        ]

        for label, default, key in overview_items:
            row_frame = ttk.Frame(info_frame)
            row_frame.pack(fill=tk.X, pady=2)

            label_widget = ttk.Label(row_frame, text=label + ':', font=('微软雅黑', 9))
            label_widget.pack(side=tk.LEFT)

            value_var = tk.StringVar(value=default)
            value_label = ttk.Label(row_frame, textvariable=value_var,
                                     font=('微软雅黑', 9, 'bold'), foreground='darkgreen')
            value_label.pack(side=tk.RIGHT)

            self.overview_vars[key] = value_var

        ttk.Separator(overview_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, pady=10)

        hint_label = ttk.Label(overview_frame, text="切换分类或重新导入时\n数据将自动刷新",
                               foreground='gray', font=('微软雅黑', 8), justify=tk.CENTER)
        hint_label.pack(anchor=tk.CENTER)

        self._log("系统就绪，请选择数据文件夹和人群分类")

    def _on_category_change(self, event=None):
        intervals = {'慢病': 90, '孕产妇': 30, '儿童': 90, '老年人': 180}
        category = self.current_category.get()
        self.interval_label.config(text=f"{intervals.get(category, 90)} 天")
        self.overview_category_label.config(text=category)
        self._reset_category_state()
        self._log(f"已切换到「{category}」分类工作台，数据已重置")
        if self.input_folder.get():
            self._scan_files()
        self._update_overview()

    def _reset_category_state(self):
        self.current_data = None
        self.validation_result = None
        self.merged_data = None
        self.deduped_data = None
        self.overdue_result = None
        self.next_month_result = None
        self.export_result = None
        self.report_generated = False
        self.reminder_generated = False
        self.archive_generated = False
        self.archive_path = ''
        self.file_info_list = []

    def _reset_import_state(self):
        self.validation_result = None
        self.deduped_data = None
        self.overdue_result = None
        self.next_month_result = None
        self.export_result = None
        self.report_generated = False
        self.reminder_generated = False
        self.archive_generated = False
        self.archive_path = ''

    def _update_overview(self):
        category = self._get_category()
        self.overview_category_label.config(text=category)

        vars_map = self.overview_vars

        vars_map['file_count'].set(str(len(self.file_list)) if self.file_list else '-')

        if self.current_data is not None:
            vars_map['total_records'].set(str(len(self.current_data)))
        else:
            vars_map['total_records'].set('-')

        if self.validation_result:
            vars_map['valid_rate'].set(self.validation_result.get('有效率', '-'))
        else:
            vars_map['valid_rate'].set('-')

        if self.deduped_data is not None:
            from core import DataMerger
            merger = DataMerger(category)
            person_result = merger.merge_by_person(self.deduped_data)
            vars_map['person_count'].set(str(person_result['人数']))
        elif self.current_data is not None:
            from core import DataMerger
            merger = DataMerger(category)
            person_result = merger.merge_by_person(self.current_data)
            vars_map['person_count'].set(str(person_result['人数']))
        else:
            vars_map['person_count'].set('-')

        if self.overdue_result:
            vars_map['overdue_count'].set(f"{self.overdue_result.get('超期人数', 0)} 人")
        else:
            vars_map['overdue_count'].set('-')

        if self.next_month_result:
            vars_map['next_month_count'].set(f"{self.next_month_result.get('待随访人数', 0)} 人")
        else:
            vars_map['next_month_count'].set('-')

        if self.validation_result and self.validation_result.get('错误明细'):
            vars_map['error_file'].set('已生成')
        else:
            vars_map['error_file'].set('未生成')

        if self.report_generated:
            vars_map['report_file'].set('已生成')
        else:
            vars_map['report_file'].set('未生成')

        if self.reminder_generated:
            vars_map['reminder_file'].set('已生成')
        else:
            vars_map['reminder_file'].set('未生成')

        if self.archive_generated:
            vars_map['archive_file'].set('已生成')
        else:
            vars_map['archive_file'].set('未生成')

    def _select_input_folder(self):
        folder = filedialog.askdirectory(title="选择数据文件夹")
        if folder:
            self.input_folder.set(folder)
            self._scan_files()

    def _select_output_folder(self):
        folder = filedialog.askdirectory(title="选择输出文件夹")
        if folder:
            self.output_folder.set(folder)
            self._log(f"输出文件夹: {folder}")

    def _scan_files(self):
        folder = self.input_folder.get()
        if not folder or not os.path.exists(folder):
            return

        current_category = self._get_category()
        all_categories = ['慢病', '孕产妇', '儿童', '老年人']

        self.file_list = []
        for filename in os.listdir(folder):
            if filename.endswith(('.xlsx', '.xls', '.csv')) and not filename.startswith('~$'):
                matches_current = current_category in filename
                matches_other = any(cat in filename for cat in all_categories if cat != current_category)
                if matches_current or not matches_other:
                    self.file_list.append(os.path.join(folder, filename))

        filtered_count = len(self.file_list)
        total_count = len([f for f in os.listdir(folder) if f.endswith(('.xlsx', '.xls', '.csv')) and not f.startswith('~$')])

        self._log(f"扫描到 {total_count} 个数据文件，筛选出 {filtered_count} 个与「{current_category}」相关的文件")
        for f in self.file_list:
            self._log(f"  - {os.path.basename(f)}")

        self._update_overview()

    def _log(self, message):
        self.log_text.config(state=tk.NORMAL)
        timestamp = datetime.now().strftime('%H:%M:%S')
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()

    def _get_category(self):
        return self.current_category.get()

    def _get_output_dir(self):
        output_dir = self.output_folder.get()
        if not output_dir:
            output_dir = os.path.join(os.getcwd(), 'output')
            self.output_folder.set(output_dir)
        os.makedirs(output_dir, exist_ok=True)
        return output_dir

    def action_import(self, silent=False):
        if not self.file_list:
            if self.input_folder.get():
                self._scan_files()
            else:
                if not silent:
                    messagebox.showwarning("提示", "请先选择数据文件夹")
                return False

        if not self.file_list:
            if not silent:
                messagebox.showwarning("提示", "数据文件夹中没有找到 Excel 或 CSV 文件")
            return False

        self._reset_import_state()
        self._log("开始导入数据...")
        try:
            merger = DataMerger(self._get_category())
            result = merger.merge_multiple_files(self.file_list)
            self.merged_data = result['合并数据']
            self.current_data = self.merged_data.copy()
            self.file_info_list = result['文件信息']

            self._log(f"成功导入 {result['总记录数']} 条记录")
            for info in result['文件信息']:
                status = f"{info['记录数']} 条" if '错误' not in info else f"错误: {info['错误']}"
                self._log(f"  - {info['文件名']}: {status}")

            if not silent:
                messagebox.showinfo("成功", f"成功导入 {result['总记录数']} 条记录")

            self._update_overview()
            return True
        except Exception as e:
            self._log(f"导入失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"导入失败: {str(e)}")
            return False

    def action_validate(self, silent=False):
        if self.current_data is None:
            if not silent:
                messagebox.showwarning("提示", "请先导入数据")
            return False

        self._log("开始校验字段...")
        try:
            validator = DataValidator(self._get_category())

            missing_fields = validator.get_missing_fields(self.current_data)
            if missing_fields:
                self._log(f"缺少必要字段: {', '.join(missing_fields)}")
                if not silent:
                    messagebox.showwarning("字段缺失", f"数据中缺少以下必要字段:\n{', '.join(missing_fields)}")

            self.validation_result = validator.validate_dataframe(self.current_data)

            self._log(f"校验完成: 共 {self.validation_result['总记录数']} 条记录")
            self._log(f"  有效记录: {self.validation_result['有效记录数']} 条")
            self._log(f"  无效记录: {self.validation_result['无效记录数']} 条")
            self._log(f"  有效率: {self.validation_result['有效率']}")

            if self.validation_result['错误汇总']:
                self._log("  错误类型统计:")
                for err_type, count in sorted(self.validation_result['错误汇总'].items(),
                                              key=lambda x: x[1], reverse=True):
                    self._log(f"    - {err_type}: {count} 条")

            output_dir = self._get_output_dir()
            error_path = os.path.join(output_dir, f"{self._get_category()}_异常明细.xlsx")
            if self.validation_result['错误明细']:
                error_df = pd.DataFrame(self.validation_result['错误明细'])
                error_df.to_excel(error_path, index=False)
                self._log(f"异常明细已保存: {error_path}")

            if not silent:
                messagebox.showinfo("校验完成",
                                    f"有效率: {self.validation_result['有效率']}\n"
                                    f"有效记录: {self.validation_result['有效记录数']} 条\n"
                                    f"无效记录: {self.validation_result['无效记录数']} 条")

            self._update_overview()
            return True
        except Exception as e:
            self._log(f"校验失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"校验失败: {str(e)}")
            return False

    def action_merge(self, silent=False):
        if self.current_data is None:
            if not silent:
                messagebox.showwarning("提示", "请先导入数据")
            return False

        self._log("开始合并去重...")
        try:
            merger = DataMerger(self._get_category())
            result = merger.remove_duplicates(self.current_data)
            self.deduped_data = result['去重后数据']
            self.current_data = self.deduped_data.copy()

            self._log(f"去重完成: 去除 {result['重复记录数']} 条重复记录")
            self._log(f"  去重前: {result['去重后记录数'] + result['重复记录数']} 条")
            self._log(f"  去重后: {result['去重后记录数']} 条")

            if result['重复明细']:
                self._log(f"  发现 {len(result['重复明细'])} 组重复记录")

            person_result = merger.merge_by_person(self.deduped_data)
            self._log(f"  管理人数: {person_result['人数']} 人")
            self._log(f"  总随访次数: {person_result['总随访次数']} 次")

            if not silent:
                messagebox.showinfo("合并完成",
                                    f"去除重复记录: {result['重复记录数']} 条\n"
                                    f"剩余记录: {result['去重后记录数']} 条\n"
                                    f"管理人数: {person_result['人数']} 人")

            self._update_overview()
            return True
        except Exception as e:
            self._log(f"合并失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"合并失败: {str(e)}")
            return False

    def action_remind(self, silent=False):
        if self.current_data is None:
            if not silent:
                messagebox.showwarning("提示", "请先导入数据")
            return False

        self._log("开始生成提醒...")
        try:
            reminder = ReminderGenerator(self._get_category())

            self.overdue_result = reminder.find_overdue_persons(self.current_data)
            if not self.overdue_result['超期人员'].empty:
                self.overdue_result['超期人员'] = reminder.add_overdue_level(
                    self.overdue_result['超期人员']
                )

            self.next_month_result = reminder.generate_next_month_reminder(self.current_data)
            self.reminder_generated = True

            self._log(f"超期未访统计:")
            self._log(f"  管理总人数: {self.overdue_result['总人数']} 人")
            self._log(f"  超期人数: {self.overdue_result['超期人数']} 人")
            self._log(f"  超期率: {self.overdue_result['超期率']}")
            self._log(f"  参考日期: {self.overdue_result['参考日期']}")
            self._log(f"  随访间隔: {self.overdue_result['随访间隔天数']} 天")

            self._log(f"下月待随访: {self.next_month_result['待随访人数']} 人")

            village_result = reminder.generate_village_reminder(self.current_data)
            self._log(f"涉及村居: {village_result['村居数量']} 个")

            if not silent:
                messagebox.showinfo("提醒生成完成",
                                    f"超期人数: {self.overdue_result['超期人数']} 人\n"
                                    f"超期率: {self.overdue_result['超期率']}\n"
                                    f"下月待随访: {self.next_month_result['待随访人数']} 人")

            self._update_overview()
            return True
        except Exception as e:
            self._log(f"生成提醒失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"生成提醒失败: {str(e)}")
            return False

    def action_export(self, silent=False):
        if self.current_data is None:
            if not silent:
                messagebox.showwarning("提示", "请先导入数据")
            return False

        output_dir = self._get_output_dir()
        self._log("开始分组导出...")
        try:
            exporter = DataExporter(self._get_category())

            village_dir = os.path.join(output_dir, '村居分组')
            result = exporter.export_by_village(self.current_data, village_dir)
            self.export_result = result

            if result['成功']:
                self._log(f"成功导出 {len(result['文件列表'])} 个村居文件")
                for f in result['文件列表']:
                    self._log(f"  - {f['文件名']}: {f['记录数']} 条")
            else:
                self._log(f"导出失败: {result['消息']}")

            if self.overdue_result is not None and not self.overdue_result['超期人员'].empty:
                overdue_dir = os.path.join(output_dir, '超期待访')
                overdue_result = exporter.export_overdue_by_village(
                    self.overdue_result['超期人员'], overdue_dir
                )
                if overdue_result['成功']:
                    self._log(f"导出超期待访清单: {len(overdue_result['文件列表'])} 个村居")

            if not silent:
                messagebox.showinfo("导出完成", f"已导出到: {output_dir}")

            self._update_overview()
            return result['成功']
        except Exception as e:
            self._log(f"导出失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"导出失败: {str(e)}")
            return False

    def action_report(self, silent=False):
        if self.current_data is None:
            if not silent:
                messagebox.showwarning("提示", "请先导入数据")
            return False

        output_dir = self._get_output_dir()
        self._log("开始生成汇总报告...")
        try:
            reporter = ReportGenerator(self._get_category())
            reminder = ReminderGenerator(self._get_category())

            report_path = os.path.join(output_dir, f"{self._get_category()}_汇总报告.xlsx")
            merge_result = {
                '总记录数': len(self.merged_data) if self.merged_data is not None else len(self.current_data),
                '重复记录数': len(self.merged_data) - len(self.deduped_data)
                if (self.merged_data is not None and self.deduped_data is not None) else 0,
                '去重后记录数': len(self.deduped_data) if self.deduped_data is not None else len(self.current_data),
            }

            if self.validation_result is None:
                validator = DataValidator(self._get_category())
                self.validation_result = validator.validate_dataframe(self.current_data)

            if self.overdue_result is None:
                self.overdue_result = reminder.find_overdue_persons(self.current_data)
                if not self.overdue_result['超期人员'].empty:
                    self.overdue_result['超期人员'] = reminder.add_overdue_level(
                        self.overdue_result['超期人员']
                    )

            if self.next_month_result is None:
                self.next_month_result = reminder.generate_next_month_reminder(self.current_data)
                self.reminder_generated = True

            result = reporter.generate_summary_report(
                self.current_data,
                self.validation_result,
                self.overdue_result,
                merge_result,
                report_path,
                self.file_info_list,
                self.next_month_result.get('下月待随访人员') if self.next_month_result else None
            )

            if result['成功']:
                self._log(f"汇总报告已生成: {report_path}")
                self.report_generated = True
            else:
                self._log(f"生成报告失败: {result['消息']}")

            reminder_path = os.path.join(output_dir, f"{self._get_category()}_下月提醒.xlsx")
            rem_result = reporter.generate_next_month_reminder_file(
                self.next_month_result['下月待随访人员'],
                reminder_path
            )
            if rem_result['成功']:
                if rem_result['待随访人数'] == 0:
                    self._log(f"下月提醒文件已生成（无待随访人员）: {reminder_path}")
                else:
                    self._log(f"下月提醒文件已生成: {reminder_path}")
                self.reminder_generated = True

            if not silent:
                messagebox.showinfo("报告生成完成",
                                    f"汇总报告已保存到:\n{report_path}\n\n"
                                    f"下月提醒已保存到:\n{reminder_path}")

            self._update_overview()
            return result['成功']
        except Exception as e:
            self._log(f"生成报告失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"生成报告失败: {str(e)}")
            return False

    def action_archive(self, silent=False):
        output_dir = self._get_output_dir()
        category = self._get_category()

        if not self.report_generated:
            if not silent:
                messagebox.showwarning("提示", "请先生成汇总报告")
            return False

        self._log("开始生成月底归档包...")
        try:
            now = datetime.now()
            archive_folder_name = f"{category}_{now.strftime('%Y%m')}月_归档"
            archive_path = os.path.join(output_dir, archive_folder_name)

            if os.path.exists(archive_path):
                shutil.rmtree(archive_path)
            os.makedirs(archive_path, exist_ok=True)

            file_list = []
            success_count = 0
            fail_count = 0

            report_file = f"{category}_汇总报告.xlsx"
            report_src = os.path.join(output_dir, report_file)
            if self.report_generated and os.path.exists(report_src):
                shutil.copy2(report_src, os.path.join(archive_path, report_file))
                file_list.append({'文件名': report_file, '类型': '汇总报告', '状态': '成功'})
                success_count += 1
            else:
                fail_count += 1

            reminder_file = f"{category}_下月提醒.xlsx"
            reminder_src = os.path.join(output_dir, reminder_file)
            if self.reminder_generated and os.path.exists(reminder_src):
                shutil.copy2(reminder_src, os.path.join(archive_path, reminder_file))
                file_list.append({'文件名': reminder_file, '类型': '下月提醒', '状态': '成功'})
                success_count += 1
            else:
                fail_count += 1

            error_file = f"{category}_异常明细.xlsx"
            error_src = os.path.join(output_dir, error_file)
            if self.validation_result and self.validation_result.get('错误明细') and os.path.exists(error_src):
                shutil.copy2(error_src, os.path.join(archive_path, error_file))
                file_list.append({'文件名': error_file, '类型': '异常明细', '状态': '成功'})
                success_count += 1

            if self.export_result and self.export_result.get('成功') and self.export_result.get('文件列表'):
                village_dir_dst = os.path.join(archive_path, '村居分组')
                os.makedirs(village_dir_dst, exist_ok=True)
                village_count = 0
                for vf_info in self.export_result['文件列表']:
                    vf_name = vf_info['文件名']
                    vf_src = os.path.join(output_dir, '村居分组', vf_name)
                    if os.path.exists(vf_src):
                        shutil.copy2(vf_src, os.path.join(village_dir_dst, vf_name))
                        file_list.append({'文件名': f'村居分组/{vf_name}', '类型': '村居文件', '状态': '成功'})
                        village_count += 1
                success_count += village_count

            summary_path = os.path.join(archive_path, '核对摘要.txt')
            with open(summary_path, 'w', encoding='utf-8') as f:
                f.write(f"{'=' * 55}\n")
                f.write(f"           {category} 随访资料核对摘要\n")
                f.write(f"{'=' * 55}\n\n")
                f.write(f"归档时间: {now.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"归档月份: {now.strftime('%Y年%m月')}\n")
                f.write(f"分类: {category}\n\n")

                f.write(f"【数据导入】\n")
                f.write(f"  导入文件数: {len(self.file_info_list)} 个\n")
                if self.current_data is not None:
                    f.write(f"  记录总数: {len(self.current_data)} 条\n")
                else:
                    f.write(f"  记录总数: 0 条\n")

                if self.file_info_list:
                    for info in self.file_info_list:
                        if '错误' in info:
                            f.write(f"    - {info['文件名']}: 读取失败 - {info['错误']}\n")
                        else:
                            f.write(f"    - {info['文件名']}: {info['记录数']} 条\n")
                f.write("\n")

                f.write(f"【数据质量】\n")
                if self.validation_result:
                    f.write(f"  数据有效率: {self.validation_result.get('有效率', '0%')}\n")
                    f.write(f"  有效记录数: {self.validation_result.get('有效记录数', 0)} 条\n")
                    f.write(f"  无效记录数: {self.validation_result.get('无效记录数', 0)} 条\n")
                    error_types = self.validation_result.get('错误汇总', {})
                    if error_types:
                        f.write(f"  异常类型:\n")
                        for err_type, count in sorted(error_types.items(), key=lambda x: x[1], reverse=True):
                            f.write(f"    - {err_type}: {count} 条\n")
                else:
                    f.write(f"  未校验\n")
                f.write("\n")

                f.write(f"【随访统计】\n")
                if self.overdue_result:
                    f.write(f"  管理总人数: {self.overdue_result.get('总人数', 0)} 人\n")
                    f.write(f"  超期未访: {self.overdue_result.get('超期人数', 0)} 人\n")
                    f.write(f"  超期率: {self.overdue_result.get('超期率', '0%')}\n")
                else:
                    f.write(f"  未统计\n")
                f.write("\n")

                f.write(f"【下月随访】\n")
                if self.next_month_result:
                    f.write(f"  待随访人数: {self.next_month_result.get('待随访人数', 0)} 人\n")
                    village_stats = self.next_month_result.get('村居统计', pd.DataFrame())
                    if not village_stats.empty:
                        f.write(f"  村居分布:\n")
                        for _, row in village_stats.head(10).iterrows():
                            f.write(f"    - {row['村居']}: {row['人数']} 人\n")
                        if len(village_stats) > 10:
                            f.write(f"    ... 共 {len(village_stats)} 个村居\n")
                else:
                    f.write(f"  未生成\n")
                f.write("\n")

                f.write(f"【输出文件】\n")
                f.write(f"  汇总报告: {'已生成' if self.report_generated else '未生成'}\n")
                f.write(f"  下月提醒: {'已生成' if self.reminder_generated else '未生成'}\n")
                if self.validation_result and self.validation_result.get('错误明细'):
                    error_count = len(self.validation_result['错误明细'])
                    f.write(f"  异常明细: 已生成 ({error_count} 条)\n")
                else:
                    f.write(f"  异常明细: 未生成\n")
                if self.export_result and self.export_result.get('成功'):
                    f.write(f"  村居分组: 已生成 ({len(self.export_result['文件列表'])} 个文件)\n")
                else:
                    f.write(f"  村居分组: 未生成\n")
                f.write("\n")

                f.write(f"【归档信息】\n")
                f.write(f"  归档文件数: {success_count} 个\n")
                if fail_count > 0:
                    f.write(f"  缺失文件数: {fail_count} 个\n")
                f.write(f"  归档路径: {archive_path}\n\n")

                f.write(f"{'=' * 55}\n")
                f.write(f"  核对人: ______________    核对日期: ______________\n")
                f.write(f"{'=' * 55}\n")

            file_list.insert(0, {'文件名': '核对摘要.txt', '类型': '摘要文件', '状态': '成功'})
            success_count += 1

            manifest_path = os.path.join(archive_path, '文件清单.txt')
            with open(manifest_path, 'w', encoding='utf-8') as f:
                f.write(f"{category} 随访资料归档清单\n")
                f.write(f"归档时间: {now.strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"归档月份: {now.strftime('%Y年%m月')}\n")
                f.write("=" * 55 + "\n\n")
                f.write(f"总文件数: {success_count} 个\n")
                if fail_count > 0:
                    f.write(f"缺失文件: {fail_count} 个\n")
                f.write("\n")
                f.write("-" * 55 + "\n")
                f.write(f"{'序号':<5}{'文件名':<35}{'类型':<12}{'状态':<10}\n")
                f.write("-" * 55 + "\n")
                for idx, item in enumerate(file_list, 1):
                    f.write(f"{idx:<5}{item['文件名']:<35}{item['类型']:<12}{item['状态']:<10}\n")
                f.write("\n")
                f.write("=" * 55 + "\n")
                f.write(f"归档路径: {archive_path}\n")

            self.archive_generated = True
            self.archive_path = archive_path
            self.archive_file_count = success_count
            self.archive_fail_count = fail_count

            self._log(f"归档包已生成: {archive_path}")
            self._log(f"  包含文件: {success_count} 个")
            if fail_count > 0:
                self._log(f"  缺失文件: {fail_count} 个")

            if not silent:
                msg = f"归档包已生成:\n{archive_path}\n\n"
                msg += f"成功文件: {success_count} 个\n"
                if fail_count > 0:
                    msg += f"缺失文件: {fail_count} 个\n"
                messagebox.showinfo("归档完成", msg)

            self._update_overview()
            return True
        except Exception as e:
            self._log(f"生成归档包失败: {str(e)}")
            if not silent:
                messagebox.showerror("错误", f"生成归档包失败: {str(e)}")
            return False

    def action_all(self):
        self._log("=" * 50)
        self._log("开始执行全部操作...")
        self._log("=" * 50)

        steps = [
            ('导入名单', 'import'),
            ('校验字段', 'validate'),
            ('合并随访', 'merge'),
            ('生成提醒', 'remind'),
            ('分组导出', 'export'),
            ('汇总报告', 'report'),
            ('生成归档包', 'archive'),
        ]

        step_results = {}
        failed_step = None
        failed_message = ''

        for step_name, step_key in steps:
            success = False
            try:
                if step_key == 'import':
                    success = self.action_import(silent=True)
                elif step_key == 'validate':
                    success = self.action_validate(silent=True)
                elif step_key == 'merge':
                    success = self.action_merge(silent=True)
                elif step_key == 'remind':
                    success = self.action_remind(silent=True)
                elif step_key == 'export':
                    success = self.action_export(silent=True)
                elif step_key == 'report':
                    success = self.action_report(silent=True)
                elif step_key == 'archive':
                    success = self.action_archive(silent=True)
            except Exception as e:
                failed_message = str(e)
                self._log(f"{step_name}异常: {str(e)}")

            step_results[step_key] = success
            if not success and step_key != 'archive':
                failed_step = step_name
                if not failed_message:
                    failed_message = f'{step_name}执行失败'
                self._log(f"步骤失败: {step_name}，终止后续操作")
                break
            elif not success and step_key == 'archive':
                self._log(f"归档包生成失败，不影响其他步骤")

        self._log("=" * 50)
        output_dir = self._get_output_dir()

        output_files = []

        if self.report_generated:
            output_files.append({'名称': '汇总报告', '文件': f'{self._get_category()}_汇总报告.xlsx', '状态': '已生成', '类型': '主要文件'})
        else:
            output_files.append({'名称': '汇总报告', '文件': f'{self._get_category()}_汇总报告.xlsx', '状态': '未生成', '类型': '主要文件'})

        if self.reminder_generated:
            output_files.append({'名称': '下月提醒', '文件': f'{self._get_category()}_下月提醒.xlsx', '状态': '已生成', '类型': '主要文件'})
        else:
            output_files.append({'名称': '下月提醒', '文件': f'{self._get_category()}_下月提醒.xlsx', '状态': '未生成', '类型': '主要文件'})

        if self.validation_result and self.validation_result.get('错误明细'):
            error_count = len(self.validation_result['错误明细'])
            output_files.append({'名称': '异常明细', '文件': f'{self._get_category()}_异常明细.xlsx', '状态': f'已生成 ({error_count}条)', '类型': '主要文件'})
        else:
            output_files.append({'名称': '异常明细', '文件': f'{self._get_category()}_异常明细.xlsx', '状态': '未生成', '类型': '主要文件'})

        if self.export_result and self.export_result.get('成功'):
            village_count = len(self.export_result['文件列表'])
            output_files.append({'名称': '村居分组', '文件': f'{village_count} 个村居文件', '状态': '已生成', '类型': '分组文件'})
        else:
            output_files.append({'名称': '村居分组', '文件': '-', '状态': '未生成', '类型': '分组文件'})

        if self.archive_generated:
            archive_count = getattr(self, 'archive_file_count', 0)
            output_files.append({'名称': '归档包', '文件': f'{self._get_category()}_{datetime.now().strftime("%Y%m")}月_归档', '状态': f'已生成 ({archive_count}个文件)', '类型': '归档'})
        else:
            output_files.append({'名称': '归档包', '文件': '-', '状态': '未生成', '类型': '归档'})

        main_files = [f for f in output_files if f['类型'] == '主要文件' and f['状态'].startswith('已生成')]
        group_files = [f for f in output_files if f['类型'] == '分组文件' and f['状态'].startswith('已生成')]

        if failed_step:
            self._log(f"执行失败，失败步骤: {failed_step}")
            self._log(f"输出目录: {output_dir}")
            self._log("=" * 50)

            msg_lines = [
                f"批量执行失败！",
                f"",
                f"失败步骤: {failed_step}",
                f"错误信息: {failed_message}",
                f"",
                f"各步骤状态:",
            ]
            for step_name, step_key in steps:
                if step_results.get(step_key):
                    status = '✓ 成功'
                elif step_key in step_results:
                    status = '✗ 失败'
                else:
                    status = '○ 未执行'
                msg_lines.append(f"  {status} - {step_name}")

            msg_lines.append("")
            msg_lines.append("输出文件状态:")
            for f in output_files:
                icon = '✓' if f['状态'].startswith('已生成') else '○'
                msg_lines.append(f"  {icon} {f['名称']}: {f['状态']}")

            if self.archive_generated and self.archive_path:
                msg_lines.append("")
                msg_lines.append(f"归档包位置: {self.archive_path}")

            msg_lines.append("")
            msg_lines.append(f"输出目录: {output_dir}")

            full_msg = "\n".join(msg_lines)
            messagebox.showerror("执行失败", full_msg)
        else:
            self._log("全部操作执行完成！")
            self._log(f"输出目录: {output_dir}")
            if self.archive_generated:
                self._log(f"归档包: {self.archive_path}")
            self._log("=" * 50)

            summary_lines = []
            if self.validation_result:
                summary_lines.append(f"数据有效率: {self.validation_result['有效率']}")
            if self.overdue_result:
                summary_lines.append(f"超期未访: {self.overdue_result['超期人数']} 人 ({self.overdue_result['超期率']})")
            if self.next_month_result:
                summary_lines.append(f"下月待随访: {self.next_month_result['待随访人数']} 人")

            msg_lines = ["全部操作执行完成！", ""]
            if summary_lines:
                msg_lines.extend(summary_lines)
                msg_lines.append("")

            msg_lines.append("输出文件:")
            for f in output_files:
                icon = '✓' if f['状态'].startswith('已生成') else '○'
                msg_lines.append(f"  {icon} {f['名称']}: {f['状态']}")

            msg_lines.append("")
            if self.archive_generated and self.archive_path:
                msg_lines.append(f"归档包位置: {self.archive_path}")
                if hasattr(self, 'archive_file_count'):
                    msg_lines.append(f"归档文件数: {self.archive_file_count} 个")
            msg_lines.append(f"输出目录: {output_dir}")

            full_msg = "\n".join(msg_lines)
            messagebox.showinfo("完成", full_msg)

    def run(self):
        self.root.mainloop()
