"""
底稿整理助手 v1.1.0 - 报告格式调整工具
纯本地运行，无需联网
支持：模板导入/管理/应用，报告格式一键调整
"""

import os
import sys
import json
import logging
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from datetime import datetime

__version__ = "1.1.0"
__app_name__ = "底稿整理助手"

# 配置日志
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(name)s: %(message)s',
    handlers=[
        logging.FileHandler(
            os.path.join(os.path.expanduser('~'), '.底稿整理助手', 'app.log'),
            encoding='utf-8'
        )
    ]
)
logger = logging.getLogger(__name__)

# 确保能找到同级模块（兼容 PyInstaller 打包）
if getattr(sys, 'frozen', False):
    # PyInstaller 打包后的运行目录
    script_dir = os.path.dirname(sys.executable)
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from docx_handler import (
    extract_styles_from_docx, apply_template_to_docx,
    init_templates_dir, save_template, load_template, list_templates,
    delete_template, backup_file, get_default_template, detect_heading_paragraphs,
    CN_FONTS, EN_FONTS, CN_FONT_SIZES, LINE_SPACING_OPTIONS,
    PAGE_NUM_FORMATS, DATE_FORMATS, TEMPLATE_CATEGORIES,
    cn_to_pt, pt_to_cn_size
)


class FormatTool:
    """底稿整理助手 - 格式调整主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title(f"{__app_name__} v{__version__} - 报告格式调整")
        self.root.geometry("1100x750")
        self.root.minsize(900, 650)

        # 当前状态
        self.current_file = None
        self.current_file_info = None
        self.current_template = get_default_template()
        self.template_list = []

        # 初始化模板目录
        user_data_dir = self._get_user_data_dir()
        init_templates_dir(user_data_dir)

        # 设置样式
        self._setup_styles()

        # 构建界面
        self._build_ui()

        # 加载模板列表
        self._refresh_template_list()

        # 绑定关闭事件
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        logger.info(f"{__app_name__} v{__version__} 启动成功")

    def _get_user_data_dir(self):
        """获取用户数据目录（模板存储位置）"""
        if getattr(sys, 'frozen', False):
            # PyInstaller 打包后使用 exe 所在目录
            base = os.path.dirname(sys.executable)
        elif sys.platform == 'win32':
            base = os.environ.get('APPDATA', os.path.expanduser('~'))
        else:
            base = os.path.expanduser('~')
        data_dir = os.path.join(base, '.底稿整理助手')
        if not os.path.exists(data_dir):
            os.makedirs(data_dir)
        return data_dir

    def _setup_styles(self):
        """设置 ttk 样式"""
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Title.TLabel', font=('微软雅黑', 12, 'bold'))
        style.configure('Status.TLabel', font=('微软雅黑', 9))
        style.configure('Action.TButton', font=('微软雅黑', 10))
        style.configure('Apply.TButton', font=('微软雅黑', 11, 'bold'), foreground='#ffffff', background='#2d5a27')
        style.map('Apply.TButton', background=[('active', '#3d7a37')])

    def _build_ui(self):
        """构建主界面"""
        # ---- 顶部工具栏 ----
        top_frame = ttk.Frame(self.root, padding=(10, 5, 10, 5))
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text=f"📄 {__app_name__} v{__version__}", style='Title.TLabel').pack(side=tk.LEFT)
        ttk.Label(top_frame, text="纯本地运行 · 无需联网", font=('微软雅黑', 8), foreground='gray').pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        # ---- 文件选择区 ----
        file_frame = ttk.LabelFrame(self.root, text="  报告文件  ", padding=(10, 8))
        file_frame.pack(fill=tk.X, padx=10, pady=(5, 0))

        file_row = ttk.Frame(file_frame)
        file_row.pack(fill=tk.X)

        self.file_path_var = tk.StringVar(value="未选择文件")
        ttk.Entry(file_row, textvariable=self.file_path_var, state='readonly').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        ttk.Button(file_row, text="📂 打开报告", command=self._open_file, width=15).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(file_row, text="📋 文件信息", command=self._show_file_info, width=12).pack(side=tk.RIGHT)

        # ---- 主内容区 ----
        main_frame = ttk.Frame(self.root, padding=(10, 5, 10, 5))
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧 - 模板管理
        left_frame = ttk.LabelFrame(main_frame, text="  模板管理  ", padding=(8, 8), width=280)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.pack_propagate(False)

        self._build_template_panel(left_frame)

        # 右侧 - 格式设置
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        self._build_format_tabs()

        # ---- 底部操作栏 ----
        bottom_frame = ttk.Frame(self.root, padding=(10, 5, 10, 10))
        bottom_frame.pack(fill=tk.X)

        self.backup_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom_frame, text="应用前自动备份原文件", variable=self.backup_var).pack(side=tk.LEFT)

        ttk.Button(
            bottom_frame, text="🚀 一键应用模板", style='Apply.TButton',
            command=self._apply_template, width=20
        ).pack(side=tk.RIGHT)

        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10)

        # ---- 状态栏 ----
        status_frame = ttk.Frame(self.root, padding=(10, 3, 10, 3))
        status_frame.pack(fill=tk.X)

        self.status_var = tk.StringVar(value="就绪")
        ttk.Label(status_frame, textvariable=self.status_var, style='Status.TLabel').pack(side=tk.LEFT)

        self.file_info_var = tk.StringVar(value="")
        ttk.Label(status_frame, textvariable=self.file_info_var, style='Status.TLabel').pack(side=tk.RIGHT)

    # ========== 模板管理面板 ==========

    def _build_template_panel(self, parent):
        """构建模板管理面板"""
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Button(btn_frame, text="📥 导入模板", command=self._import_template, width=15).pack(fill=tk.X, pady=1)
        ttk.Button(btn_frame, text="💾 保存为模板", command=self._save_as_template, width=15).pack(fill=tk.X, pady=1)
        ttk.Button(btn_frame, text="🔄 恢复默认", command=self._reset_to_default, width=15).pack(fill=tk.X, pady=1)
        ttk.Button(btn_frame, text="🗑️ 删除模板", command=self._delete_template, width=15).pack(fill=tk.X, pady=1)

        tree_frame = ttk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(tree_frame, text="已保存的模板：", font=('微软雅黑', 9, 'bold')).pack(anchor=tk.W)

        self.template_tree = ttk.Treeview(tree_frame, columns=("source",), show="tree", height=12)
        self.template_tree.pack(fill=tk.BOTH, expand=True)
        self.template_tree.column("#0", width=200, minwidth=150)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.template_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.template_tree.configure(yscrollcommand=scrollbar.set)

        self.template_tree.bind("<<TreeviewSelect>>", self._on_template_select)

    def _refresh_template_list(self):
        """刷新模板列表"""
        self.template_tree.delete(*self.template_tree.get_children())
        self.template_list = list_templates()

        categories = {}
        for tpl in self.template_list:
            cat = tpl["category"]
            if cat not in categories:
                categories[cat] = []
            categories[cat].append(tpl)

        for cat in TEMPLATE_CATEGORIES:
            if cat in categories and categories[cat]:
                cat_id = self.template_tree.insert("", "end", text=f"📁 {cat}", open=True)
                for tpl in categories[cat]:
                    self.template_tree.insert(
                        cat_id, "end", text=f"  📄 {tpl['name']}",
                        values=(tpl["source"],),
                        tags=(tpl["filepath"],)
                    )

    def _on_template_select(self, event):
        """选择模板事件"""
        selection = self.template_tree.selection()
        if not selection:
            return

        item = selection[0]
        parent = self.template_tree.parent(item)
        if not parent:
            return

        filepath = self.template_tree.item(item, "tags")[0] if self.template_tree.item(item, "tags") else None
        if filepath:
            template = load_template(filepath)
            if template:
                self.current_template = template
                self._load_template_to_ui()
                self.status_var.set(f"已加载模板：{self.template_tree.item(item, 'text').replace('  📄 ', '')}")
            else:
                messagebox.showerror("错误", "模板文件损坏，无法加载")

    # ========== 格式设置分页 ==========

    def _build_format_tabs(self):
        """构建所有格式设置分页"""
        self._build_cover_tab()
        self._build_headings_tab()
        self._build_body_tab()
        self._build_table_tab()
        self._build_image_tab()
        self._build_header_footer_tab()
        self._build_toc_signature_tab()

    def _add_combo_row(self, parent, label, options, default=None, row=0, column_span=1):
        """添加一行下拉选择"""
        ttk.Label(parent, text=label, width=10, anchor=tk.E).grid(row=row, column=0, sticky=tk.W, padx=(0, 5), pady=2)
        var = tk.StringVar(value=default if default else (options[0] if options else ""))
        combo = ttk.Combobox(parent, textvariable=var, values=options, state="readonly", width=column_span * 12 - 5)
        combo.grid(row=row, column=1, columnspan=column_span, sticky=tk.W, pady=2)
        return var

    def _add_check_row(self, parent, text, default=True, row=0, column_span=1):
        """添加一行复选框"""
        var = tk.BooleanVar(value=default)
        cb = ttk.Checkbutton(parent, text=text, variable=var)
        cb.grid(row=row, column=0, columnspan=1 + column_span, sticky=tk.W, pady=2)
        return var

    def _build_cover_tab(self):
        """封面设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="📄 封面")

        self.cover_auto = self._add_check_row(frame, "☑ 自动生成封面（若报告没有封面页）", default=False, row=0)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=5)
        ttk.Label(frame, text="报告标题", font=('微软雅黑', 9, 'bold')).grid(row=2, column=0, columnspan=3, sticky=tk.W)

        self.cover_title_font = self._add_combo_row(frame, "字体", CN_FONTS, "宋体", row=3)
        self.cover_title_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "一号", row=4)
        self.cover_title_bold = self._add_check_row(frame, "加粗", True, row=5)
        self.cover_title_align = self._add_combo_row(frame, "对齐", ["居中", "左对齐", "右对齐"], "居中", row=6)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=5)
        ttk.Label(frame, text="单位名称", font=('微软雅黑', 9, 'bold')).grid(row=8, column=0, columnspan=3, sticky=tk.W)

        self.cover_company_font = self._add_combo_row(frame, "字体", CN_FONTS, "宋体", row=9)
        self.cover_company_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "三号", row=10)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=11, column=0, columnspan=3, sticky=tk.EW, pady=5)
        ttk.Label(frame, text="日期", font=('微软雅黑', 9, 'bold')).grid(row=12, column=0, columnspan=3, sticky=tk.W)

        self.cover_date_font = self._add_combo_row(frame, "字体", CN_FONTS, "宋体", row=13)
        self.cover_date_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "四号", row=14)

    def _build_headings_tab(self):
        """标题设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="🏷️ 标题")

        ttk.Label(frame, text="一级标题", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.h1_font = self._add_combo_row(frame, "字体", CN_FONTS, "黑体", row=2)
        self.h1_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "三号", row=3)
        self.h1_bold = self._add_check_row(frame, "加粗", True, row=4)
        self.h1_align = self._add_combo_row(frame, "对齐", ["居中", "左对齐", "右对齐"], "居中", row=5)
        self.h1_space_before = self._add_combo_row(frame, "段前距", ["6", "8", "12", "18", "24"], "12", row=6)
        self.h1_space_after = self._add_combo_row(frame, "段后距", ["3", "6", "8", "12"], "6", row=7)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=8, column=0, columnspan=3, sticky=tk.EW, pady=(15, 3))
        ttk.Label(frame, text="二级标题", font=('微软雅黑', 10, 'bold')).grid(row=9, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=10, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.h2_font = self._add_combo_row(frame, "字体", CN_FONTS, "黑体", row=11)
        self.h2_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "四号", row=12)
        self.h2_bold = self._add_check_row(frame, "加粗", True, row=13)
        self.h2_align = self._add_combo_row(frame, "对齐", ["左对齐", "居中", "右对齐", "两端对齐"], "左对齐", row=14)
        self.h2_space_before = self._add_combo_row(frame, "段前距", ["3", "6", "8", "12", "18"], "8", row=15)
        self.h2_space_after = self._add_combo_row(frame, "段后距", ["3", "6", "8"], "4", row=16)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=17, column=0, columnspan=3, sticky=tk.EW, pady=(15, 3))
        ttk.Label(frame, text="三级标题", font=('微软雅黑', 10, 'bold')).grid(row=18, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=19, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.h3_font = self._add_combo_row(frame, "字体", CN_FONTS, "仿宋", row=20)
        self.h3_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "小四", row=21)
        self.h3_bold = self._add_check_row(frame, "加粗", True, row=22)
        self.h3_align = self._add_combo_row(frame, "对齐", ["左对齐", "居中", "右对齐", "两端对齐"], "左对齐", row=23)
        self.h3_space_before = self._add_combo_row(frame, "段前距", ["3", "6", "8", "12"], "6", row=24)
        self.h3_space_after = self._add_combo_row(frame, "段后距", ["3", "6", "8"], "3", row=25)

    def _build_body_tab(self):
        """正文设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="📝 正文")

        self.body_font = self._add_combo_row(frame, "字体", CN_FONTS, "仿宋", row=0)
        self.body_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "小四", row=1)
        self.body_bold = self._add_check_row(frame, "加粗", False, row=2)
        self.body_align = self._add_combo_row(frame, "对齐", ["两端对齐", "左对齐", "居中", "右对齐"], "两端对齐", row=3)
        self.body_line_spacing = self._add_combo_row(frame, "行距", LINE_SPACING_OPTIONS, "1.5", row=4)
        self.body_indent = self._add_combo_row(frame, "首行缩进(字符)", ["0", "1", "2", "3", "4"], "2", row=5)

    def _build_table_tab(self):
        """表格设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="📊 表格")

        self.table_header_bold = self._add_check_row(frame, "表头文字加粗", True, row=0)
        self.table_header_shading = self._add_check_row(frame, "表头添加灰色底纹", True, row=1)
        self.table_content_align = self._add_combo_row(frame, "单元格对齐", ["居中", "左对齐", "右对齐"], "居中", row=2, column_span=2)
        self.table_auto_fit = self._add_check_row(frame, "自动调整表格宽度", True, row=3)

    def _build_image_tab(self):
        """图片设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="🖼️ 图片")

        self.image_center = self._add_check_row(frame, "图片居中对齐", True, row=0)
        self.image_max_width = self._add_combo_row(frame, "最大宽度(cm)", ["10", "12", "14", "15", "16", "18", "20"], "15", row=1, column_span=2)

    def _build_header_footer_tab(self):
        """页眉页脚页码设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="📑 页眉/页码")

        ttk.Label(frame, text="页眉设置", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=3)

        ttk.Label(frame, text="页眉内容", width=10, anchor=tk.E).grid(row=2, column=0, sticky=tk.W, padx=(0, 5), pady=2)
        self.header_text_var = tk.StringVar(value="")
        ttk.Entry(frame, textvariable=self.header_text_var, width=30).grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=2)

        self.header_font = self._add_combo_row(frame, "字体", CN_FONTS, "宋体", row=3)
        self.header_size = self._add_combo_row(frame, "字号", list(CN_FONT_SIZES.keys()), "小五", row=4)
        self.header_align = self._add_combo_row(frame, "对齐", ["居中", "左对齐", "右对齐"], "居中", row=5)
        self.header_diff_first = self._add_check_row(frame, "首页不显示页眉", True, row=6)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=7, column=0, columnspan=3, sticky=tk.EW, pady=(15, 3))
        ttk.Label(frame, text="页码设置", font=('微软雅黑', 10, 'bold')).grid(row=8, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=9, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.page_num_format = self._add_combo_row(frame, "页码格式", PAGE_NUM_FORMATS, "第1页 / 共N页", row=10)
        self.page_align = self._add_combo_row(frame, "位置", ["居中", "左对齐", "右对齐"], "居中", row=11)
        self.page_start = self._add_combo_row(frame, "起始页码", ["1", "2", "3", "4", "5"], "1", row=12)
        self.page_show_first = self._add_check_row(frame, "首页显示页码", False, row=13)

    def _build_toc_signature_tab(self):
        """目录和落款设置"""
        frame = ttk.Frame(self.notebook, padding=(15, 10))
        self.notebook.add(frame, text="📋 目录/落款")

        ttk.Label(frame, text="目录设置", font=('微软雅黑', 10, 'bold')).grid(row=0, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.toc_include = self._add_check_row(frame, "插入目录页", False, row=2)
        ttk.Label(frame, text="注意：目录页插入后，需在 Word 中按 Ctrl+A → F9 更新",
                 font=('微软雅黑', 8), foreground='gray').grid(row=3, column=0, columnspan=3, sticky=tk.W, padx=20)

        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=(15, 3))
        ttk.Label(frame, text="落款设置", font=('微软雅黑', 10, 'bold')).grid(row=5, column=0, columnspan=3, sticky=tk.W)
        ttk.Separator(frame, orient=tk.HORIZONTAL).grid(row=6, column=0, columnspan=3, sticky=tk.EW, pady=3)

        self.sig_align_right = self._add_check_row(frame, "落款（盖章/日期）右对齐", True, row=7)
        self.sig_include_date = self._add_check_row(frame, "包含日期", True, row=8)
        self.sig_date_format = self._add_combo_row(frame, "日期格式", DATE_FORMATS, DATE_FORMATS[0], row=9, column_span=2)

    # ========== UI ↔ 模板数据同步 ==========

    def _load_template_to_ui(self):
        """将当前模板数据加载到 UI 各控件"""
        tpl = self.current_template

        cover = tpl.get("cover", {})
        self.cover_auto.set(cover.get("auto_generate", False))
        self.cover_title_font.set(cover.get("title_font", "宋体"))
        self.cover_title_size.set(cover.get("title_size", "一号"))
        self.cover_title_bold.set(cover.get("title_bold", True))
        self.cover_title_align.set(cover.get("title_align", "居中"))
        self.cover_company_font.set(cover.get("company_font", "宋体"))
        self.cover_company_size.set(cover.get("company_size", "三号"))
        self.cover_date_font.set(cover.get("date_font", "宋体"))
        self.cover_date_size.set(cover.get("date_size", "四号"))

        h1 = tpl.get("headings", {}).get("heading1", {})
        self.h1_font.set(h1.get("font", "黑体"))
        self.h1_size.set(h1.get("size", "三号"))
        self.h1_bold.set(h1.get("bold", True))
        self.h1_align.set(h1.get("align", "居中"))
        self.h1_space_before.set(str(h1.get("space_before", "12")))
        self.h1_space_after.set(str(h1.get("space_after", "6")))

        h2 = tpl.get("headings", {}).get("heading2", {})
        self.h2_font.set(h2.get("font", "黑体"))
        self.h2_size.set(h2.get("size", "四号"))
        self.h2_bold.set(h2.get("bold", True))
        self.h2_align.set(h2.get("align", "左对齐"))
        self.h2_space_before.set(str(h2.get("space_before", "8")))
        self.h2_space_after.set(str(h2.get("space_after", "4")))

        h3 = tpl.get("headings", {}).get("heading3", {})
        self.h3_font.set(h3.get("font", "仿宋"))
        self.h3_size.set(h3.get("size", "小四"))
        self.h3_bold.set(h3.get("bold", True))
        self.h3_align.set(h3.get("align", "左对齐"))
        self.h3_space_before.set(str(h3.get("space_before", "6")))
        self.h3_space_after.set(str(h3.get("space_after", "3")))

        body = tpl.get("body", {})
        self.body_font.set(body.get("font", "仿宋"))
        self.body_size.set(body.get("size", "小四"))
        self.body_bold.set(body.get("bold", False))
        self.body_align.set(body.get("align", "两端对齐"))
        ls = body.get("line_spacing", 1.5)
        self.body_line_spacing.set(str(ls) if ls else "1.5")
        self.body_indent.set(str(int(body.get("first_line_indent", 2))))

        tbl = tpl.get("table", {})
        self.table_header_bold.set(tbl.get("header_bold", True))
        self.table_header_shading.set(tbl.get("header_shading", True))
        self.table_content_align.set(tbl.get("content_align", "居中"))
        self.table_auto_fit.set(tbl.get("auto_fit", True))

        img = tpl.get("image", {})
        self.image_center.set(img.get("center_align", True))
        self.image_max_width.set(str(img.get("max_width_cm", 15)))

        header = tpl.get("header", {})
        self.header_text_var.set(header.get("text", ""))
        self.header_font.set(header.get("font", "宋体"))
        self.header_size.set(header.get("size", "小五"))
        self.header_align.set(header.get("align", "居中"))
        self.header_diff_first.set(header.get("different_first_page", True))

        footer = tpl.get("footer", {})
        self.page_num_format.set(footer.get("page_num_format", "第1页 / 共N页"))
        self.page_align.set(footer.get("align", "居中"))
        self.page_start.set(str(footer.get("start_from", 1)))
        self.page_show_first.set(footer.get("show_on_first_page", False))

        toc = tpl.get("toc", {})
        self.toc_include.set(toc.get("include", False))

        sig = tpl.get("signature", {})
        self.sig_align_right.set(sig.get("align_right", True))
        self.sig_include_date.set(sig.get("include_date", True))
        self.sig_date_format.set(sig.get("date_format", DATE_FORMATS[0]))

    def _collect_ui_to_template(self):
        """从 UI 控件收集数据到模板字典"""
        tpl = get_default_template()
        tpl["name"] = self.current_template.get("name", "未命名")
        tpl["source"] = self.current_template.get("source", "手动设置")

        tpl["cover"] = {
            "auto_generate": self.cover_auto.get(),
            "title_font": self.cover_title_font.get(),
            "title_size": self.cover_title_size.get(),
            "title_bold": self.cover_title_bold.get(),
            "title_align": self.cover_title_align.get(),
            "company_font": self.cover_company_font.get(),
            "company_size": self.cover_company_size.get(),
            "date_font": self.cover_date_font.get(),
            "date_size": self.cover_date_size.get(),
        }

        tpl["headings"] = {
            "heading1": {
                "font": self.h1_font.get(), "size": self.h1_size.get(),
                "bold": self.h1_bold.get(), "align": self.h1_align.get(),
                "space_before": int(self.h1_space_before.get()), "space_after": int(self.h1_space_after.get()),
            },
            "heading2": {
                "font": self.h2_font.get(), "size": self.h2_size.get(),
                "bold": self.h2_bold.get(), "align": self.h2_align.get(),
                "space_before": int(self.h2_space_before.get()), "space_after": int(self.h2_space_after.get()),
            },
            "heading3": {
                "font": self.h3_font.get(), "size": self.h3_size.get(),
                "bold": self.h3_bold.get(), "align": self.h3_align.get(),
                "space_before": int(self.h3_space_before.get()), "space_after": int(self.h3_space_after.get()),
            },
        }

        ls_val = self.body_line_spacing.get()
        try:
            ls = float(ls_val)
        except ValueError:
            ls = 1.5

        tpl["body"] = {
            "font": self.body_font.get(), "size": self.body_size.get(),
            "bold": self.body_bold.get(), "align": self.body_align.get(),
            "line_spacing": ls, "line_spacing_rule": "multiple",
            "first_line_indent": int(self.body_indent.get()),
        }

        tpl["table"] = {
            "header_bold": self.table_header_bold.get(),
            "header_shading": self.table_header_shading.get(),
            "header_shading_color": "D9E2F3",
            "content_align": self.table_content_align.get(),
            "auto_fit": self.table_auto_fit.get(),
        }

        tpl["image"] = {
            "center_align": self.image_center.get(),
            "max_width_cm": int(self.image_max_width.get()),
        }

        tpl["header"] = {
            "text": self.header_text_var.get(),
            "font": self.header_font.get(),
            "size": self.header_size.get(),
            "align": self.header_align.get(),
            "different_first_page": self.header_diff_first.get(),
        }

        tpl["footer"] = {
            "page_num_format": self.page_num_format.get(),
            "align": self.page_align.get(),
            "start_from": int(self.page_start.get()),
            "font": self.header_font.get(),
            "size": self.header_size.get(),
            "show_on_first_page": self.page_show_first.get(),
        }

        tpl["toc"] = {"include": self.toc_include.get()}

        tpl["signature"] = {
            "align_right": self.sig_align_right.get(),
            "include_date": self.sig_include_date.get(),
            "date_format": self.sig_date_format.get(),
        }

        return tpl

    # ========== 操作事件 ==========

    def _open_file(self):
        """打开报告文件"""
        filepath = filedialog.askopenfilename(
            title="选择报告文件",
            filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
        )
        if not filepath:
            return

        self.current_file = filepath
        self.file_path_var.set(filepath)

        stats = detect_heading_paragraphs(filepath)
        if stats:
            basename = os.path.basename(filepath)
            size = os.path.getsize(filepath)
            size_str = f"{size/1024:.1f} KB" if size < 1024*1024 else f"{size/1024/1024:.1f} MB"
            self.current_file_info = f"{basename} ({size_str})"
            self.file_info_var.set(self.current_file_info)

            info = f"已打开：{basename} | H1:{stats['h1']} H2:{stats['h2']} H3:{stats['h3']} 正文段:{stats['body']} 表格:{stats['tables']}"
            self.status_var.set(info)
            logger.info(f"打开文件：{filepath}")
        else:
            self.status_var.set(f"已打开：{os.path.basename(filepath)}（无法识别文档结构）")

    def _show_file_info(self):
        """显示文件详细信息"""
        if not self.current_file:
            messagebox.showinfo("提示", "请先打开一个报告文件")
            return

        stats = detect_heading_paragraphs(self.current_file)
        if stats:
            msg = (
                f"文件：{os.path.basename(self.current_file)}\n"
                f"大小：{os.path.getsize(self.current_file)/1024:.1f} KB\n"
                f"修改时间：{datetime.fromtimestamp(os.path.getmtime(self.current_file)).strftime('%Y-%m-%d %H:%M')}\n\n"
                f"文档结构：\n"
                f"  一级标题：{stats['h1']} 个\n"
                f"  二级标题：{stats['h2']} 个\n"
                f"  三级标题：{stats['h3']} 个\n"
                f"  正文段落：{stats['body']} 段\n"
                f"  表格数量：{stats['tables']} 个"
            )
            messagebox.showinfo("文件信息", msg)

    def _import_template(self):
        """从 docx 导入模板"""
        if not self.current_file:
            filepath = filedialog.askopenfilename(
                title="选择要导入的模板文件（docx）",
                filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
            )
            if not filepath:
                return
        else:
            result = messagebox.askyesno("导入模板", "是否使用当前打开的报告作为模板？\n\n选「是」使用当前报告\n选「否」另外选择文件")
            if result:
                filepath = self.current_file
            else:
                filepath = filedialog.askopenfilename(
                    title="选择要导入的模板文件（docx）",
                    filetypes=[("Word 文档", "*.docx"), ("所有文件", "*.*")]
                )
                if not filepath:
                    return

        self.status_var.set(f"正在分析模板：{os.path.basename(filepath)}...")
        self.root.update()

        try:
            template = extract_styles_from_docx(filepath)
            self.current_template = template
            self._load_template_to_ui()

            name = simpledialog.askstring("保存模板", "请为导入的模板命名：", initialvalue=os.path.splitext(os.path.basename(filepath))[0])
            if name:
                category = simpledialog.askstring("模板分类", f"选择分类：{', '.join(TEMPLATE_CATEGORIES)}", initialvalue="其他")
                if category and category in TEMPLATE_CATEGORIES:
                    save_template(template, name, category)
                    self._refresh_template_list()
                    self.status_var.set(f"✅ 已导入并保存模板：{name}")
                else:
                    self.status_var.set('✅ 模板已加载到界面，但未保存（点击"保存为模板"永久保存）')
            else:
                self.status_var.set("✅ 模板已加载到界面，可随时调整后保存")

        except Exception as e:
            logger.error(f"模板导入失败：{e}")
            messagebox.showerror("导入失败", f"无法从文件提取格式：\n{str(e)}")
            self.status_var.set("❌ 模板导入失败")

    def _save_as_template(self):
        """保存当前设置为模板"""
        template = self._collect_ui_to_template()

        name = simpledialog.askstring("保存模板", "请输入模板名称：", initialvalue=template.get("name", "我的模板"))
        if not name:
            return

        cat_window = tk.Toplevel(self.root)
        cat_window.title("选择分类")
        cat_window.geometry("300x200")
        cat_window.transient(self.root)
        cat_window.grab_set()

        ttk.Label(cat_window, text="选择模板分类：", font=('微软雅黑', 10)).pack(pady=(15, 5))

        cat_var = tk.StringVar(value="通用模板")
        for cat in TEMPLATE_CATEGORIES:
            ttk.Radiobutton(cat_window, text=cat, variable=cat_var, value=cat).pack(anchor=tk.W, padx=30, pady=3)

        def do_save():
            category = cat_var.get()
            success, msg = save_template(template, name, category)
            if success:
                self._refresh_template_list()
                self.status_var.set(f"✅ {msg}")
            else:
                messagebox.showerror("保存失败", msg)
            cat_window.destroy()

        ttk.Button(cat_window, text="确定保存", command=do_save).pack(pady=15)

    def _delete_template(self):
        """删除选中的模板"""
        selection = self.template_tree.selection()
        if not selection:
            messagebox.showinfo("提示", "请先在左侧选中要删除的模板")
            return

        item = selection[0]
        parent = self.template_tree.parent(item)
        if not parent:
            messagebox.showinfo("提示", "请选择具体的模板，而非分类文件夹")
            return

        name = self.template_tree.item(item, "text").replace("  📄 ", "")
        filepath = self.template_tree.item(item, "tags")[0]

        if messagebox.askyesno("确认删除", f"确定要删除模板「{name}」吗？"):
            success, msg = delete_template(filepath)
            if success:
                self._refresh_template_list()
                self.status_var.set(f"🗑️ 已删除：{name}")
            else:
                messagebox.showerror("删除失败", msg)

    def _reset_to_default(self):
        """恢复默认模板"""
        if messagebox.askyesno("确认", "恢复默认将清空当前所有设置，确定吗？"):
            self.current_template = get_default_template()
            self._load_template_to_ui()
            self.status_var.set("已恢复默认模板")

    def _apply_template(self):
        """一键应用模板到当前报告"""
        if not self.current_file:
            messagebox.showwarning("提示", "请先打开一个报告文件（.docx）")
            return

        basename = os.path.basename(self.current_file)
        if not messagebox.askyesno("确认应用", f"确定要将当前模板应用到：\n\n{basename}\n\n吗？"):
            return

        # 备份
        if self.backup_var.get():
            self.status_var.set("正在备份原文件...")
            self.root.update()
            success, backup_path = backup_file(self.current_file)
            if success:
                self.status_var.set(f"✅ 已备份：{os.path.basename(backup_path)}")
            else:
                if not messagebox.askyesno("备份失败", f"文件备份失败：{backup_path}\n\n是否继续应用模板？"):
                    return

        template = self._collect_ui_to_template()

        self.status_var.set("⏳ 正在应用模板格式...")
        self.root.update()

        success, msg = apply_template_to_docx(self.current_file, template)

        if success:
            messagebox.showinfo("✅ 完成", f"格式已成功应用到：\n{basename}\n\n{msg}")
            self.status_var.set(f"✅ 格式应用完成：{basename}")
            logger.info(f"模板应用成功：{self.current_file}")
        else:
            messagebox.showerror("❌ 失败", msg)
            self.status_var.set("❌ 格式应用失败")
            logger.error(f"模板应用失败：{msg}")

    def _on_close(self):
        """关闭程序（带确认）"""
        if messagebox.askyesno("退出确认", "确定要退出底稿整理助手吗？"):
            logger.info("程序退出")
            self.root.destroy()


def main():
    """主入口"""
    root = tk.Tk()
    app = FormatTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
