import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
from PIL import Image, ImageTk, ImageDraw, ImageFont
import tempfile
import subprocess
import win32print
import sys
from datetime import datetime
import json
import glob
import zipfile
import xml.etree.ElementTree as ET

class BarcodeLabelTool:
    def __init__(self, root):
        self.root = root
        self.root.title("扫码出标签工具")
        self.root.geometry("800x600")
        
        # 数据存储
        self.data = None
        self.data_columns = []
        self.order_column = None
        self.tracking_column = None
        self.auto_print = tk.BooleanVar(value=True)  # 默认开启自动打印
        
        # DPI设置 (用于毫米到像素的转换)
        self.dpi = 300  # 标准打印DPI
        
        # 标签尺寸定义 (单位: 毫米)
        self.label_sizes = {
            "100x100": (100, 100),
            "100x70": (100, 70),
            "100x150": (100, 150)
        }
        
        # 条形码固定参数
        self.barcode_width_mm = 80
        self.barcode_height_mm = 20
        self.top_margin_mm = 10
        self.barcode_density_scale = 1.6

        # 文字显示配置（独立于条码自带文本）
        self.text_font_size_pt = 20  # 字号为20
        self.text_margin_top_mm = 3  # 条码下方到文字的间距（毫米）
        self.text_color = 'black'    # 文字颜色
        
        # 创建界面
        self.create_widgets()
        
        # 启动时清理非当天文件
        self.cleanup_old_files()
        
    def cleanup_old_files(self):
        """清理非当天的条码文件"""
        try:
            today = datetime.now().date()
            deleted_count = 0
            
            # 查找所有可能的条码文件
            file_patterns = [
                "label_*.png",
                "label_*.pdf", 
                "barcode_*.png"
            ]
            
            for pattern in file_patterns:
                for file_path in glob.glob(pattern):
                    try:
                        # 获取文件修改时间
                        file_time = datetime.fromtimestamp(os.path.getmtime(file_path))
                        file_date = file_time.date()
                        
                        # 如果不是今天的文件，删除
                        if file_date != today:
                            os.remove(file_path)
                            deleted_count += 1
                            self.log_event(f"清理旧文件: {file_path}")
                    except Exception as e:
                        # 如果删除失败，记录但继续处理其他文件
                        self.log_event(f"清理文件失败 {file_path}: {str(e)}")
                        continue
            
            if deleted_count > 0:
                self.log_event(f"清理完成，共删除 {deleted_count} 个非当天文件")
            else:
                self.log_event("无需要清理的旧文件")
                
        except Exception as e:
            self.log_event(f"清理旧文件时出错: {str(e)}")
        
    def mm_to_pixels(self, mm):
        """将毫米转换为像素"""
        # 使用四舍五入提高尺寸精度，保证打印物理尺寸更贴近设定值
        return round(mm * self.dpi / 25.4)

    def pt_to_pixels(self, pt):
        """将磅值(pt)转换为像素(px)，基于DPI"""
        return round(pt * self.dpi / 72)

    def load_font(self, pixel_size):
        """尝试加载常见中文/英文字体，失败则回退默认字体"""
        font_dirs = []
        windir = os.environ.get('WINDIR', 'C:\\Windows')
        font_dirs.append(os.path.join(windir, 'Fonts'))
        candidates = [
            'msyh.ttc',      # 微软雅黑（TrueType集合）
            'msyh.ttf',      # 微软雅黑（单字体）
            'SimHei.ttf',    # 黑体
            'arial.ttf',     # Arial
        ]
        for d in font_dirs:
            for name in candidates:
                fp = os.path.join(d, name)
                try:
                    return ImageFont.truetype(fp, pixel_size)
                except Exception:
                    pass
        return ImageFont.load_default()
    
    def create_widgets(self):
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 扫码处理（置顶，输入框更大更醒目）
        scan_frame = ttk.LabelFrame(main_frame, text="扫码处理", padding="8")
        scan_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        tk.Label(scan_frame, text="扫描订单号:", font=("Microsoft YaHei", 10, "bold")).grid(row=0, column=0, padx=(0, 8))
        self.scan_entry = tk.Entry(scan_frame, width=50, font=("Microsoft YaHei", 13))
        self.scan_entry.grid(row=0, column=1, padx=(0, 12), sticky=(tk.W, tk.E))
        self.scan_entry.bind('<Return>', self.process_scan)
        self.scan_entry.focus_set()

        # 不再在界面显示转单号；仅保留内部变量用于流程
        self.tracking_var = tk.StringVar()
        # 自动打印复选框始终显示，字体与输入框一致
        self.auto_print_chk = tk.Checkbutton(scan_frame, text="自动打印", variable=self.auto_print, font=("Microsoft YaHei", 13))
        self.auto_print_chk.grid(row=0, column=2, padx=(10, 0))

        # 表格映射（合并数据导入与列映射）
        self.mapping_frame = ttk.LabelFrame(main_frame, text="表格映射", padding="8")
        self.mapping_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # 上行：导入文件
        ttk.Button(self.mapping_frame, text="导入Excel文件", command=self.import_excel).grid(row=0, column=0, padx=(0, 10), sticky=tk.W)
        self.file_label = ttk.Label(self.mapping_frame, text="未选择文件")
        self.file_label.grid(row=0, column=1, sticky=tk.W)

        # 下行：列映射
        ttk.Label(self.mapping_frame, text="订单号列:").grid(row=1, column=0, padx=(0, 5), pady=(6,0))
        self.order_combo = ttk.Combobox(self.mapping_frame, state="readonly")
        self.order_combo.grid(row=1, column=1, padx=(0, 10), pady=(6,0))

        ttk.Label(self.mapping_frame, text="转单号列:").grid(row=1, column=2, padx=(0, 5), pady=(6,0))
        self.tracking_combo = ttk.Combobox(self.mapping_frame, state="readonly")
        self.tracking_combo.grid(row=1, column=3, padx=(0, 10), pady=(6,0))

        ttk.Button(self.mapping_frame, text="确认映射", command=self.confirm_mapping).grid(row=1, column=4, pady=(6,0))

        # 打印设置（合并标签与打印机）
        print_settings = ttk.LabelFrame(main_frame, text="打印设置", padding="8")
        print_settings.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # 第一行：标签格式（下拉）与实际尺寸
        ttk.Label(print_settings, text="标签格式:").grid(row=0, column=0, padx=(0, 5))
        self.label_format_var = tk.StringVar(value="100x100")
        label_formats = list(self.label_sizes.keys())
        self.label_format_combo = ttk.Combobox(print_settings, textvariable=self.label_format_var, values=label_formats, state="readonly", width=10)
        self.label_format_combo.grid(row=0, column=1, padx=(0, 10))
        self.label_format_combo.bind('<<ComboboxSelected>>', self.update_preview)

        ttk.Label(print_settings, text="实际尺寸(毫米):").grid(row=0, column=2, padx=(10, 5))
        self.actual_size_label = ttk.Label(print_settings, text="100x100")
        self.actual_size_label.grid(row=0, column=3)

        # 第二行：打印机下拉与手动打印按钮
        ttk.Label(print_settings, text="选择打印机:").grid(row=1, column=0, padx=(0, 5), pady=(6,0))
        self.printer_combo = ttk.Combobox(print_settings, state="readonly", width=30)
        self.printer_combo.grid(row=1, column=1, padx=(0, 10), pady=(6,0))
        self.load_printers()

        ttk.Button(print_settings, text="手动打印条形码", command=self.print_barcode).grid(row=1, column=2, pady=(6,0))
        ttk.Button(print_settings, text="保存当前配置", command=self.save_config).grid(row=1, column=3, pady=(6,0))

        # 日志区域（移除预览，仅显示日志）
        log_frame = ttk.LabelFrame(main_frame, text="日志", padding="8")
        log_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        log_container = ttk.Frame(log_frame)
        log_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_container.columnconfigure(0, weight=1)
        log_container.rowconfigure(0, weight=1)
        log_scroll = ttk.Scrollbar(log_container, orient=tk.VERTICAL)
        self.log_text = tk.Text(log_container, height=18, font=("Microsoft YaHei", 9), yscrollcommand=log_scroll.set)
        log_scroll.config(command=self.log_text.yview)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        log_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.log_text.config(state=tk.DISABLED)

        # 配置网格权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)
        
        # 应用已保存配置（如果存在）
        self.apply_config_defaults()
        
    def _col_letters_to_index(self, letters):
        result = 0
        for ch in letters:
            if ch.isalpha():
                result = result * 26 + (ord(ch.upper()) - ord("A") + 1)
        return result - 1

    def load_xlsx_simple(self, file_path):
        with zipfile.ZipFile(file_path) as zf:
            shared_strings = []
            try:
                with zf.open("xl/sharedStrings.xml") as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                    for si in root.findall("s:si", ns):
                        parts = []
                        for t in si.findall(".//s:t", ns):
                            if t.text:
                                parts.append(t.text)
                        shared_strings.append("".join(parts))
            except KeyError:
                shared_strings = []

            with zf.open("xl/worksheets/sheet1.xml") as f:
                tree = ET.parse(f)
                root = tree.getroot()
                ns = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
                rows = []
                for row in root.findall("s:sheetData/s:row", ns):
                    cells = {}
                    max_col = -1
                    for c in row.findall("s:c", ns):
                        ref = c.get("r")
                        col_letters = ""
                        if ref:
                            for ch in ref:
                                if ch.isalpha():
                                    col_letters += ch
                        col_idx = self._col_letters_to_index(col_letters) if col_letters else 0
                        if col_idx > max_col:
                            max_col = col_idx
                        cell_type = c.get("t")
                        v = c.find("s:v", ns)
                        value = ""
                        if cell_type == "s":
                            if v is not None and v.text is not None:
                                idx = int(v.text)
                                if 0 <= idx < len(shared_strings):
                                    value = shared_strings[idx]
                        else:
                            if v is not None and v.text is not None:
                                value = v.text
                        cells[col_idx] = "" if value is None else str(value)
                    if max_col == -1:
                        rows.append([])
                    else:
                        row_vals = []
                        for i in range(max_col + 1):
                            row_vals.append(cells.get(i, ""))
                        rows.append(row_vals)

        if not rows:
            return [], []
        header_row = rows[0]
        headers = []
        for i, h in enumerate(header_row):
            s = str(h).strip() if h is not None else ""
            if not s:
                s = f"列{i+1}"
            headers.append(s)
        data_rows = []
        for r in rows[1:]:
            if not any(str(v).strip() for v in r):
                continue
            row_dict = {}
            for i, h in enumerate(headers):
                value = ""
                if i < len(r) and r[i] is not None:
                    value = str(r[i])
                row_dict[h] = value
            data_rows.append(row_dict)
        return headers, data_rows

    def import_excel(self):
        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                headers, rows = self.load_xlsx_simple(file_path)
                self.data = rows
                self.data_columns = headers
                self.file_label.config(text=os.path.basename(file_path))
                
                self.order_column = None
                self.tracking_column = None
                
                for col in headers:
                    col_lower = col.lower()
                    if any(keyword in col_lower for keyword in ['订单', 'order', '编号', 'id']):
                        self.order_column = col
                    elif any(keyword in col_lower for keyword in ['转单', 'tracking', '快递', '运单']):
                        self.tracking_column = col
                
                self.mapping_frame.grid()
                
                columns = headers
                self.order_combo['values'] = columns
                self.tracking_combo['values'] = columns
                
                if self.order_column:
                    self.order_combo.set(self.order_column)
                if self.tracking_column:
                    self.tracking_combo.set(self.tracking_column)
                    
                self.log_event(f"成功导入文件，共{len(self.data)}条记录")
                
            except Exception as e:
                self.log_event(f"导入文件失败: {str(e)}")
    
    def confirm_mapping(self):
        self.order_column = self.order_combo.get()
        self.tracking_column = self.tracking_combo.get()
        
        if not self.order_column or not self.tracking_column:
            self.log_event("错误：请选择订单号和转单号列")
            return
            
        self.log_event(f"列映射已设置: 订单号列={self.order_column}, 转单号列={self.tracking_column}")
    
    def find_row_by_order(self, order_number):
        if not self.data:
            return None
        if not self.order_column:
            return None
        target = str(order_number).strip()
        if not target:
            return None
        for row in self.data:
            value = str(row.get(self.order_column, "")).strip()
            if value == target:
                return row
        return None
    
    def process_scan(self, event=None):
        self.last_action_start = datetime.now()
        if not self.data:
            self.log_event("错误：请先导入Excel文件")
            return
            
        if not self.order_column or not self.tracking_column:
            self.log_event("错误：请先设置列映射")
            return
            
        order_number = self.scan_entry.get().strip()
        if not order_number:
            self.log_event("错误：请输入订单号")
            return
            
        row = self.find_row_by_order(order_number)
        if row is None:
            self.log_event(f"错误：未找到订单号: {order_number}")
            return
            
        tracking_number = str(row.get(self.tracking_column, "")).strip()
        if not tracking_number:
            self.log_event(f"错误：订单号 {order_number} 的转单号为空")
            return
        self.tracking_var.set(tracking_number)
        self.last_tracking_number = tracking_number
        
        self.log_event(f"找到匹配订单: {order_number} -> 转单号: {tracking_number}")
        
        # 生成条形码和标签
        success = self.generate_label(tracking_number)
        
        # 如果启用了自动打印，且生成成功，再打印
        if success and self.auto_print.get():
            self.print_barcode()
        
        # 清空扫描输入框，准备下一次扫描
        self.scan_entry.delete(0, tk.END)
    
    def generate_label(self, tracking_number):
        """生成完整的标签图片和PDF。成功返回True，失败返回False"""
        try:
            # 清理旧文件，避免打印旧PDF
            label_png = f"label_{tracking_number}.png"
            label_pdf = f"label_{tracking_number}.pdf"
            if os.path.exists(label_png):
                try:
                    os.remove(label_png)
                except Exception:
                    pass
            if os.path.exists(label_pdf):
                try:
                    os.remove(label_pdf)
                except Exception:
                    pass

            barcode_filename = f"barcode_{tracking_number}.png"
            self.create_code128_barcode_pil(tracking_number, barcode_filename)
            
            # 创建完整的标签
            self.create_complete_label(barcode_filename, tracking_number)
            
            # 删除临时条形码文件
            try:
                os.remove(barcode_filename)
            except Exception:
                pass
            
            self.log_event(f"成功生成条码标签: {tracking_number}")
            return True
        except Exception as e:
            self.log_event(f"生成条形码失败: {str(e)}")
            return False

    def create_code128_barcode_pil(self, value, out_png):
        """使用 Pillow 绘制 Code128 条码 PNG（不带文字）"""
        text = str(value)
        if not text:
            raise ValueError("Code128编码内容不能为空")
        if not text.isdigit():
            raise ValueError("Code128子集C模式下仅支持数字字符")

        codes = []
        n = len(text)

        if n >= 2:
            codes.append(105)
            pair_len = n if n % 2 == 0 else n - 1
            for i in range(0, pair_len, 2):
                pair_val = int(text[i:i+2])
                codes.append(pair_val)
            if n % 2 == 1:
                codes.append(100)
                last_ch = text[-1]
                codes.append(ord(last_ch) - 32)
        else:
            codes.append(104)
            codes.append(ord(text) - 32)

        checksum = codes[0]
        for i, code in enumerate(codes[1:], start=1):
            checksum += code * i
        checksum %= 103

        sequence = codes + [checksum, 106]

        patterns = [
            "212222", "222122", "222221", "121223", "121322", "131222", "122213", "122312", "132212", "221213",
            "221312", "231212", "112232", "122132", "122231", "113222", "123122", "123221", "223211", "221132",
            "221231", "213212", "223112", "312131", "311222", "321122", "321221", "312212", "322112", "322211",
            "212123", "212321", "232121", "111323", "131123", "131321", "112313", "132113", "132311", "211313",
            "231113", "231311", "112133", "112331", "132131", "113123", "113321", "133121", "313121", "211331",
            "231131", "213113", "213311", "213131", "311123", "311321", "331121", "312113", "312311", "332111",
            "314111", "221411", "431111", "111224", "111422", "121124", "121421", "141122", "141221", "112214",
            "112412", "122114", "122411", "142112", "142211", "241211", "221114", "413111", "241112", "134111",
            "111242", "121142", "121241", "114212", "124112", "124211", "411212", "421112", "421211", "212141",
            "214121", "412121", "111143", "111341", "131141", "114113", "114311", "411113", "411311", "113141",
            "114131", "311141", "411131", "211412", "211214", "211232", "2331112"
        ]

        width_px = self.mm_to_pixels(self.barcode_width_mm)
        height_px = self.mm_to_pixels(self.barcode_height_mm)

        quiet_modules = 10
        total_modules = quiet_modules * 2
        for code in sequence:
            pattern = patterns[code]
            total_modules += sum(int(x) for x in pattern)

        module_px = width_px / float(total_modules)

        img = Image.new('RGB', (width_px, height_px), 'white')
        draw = ImageDraw.Draw(img)

        xf = quiet_modules * module_px
        for code in sequence:
            pattern = patterns[code]
            for i, ch in enumerate(pattern):
                units = int(ch)
                next_xf = xf + units * module_px
                xi = round(xf)
                xj = round(next_xf)
                if i % 2 == 0 and xj > xi:
                    draw.rectangle([xi, 0, xj - 1, height_px - 1], fill="black")
                xf = next_xf

        img.save(out_png, dpi=(self.dpi, self.dpi))

    create_code39_barcode_pil = create_code128_barcode_pil
    
    def create_complete_label(self, barcode_file, tracking_number):
        """创建完整的标签图片和PDF"""
        # 获取标签尺寸 (毫米)
        label_width_mm, label_height_mm = self.label_sizes[self.label_format_var.get()]
        
        # 更新实际尺寸显示
        self.actual_size_label.config(text=f"{label_width_mm}x{label_height_mm}")
        
        # 转换为像素
        label_width_px = self.mm_to_pixels(label_width_mm)
        label_height_px = self.mm_to_pixels(label_height_mm)
        
        # 创建标签画布
        label_img = Image.new('RGB', (label_width_px, label_height_px), 'white')
        
        # 加载条形码图片
        barcode_img = Image.open(barcode_file)
        
        # 转换为像素
        barcode_width_px = self.mm_to_pixels(self.barcode_width_mm)
        barcode_height_px = self.mm_to_pixels(self.barcode_height_mm)
        
        # 将条形码调整到固定尺寸，优先保证横向尺寸为80mm
        # 使用NEAREST避免抗锯齿造成的条纹模糊，提升扫码成功率
        barcode_img = barcode_img.resize((barcode_width_px, barcode_height_px), Image.NEAREST)
        
        # 计算水平居中位置和上部位置 (毫米)
        top_margin_px = self.mm_to_pixels(self.top_margin_mm)
        
        # 计算水平居中位置
        x = (label_width_px - barcode_width_px) // 2
        y = top_margin_px
        
        # 将条形码粘贴到标签上
        label_img.paste(barcode_img, (x, y))

        # 绘制独立文字（不使用条码自带文字）
        draw = ImageDraw.Draw(label_img)
        font_px = self.pt_to_pixels(self.text_font_size_pt)
        font = self.load_font(font_px)
        text = str(tracking_number)
        try:
            # 优先使用textbbox以获得更准确的尺寸
            bbox = draw.textbbox((0, 0), text, font=font)
            text_w = bbox[2] - bbox[0]
            text_h = bbox[3] - bbox[1]
        except Exception:
            text_w, text_h = draw.textsize(text, font=font)
        text_x = (label_width_px - text_w) // 2
        text_y = y + barcode_height_px + self.mm_to_pixels(self.text_margin_top_mm)
        # 防止文字超出底部
        if text_y + text_h > label_height_px:
            text_y = max(0, label_height_px - text_h - self.mm_to_pixels(2))
        draw.text((text_x, text_y), text, fill=self.text_color, font=font)
        
        # 保存完整的标签图片用于预览
        label_filename = f"label_{tracking_number}.png"
        label_img.save(label_filename, dpi=(self.dpi, self.dpi))
        
        # 创建PDF版本
        self.create_pdf_label(label_img, tracking_number, label_width_mm, label_height_mm)
    
    def create_pdf_label(self, label_img, tracking_number, width_mm, height_mm):
        """创建PDF版本的标签"""
        try:
            # 导入reportlab库
            from reportlab.pdfgen import canvas
            from reportlab.lib.pagesizes import mm
            from reportlab.lib.utils import ImageReader
            
            # 创建PDF文件
            pdf_filename = f"label_{tracking_number}.pdf"
            c = canvas.Canvas(pdf_filename, pagesize=(width_mm*mm, height_mm*mm))
            
            # 将PIL图像转换为reportlab可用的图像
            img_reader = ImageReader(label_img)
            
            # 在PDF上绘制图像，填满整个页面
            c.drawImage(img_reader, 0, 0, width=width_mm*mm, height=height_mm*mm)
            
            # 保存PDF
            c.save()
            
        except ImportError:
            self.log_event("错误：请安装reportlab库: pip install reportlab")
        except Exception as e:
            self.log_event(f"创建PDF失败: {str(e)}")
    
    def update_preview(self, event=None):
        """更新实际尺寸文本显示。"""
        try:
            fmt = self.label_format_var.get()
            self.actual_size_label.config(text=fmt)
        except Exception as e:
            self.log_event(f"更新尺寸显示失败: {e}")

    # 配置保存/加载
    def config_path(self):
        base = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base, 'label_change_config.json')

    def save_config(self):
        cfg = {
            'label_format': self.label_format_var.get(),
            'printer_name': self.printer_combo.get(),
            'auto_print': bool(self.auto_print.get()),
        }
        try:
            with open(self.config_path(), 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            self.log_event("配置已保存")
        except Exception as e:
            self.log_event(f"配置保存失败: {e}")

    def load_config(self):
        try:
            with open(self.config_path(), 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            return {}

    def apply_config_defaults(self):
        cfg = self.load_config()
        # 自动打印
        if 'auto_print' in cfg:
            try:
                self.auto_print.set(bool(cfg['auto_print']))
            except Exception:
                pass
        # 标签格式
        if 'label_format' in cfg and cfg['label_format'] in self.label_sizes:
            try:
                self.label_format_var.set(cfg['label_format'])
                self.update_preview()
            except Exception:
                pass
        # 打印机选择
        if 'printer_name' in cfg:
            printers = list(self.printer_combo['values'])
            if cfg['printer_name'] in printers:
                try:
                    self.printer_combo.set(cfg['printer_name'])
                except Exception:
                    pass
    
    def load_printers(self):
        printers = []
        try:
            for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL):
                printers.append(printer[2])
            self.printer_combo['values'] = printers
            if printers:
                self.printer_combo.set(printers[0])
        except:
            self.log_event("错误：无法获取打印机列表")
    
    def print_barcode(self):
        tracking_number = self.tracking_var.get()
        if not tracking_number:
            self.log_event("错误：没有可打印的条形码")
            return
            
        printer_name = self.printer_combo.get()
        if not printer_name:
            self.log_event("错误：请选择打印机")
            return
            
        pdf_file = f"label_{tracking_number}.pdf"
        if not os.path.exists(pdf_file):
            self.log_event("错误：PDF文件不存在")
            return
            
        # 获取程序所在目录
        if getattr(sys, 'frozen', False):
            # 打包后的exe文件所在目录
            application_path = os.path.dirname(sys.executable)
        else:
            # 脚本文件所在目录
            application_path = os.path.dirname(os.path.abspath(__file__))
            
        sumatra_path = os.path.join(application_path, "SumatraPDF.exe")
        
        if not os.path.exists(sumatra_path):
            self.log_event(f"错误：SumatraPDF.exe不存在于程序目录: {sumatra_path}")
            return
            
        try:
            # 使用SumatraPDF静默打印
            cmd = [
                sumatra_path,
                "-print-to", printer_name,
                "-silent",
                pdf_file
            ]
            
            # 执行打印命令
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                self.log_event(f"打印任务已发送: {tracking_number} -> {printer_name}")
                # 记录日志（含耗时）
                elapsed = None
                try:
                    if hasattr(self, 'last_action_start') and self.last_action_start:
                        elapsed = (datetime.now() - self.last_action_start).total_seconds()
                except Exception:
                    pass
                if elapsed is not None:
                    self.log_event(f"打印条码 {tracking_number} 到打印机 {printer_name}，耗时 {elapsed:.2f}s")
                else:
                    self.log_event(f"打印条码 {tracking_number} 到打印机 {printer_name}")
            else:
                # 如果打印失败，尝试使用默认打印机
                self.log_event(f"打印到指定打印机失败，尝试使用默认打印机: {result.stderr}")
                self.print_with_default_printer(pdf_file, sumatra_path)
            
        except subprocess.TimeoutExpired:
            self.log_event(f"打印超时: {pdf_file}")
        except Exception as e:
            self.log_event(f"打印失败: {pdf_file}，错误: {str(e)}")
    
    def print_with_default_printer(self, pdf_file, sumatra_path):
        """使用默认打印机打印"""
        try:
            cmd = [
                sumatra_path,
                "-print-to-default",
                "-silent",
                pdf_file
            ]
            
            result = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            
            if result.returncode == 0:
                self.log_event(f"默认打印机打印成功: {pdf_file}")
            else:
                self.log_event(f"默认打印机打印失败: {pdf_file}，错误: {result.stderr}")
        except Exception as e:
            self.log_event(f"默认打印机打印失败: {pdf_file}，异常: {str(e)}")

    def log_event(self, text):
        """将事件写入右侧日志窗口，带时间戳"""
        try:
            ts = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            line = f"[{ts}] {text}\n"
            if hasattr(self, 'log_text') and self.log_text:
                self.log_text.config(state=tk.NORMAL)
                self.log_text.insert(tk.END, line)
                self.log_text.see(tk.END)
                self.log_text.config(state=tk.DISABLED)
        except Exception:
            pass

if __name__ == "__main__":
    root = tk.Tk()
    app = BarcodeLabelTool(root)
    root.mainloop()
