import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import fitz  # PyMuPDF
from PIL import Image, ImageTk, ImageChops, ImageOps
import io
import re
import xml.etree.ElementTree as ET
from tkinter import ttk


class PdfSvgGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF 交互式提取 SVG 工具")

        # 顶部工具栏
        toolbar = tk.Frame(root)
        toolbar.pack(fill=tk.X)

        tk.Button(toolbar, text="打开PDF", command=self.open_pdf).pack(side=tk.LEFT, padx=4, pady=4)
        tk.Button(toolbar, text="上一页", command=self.prev_page).pack(side=tk.LEFT, padx=4, pady=4)
        tk.Button(toolbar, text="下一页", command=self.next_page).pack(side=tk.LEFT, padx=4, pady=4)

        tk.Label(toolbar, text="DPI:").pack(side=tk.LEFT)
        self.dpi_var = tk.StringVar(value="300")
        tk.Entry(toolbar, textvariable=self.dpi_var, width=5).pack(side=tk.LEFT)

        # 选框比例设置：预设 + 自定义
        tk.Label(toolbar, text="选框比例:").pack(side=tk.LEFT, padx=(8, 0))
        self.aspect_var = tk.StringVar(value="自由")
        tk.OptionMenu(toolbar, self.aspect_var, "自由", "1:1", "4:3", "3:2", "16:9", "9:16").pack(side=tk.LEFT)
        tk.Label(toolbar, text="自定义 W:H").pack(side=tk.LEFT, padx=(8, 0))
        self.custom_aspect_var = tk.StringVar(value="")
        tk.Entry(toolbar, textvariable=self.custom_aspect_var, width=10).pack(side=tk.LEFT)

        tk.Button(toolbar, text="导出SVG", command=self.export_svg).pack(side=tk.LEFT, padx=8)
        tk.Button(toolbar, text="导出PNG", command=self.export_png).pack(side=tk.LEFT)
        tk.Button(toolbar, text="批量导出图片", command=self.batch_export_images).pack(side=tk.LEFT, padx=8)
        # 去除白底背景开关
        self.remove_bg_var = tk.BooleanVar(value=False)
        tk.Checkbutton(toolbar, text="去除白底背景", variable=self.remove_bg_var).pack(side=tk.LEFT, padx=8)

        self.page_label = tk.Label(toolbar, text="")
        self.page_label.pack(side=tk.RIGHT, padx=8)

        # 底部状态栏
        self.status_var = tk.StringVar(value="就绪")
        status_bar = tk.Label(root, textvariable=self.status_var, anchor="w")
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        # 画布与滚动
        self.canvas = tk.Canvas(root, bg="#222")
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.canvas.bind("<ButtonPress-1>", self.on_mouse_down)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_mouse_up)

        # 状态
        self.doc = None
        self.page_index = 0
        self.photo = None
        self.zoom = 2.0  # 渲染缩放（72dpi 基础上的倍率）
        self.scale = 1.0  # 图片缩放到画布的比例
        self.img_w = 0
        self.img_h = 0
        self.sel_rect = None  # (x0,y0,x1,y1) in canvas coords
        self.sel_id = None
        # 最近一次导出的/生成的 SVG 缓存
        self.last_svg = None
        self.last_svg_size = (0, 0)  # (width, height)
        self.last_svg_name = "extracted"
        self.last_rect = None  # 上次生成 SVG 的页面选区矩形
        # 批量导出选项（默认）
        self.export_formats = {
            "PNG": tk.BooleanVar(value=True),
            "WEBP": tk.BooleanVar(value=True),
            "JPG": tk.BooleanVar(value=True),
            "ICO": tk.BooleanVar(value=True),  # 新增 ICO 图标格式
        }
        # 常见尺寸（包含更小图标尺寸）
        self.export_sizes = [16, 24, 32, 48, 64, 96, 128, 256, 512, 1024]
        self.size_vars = {s: tk.BooleanVar(value=True) for s in self.export_sizes}
        self.custom_sizes_var = tk.StringVar(value="")

    def open_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf"), ("All Files", "*.*")])
        if not path:
            return
        try:
            self.doc = fitz.open(path)
            self.page_index = 0
            self.render_page()
        except Exception as e:
            messagebox.showerror("打开失败", str(e))

    def render_page(self):
        if not self.doc:
            return
        page = self.doc[self.page_index]
        mat = fitz.Matrix(self.zoom, self.zoom)
        pix = page.get_pixmap(matrix=mat, alpha=True)

        self.img_w, self.img_h = pix.width, pix.height
        # 画布尺寸
        cw = max(400, min(self.img_w, self.root.winfo_screenwidth() - 80))
        ch = max(300, min(self.img_h, self.root.winfo_screenheight() - 160))
        self.canvas.config(width=cw, height=ch)

        # 等比缩放到画布
        self.scale = min(cw / self.img_w, ch / self.img_h)
        disp_w = int(self.img_w * self.scale)
        disp_h = int(self.img_h * self.scale)

        img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
        img = img.resize((disp_w, disp_h), Image.LANCZOS)
        self.photo = ImageTk.PhotoImage(img)

        self.canvas.delete("all")
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo)
        self.page_label.config(text=f"第 {self.page_index + 1}/{self.doc.page_count} 页")
        self.sel_rect = None
        self.sel_id = None

    def prev_page(self):
        if not self.doc:
            return
        if self.page_index > 0:
            self.page_index -= 1
            self.render_page()

    def next_page(self):
        if not self.doc:
            return
        if self.page_index < self.doc.page_count - 1:
            self.page_index += 1
            self.render_page()

    def on_mouse_down(self, e):
        self.sel_rect = (e.x, e.y, e.x, e.y)
        if self.sel_id:
            self.canvas.delete(self.sel_id)
            self.sel_id = None

    def on_mouse_drag(self, e):
        if not self.sel_rect:
            return
        x0, y0, _, _ = self.sel_rect
        # 根据选项应用长宽比约束
        r = self._get_aspect_ratio()
        if r is None:
            x1, y1 = e.x, e.y
        else:
            dx = e.x - x0
            dy = e.y - y0
            # 若横向变化较明显，以宽度为基准；否则以高度为基准
            if abs(dx) >= abs(dy):
                x1 = e.x
                height = abs(dx) / r
                y1 = y0 + (1 if dy >= 0 else -1) * height
            else:
                y1 = e.y
                width = abs(dy) * r
                x1 = x0 + (1 if dx >= 0 else -1) * width
        self.sel_rect = (x0, y0, x1, y1)
        if self.sel_id:
            self.canvas.coords(self.sel_id, x0, y0, x1, y1)
        else:
            self.sel_id = self.canvas.create_rectangle(x0, y0, x1, y1, outline="#00ff88", width=2)

    def on_mouse_up(self, e):
        if not self.sel_rect:
            return
        x0, y0, _, _ = self.sel_rect
        # 在鼠标松开时同样应用长宽比约束，保持与拖拽时一致
        r = self._get_aspect_ratio()
        if r is None:
            x1, y1 = e.x, e.y
        else:
            dx = e.x - x0
            dy = e.y - y0
            if abs(dx) >= abs(dy):
                x1 = e.x
                height = abs(dx) / r
                y1 = y0 + (1 if dy >= 0 else -1) * height
            else:
                y1 = e.y
                width = abs(dy) * r
                x1 = x0 + (1 if dx >= 0 else -1) * width
        self.sel_rect = (x0, y0, x1, y1)
        if self.sel_id:
            self.canvas.coords(self.sel_id, x0, y0, x1, y1)

    def _canvas_to_page_rect(self):
        if not self.doc:
            return None
        page = self.doc[self.page_index]
        if not self.sel_rect:
            return page.rect
        x0, y0, x1, y1 = self.sel_rect
        # 归一化坐标
        if x1 < x0:
            x0, x1 = x1, x0
        if y1 < y0:
            y0, y1 = y1, y0

        # 画布坐标 → 像素坐标
        px0 = x0 / self.scale
        py0 = y0 / self.scale
        px1 = x1 / self.scale
        py1 = y1 / self.scale

        # 像素坐标 → 页面坐标（72dpi 基础）
        rect = fitz.Rect(px0 / self.zoom, py0 / self.zoom, px1 / self.zoom, py1 / self.zoom)
        # 防越界
        r0 = page.rect
        rect.x0 = max(r0.x0, rect.x0)
        rect.y0 = max(r0.y0, rect.y0)
        rect.x1 = min(r0.x1, rect.x1)
        rect.y1 = min(r0.y1, rect.y1)
        return rect

    def _get_aspect_ratio(self):
        """
        返回选框的宽高比（宽/高）。
        - 若为自由模式，返回 None。
        - 若自定义输入存在（格式如 "W:H" 或单个数值），优先解析自定义。
        - 否则使用预设映射。
        """
        s = (self.custom_aspect_var.get() or "").strip()
        if s:
            try:
                if ":" in s:
                    w_str, h_str = s.split(":", 1)
                    w = float(w_str)
                    h = float(h_str)
                    if w > 0 and h > 0:
                        return w / h
                else:
                    r = float(s)
                    if r > 0:
                        return r
            except Exception:
                # 忽略解析错误，退回到预设
                pass
        preset = self.aspect_var.get()
        mapping = {
            "自由": None,
            "1:1": 1.0,
            "4:3": 4 / 3,
            "3:2": 3 / 2,
            "16:9": 16 / 9,
            "9:16": 9 / 16,
        }
        return mapping.get(preset, None)

    def export_png(self):
        if not self.doc:
            messagebox.showinfo("提示", "请先打开 PDF")
            return
        dpi = int(self.dpi_var.get() or 300)
        page = self.doc[self.page_index]
        rect = self._canvas_to_page_rect()
        mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
        pix = page.get_pixmap(matrix=mat, clip=rect, alpha=True)
        out = filedialog.asksaveasfilename(defaultextension=".png", filetypes=[("PNG", "*.png")], initialfile="extracted.png")
        if not out:
            return
        img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
        if self.remove_bg_var.get():
            try:
                img = self._remove_white_background(img)
            except Exception:
                pass
        img.save(out, format="PNG")
        messagebox.showinfo("完成", f"已导出 PNG: {out}")

    def _remove_white_background(self, img: Image.Image, threshold: int = 250) -> Image.Image:
        """
        将近白像素（R、G、B 都 >= threshold）透明化，并与原 Alpha 叠乘。
        """
        img = img.convert("RGBA")
        r, g, b, a = img.split()
        mask_r = r.point(lambda v: 255 if v >= threshold else 0)
        mask_g = g.point(lambda v: 255 if v >= threshold else 0)
        mask_b = b.point(lambda v: 255 if v >= threshold else 0)
        white_mask = ImageChops.multiply(mask_r, ImageChops.multiply(mask_g, mask_b))
        alpha_from_white = ImageOps.invert(white_mask)  # 白色->0，非白->255
        new_alpha = ImageChops.multiply(a, alpha_from_white)
        img.putalpha(new_alpha)
        return img

    def export_svg(self):
        if not self.doc:
            messagebox.showinfo("提示", "请先打开 PDF")
            return
        page = self.doc[self.page_index]
        rect = self._canvas_to_page_rect()

        # 通过将选区作为 clip 插入到一张新页面来实现严格裁剪
        try:
            tmp_doc = fitz.open()
            new_page = tmp_doc.new_page(width=rect.width, height=rect.height)
            # 将源 PDF 页的选定区域显示到新页上（坐标原点对齐到 (0,0)）
            new_page.show_pdf_page(new_page.rect, self.doc, self.page_index, clip=rect)

            # 生成新页的 SVG（其尺寸与选区一致）
            svg = new_page.get_svg_image()
            # 可选：去除白底背景（仅移除覆盖全画布的白色矩形），保持其它元素
            if self.remove_bg_var.get():
                try:
                    svg = self._remove_white_background_in_svg(svg, (float(rect.width), float(rect.height)))
                except Exception:
                    pass
        except Exception as e:
            messagebox.showerror("导出失败", f"生成裁剪 SVG 时出错: {e}")
            return
        finally:
            try:
                tmp_doc.close()
            except Exception:
                pass

        # 缓存 SVG 及其尺寸，便于后续批量导出
        self.last_svg = svg
        self.last_svg_size = (int(rect.width), int(rect.height))
        self.last_rect = rect

        out = filedialog.asksaveasfilename(defaultextension=".svg", filetypes=[("SVG", "*.svg")], initialfile="extracted.svg")
        if not out:
            return
        Path(out).write_text(svg, encoding="utf-8")
        try:
            self.last_svg_name = Path(out).stem or "extracted"
        except Exception:
            self.last_svg_name = "extracted"
        messagebox.showinfo("完成", f"已导出 SVG: {out}")

    def _remove_white_background_in_svg(self, svg: str, size: tuple) -> str:
        """
        移除/透明化覆盖整张画布的白色背景：
        - 处理 <rect>/<polygon>/<path> 中的白底图形
        - 白色判断：white/#fff/#ffffff/#fefefe 等近白、rgb(>=250,>=250,>=250)
        - 尺寸判断：接近画布尺寸或占比 >= 95%
        若解析失败，回退到更强的正则删除常见白底元素。
        """
        def _parse_float(val: str) -> float:
            try:
                m = re.search(r"[-+]?[0-9]*\.?[0-9]+", str(val))
                return float(m.group(0)) if m else 0.0
            except Exception:
                return 0.0

        def _hex_to_rgb(hexstr: str):
            h = hexstr.lstrip('#')
            if len(h) == 3:
                h = ''.join(ch*2 for ch in h)
            if len(h) != 6:
                return None
            try:
                return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
            except Exception:
                return None

        def _parse_rgb(fill: str):
            f = fill.strip().lower()
            m = re.match(r"rgba?\(\s*([\d.]+)\s*,\s*([\d.]+)\s*,\s*([\d.]+)\s*(?:,\s*[\d.]+\s*)?\)", f)
            if m:
                vals = []
                for g in m.groups()[:3]:
                    if '%' in g:
                        try:
                            vals.append(int(round(float(g.strip('%')) * 2.55)))
                        except Exception:
                            vals.append(255)
                    else:
                        try:
                            vals.append(int(round(float(g))))
                        except Exception:
                            vals.append(255)
                return tuple(vals)
            return None

        def _is_white(fill: str) -> bool:
            if not fill:
                return False
            f = fill.strip().lower().replace(" ", "")
            if f == "white":
                return True
            if f.startswith('#'):
                rgb = _hex_to_rgb(f)
                if rgb:
                    return all(c >= 250 for c in rgb)
                return False
            rgb = _parse_rgb(f)
            if rgb:
                return all(c >= 250 for c in rgb)
            return False

        def _ensure_transparent(child, style: str):
            # 透明化而不是删除，避免影响布局
            child.attrib["fill"] = "none"
            styles = {}
            if style:
                for part in style.split(";"):
                    if ":" in part:
                        k, v = part.split(":", 1)
                        styles[k.strip().lower()] = v.strip()
            styles.pop("fill", None)
            styles["fill"] = "none"
            styles["fill-opacity"] = "0"
            child.attrib["style"] = ";".join(f"{k}:{v}" for k, v in styles.items())

        try:
            root = ET.fromstring(svg)
            canvas_w, canvas_h = size
            tol = max(1.0, 0.02 * max(canvas_w, canvas_h))  # 2% 容差
            min_ratio = 0.95
            # 遍历候选元素，找到符合条件者并透明化
            removed = False
            for parent in root.iter():
                for child in list(parent):
                    tag = child.tag
                    if isinstance(tag, str) and (tag.endswith("rect") or tag.endswith("polygon") or tag.endswith("path")):
                        fill = child.attrib.get("fill")
                        style = child.attrib.get("style", "")
                        # 从 style 中提取 fill
                        if not fill and style:
                            for part in style.split(";"):
                                if ":" in part:
                                    k, v = part.split(":", 1)
                                    if k.strip().lower() == "fill":
                                        fill = v.strip()
                                        break
                        if not _is_white(fill):
                            continue
                        # stroke 检查：若存在描边，减少误伤
                        stroke = child.attrib.get("stroke") or ""
                        stroke_in_style = False
                        if style:
                            for part in style.split(";"):
                                if ":" in part:
                                    k, v = part.split(":", 1)
                                    if k.strip().lower() == "stroke" and v.strip().lower() not in {"none", ""}:
                                        stroke_in_style = True
                                        break
                        if stroke and stroke.lower() not in {"none", ""}:
                            continue
                        if stroke_in_style:
                            continue

                        fits_canvas = False
                        if tag.endswith("rect"):
                            w = _parse_float(child.attrib.get("width", canvas_w))
                            h = _parse_float(child.attrib.get("height", canvas_h))
                            x = _parse_float(child.attrib.get("x", 0))
                            y = _parse_float(child.attrib.get("y", 0))
                            if (abs(x) <= tol and abs(y) <= tol) and (w >= min_ratio * canvas_w and h >= min_ratio * canvas_h):
                                fits_canvas = True
                        elif tag.endswith("polygon"):
                            pts = child.attrib.get("points", "")
                            nums = re.findall(r"[-+]?\d*\.?\d+(?:e[-+]?\d+)?", pts)
                            coords = [float(n) for n in nums]
                            if len(coords) >= 8:
                                xs = coords[0::2]
                                ys = coords[1::2]
                                minx, maxx = min(xs), max(xs)
                                miny, maxy = min(ys), max(ys)
                                if abs(minx) <= tol and abs(miny) <= tol and (maxx >= min_ratio * canvas_w) and (maxy >= min_ratio * canvas_h):
                                    fits_canvas = True
                        elif tag.endswith("path"):
                            d = child.attrib.get("d", "")
                            nums = re.findall(r"[-+]?\d*\.?\d+(?:e[-+]?\d+)?", d)
                            coords = [float(n) for n in nums]
                            if len(coords) >= 8:
                                xs = coords[0::2]
                                ys = coords[1::2]
                                minx, maxx = min(xs), max(xs)
                                miny, maxy = min(ys), max(ys)
                                if abs(minx) <= tol and abs(miny) <= tol and (maxx >= min_ratio * canvas_w) and (maxy >= min_ratio * canvas_h):
                                    fits_canvas = True

                        if fits_canvas:
                            _ensure_transparent(child, style)
                            removed = True
            if removed:
                return ET.tostring(root, encoding="unicode")
        except Exception:
            pass

        # 回退：正则替换首个匹配的白底背景（rect/path/polygon）
        try:
            color_pat = r"(?:#(?:fff|ffffff|fefefe)|white|rgb\(\s*25[0-9]\s*,\s*25[0-9]\s*,\s*25[0-9]\s*\))"
            style_pat = rf"style\s*=\s*\"[^\"]*fill\s*:\s*{color_pat}[^\"]*\""
            fill_pat = rf"fill\s*=\s*\"{color_pat}\""
            rect_pat = rf"<rect[^>]*?(?:{fill_pat}|{style_pat})[^>]*?>"
            poly_pat = rf"<polygon[^>]*?(?:{fill_pat}|{style_pat})[^>]*?>"
            path_pat = rf"<path[^>]*?(?:{fill_pat}|{style_pat})[^>]*?>"
            pattern = rf"({rect_pat}|{poly_pat}|{path_pat})"
            return re.sub(pattern, "", svg, count=1, flags=re.IGNORECASE)
        except Exception:
            return svg

    def batch_export_images(self):
        """
        弹出对话框让用户选择导出格式与尺寸，然后按选择进行批量导出。
        """
        self._open_export_dialog()

    def _open_export_dialog(self):
        if not self.doc:
            messagebox.showinfo("提示", "请先打开 PDF")
            return
        dlg = tk.Toplevel(self.root)
        dlg.title("批量导出选项")
        dlg.transient(self.root)
        dlg.grab_set()
        frm = ttk.Frame(dlg, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="选择导出格式").grid(row=0, column=0, sticky="w")
        fmt_frame = ttk.Frame(frm)
        fmt_frame.grid(row=1, column=0, sticky="w")
        for i, (fmt, var) in enumerate(self.export_formats.items()):
            ttk.Checkbutton(fmt_frame, text=fmt, variable=var).grid(row=0, column=i, padx=6, pady=4)

        ttk.Label(frm, text="选择导出尺寸（按长边）").grid(row=2, column=0, sticky="w", pady=(8,0))
        size_frame = ttk.Frame(frm)
        size_frame.grid(row=3, column=0, sticky="w")
        for i, s in enumerate(self.export_sizes):
            ttk.Checkbutton(size_frame, text=str(s), variable=self.size_vars[s]).grid(row=i//6, column=i%6, padx=6, pady=4)

        # 自定义尺寸输入
        custom_frame = ttk.Frame(frm)
        custom_frame.grid(row=4, column=0, sticky="w", pady=(8,0))
        ttk.Label(custom_frame, text="自定义尺寸（逗号分隔，如 20,40,80）").pack(side=tk.LEFT)
        ttk.Entry(custom_frame, textvariable=self.custom_sizes_var, width=24).pack(side=tk.LEFT, padx=8)

        btn_frame = ttk.Frame(frm)
        btn_frame.grid(row=5, column=0, sticky="e", pady=(12,0))
        ttk.Button(btn_frame, text="开始导出", command=lambda: self._on_export_dialog_confirm(dlg)).pack(side=tk.RIGHT)
        ttk.Button(btn_frame, text="取消", command=dlg.destroy).pack(side=tk.RIGHT, padx=8)

    def _on_export_dialog_confirm(self, dlg: tk.Toplevel):
        """
        在关闭对话框前先校验用户选择；若无有效选择则提示并保持对话框。
        校验通过后释放抓取并关闭对话框，再异步触发批量导出。
        """
        # 校验至少选择一个格式
        formats = [fmt for fmt, v in self.export_formats.items() if v.get()]
        if not formats:
            messagebox.showinfo("提示", "请至少选择一种导出格式")
            return
        # 校验至少选择一个尺寸（含自定义）
        sizes = [s for s, v in self.size_vars.items() if v.get()]
        extra = []
        sraw = (self.custom_sizes_var.get() or "").strip()
        if sraw:
            sraw = sraw.replace("，", ",")
            for part in sraw.split(","):
                part = part.strip()
                if not part:
                    continue
                try:
                    val = int(part)
                    if val > 0:
                        extra.append(val)
                except Exception:
                    pass
        if not sizes and not extra:
            messagebox.showinfo("提示", "请至少选择一个导出尺寸")
            return
        # 释放抓取并关闭对话框，再稍后触发导出，避免窗口销毁影响弹窗/文件选择器
        try:
            dlg.grab_release()
        except Exception:
            pass
        dlg.destroy()
        # 异步调用，确保事件循环处理完窗口销毁
        self.root.after(10, self._perform_batch_export)

    def _perform_batch_export(self):
        # 准备进度窗口
        prog = tk.Toplevel(self.root)
        prog.title("批量导出进度")
        prog.transient(self.root)
        prog.geometry("480x300")
        prog_frame = ttk.Frame(prog, padding=10)
        prog_frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(prog_frame, text="正在导出...").pack(anchor="w")
        prog_bar = ttk.Progressbar(prog_frame, mode="determinate")
        prog_bar.pack(fill=tk.X, pady=8)
        log_text = tk.Text(prog_frame, height=10)
        log_text.pack(fill=tk.BOTH, expand=True)
        def log(msg: str):
            try:
                log_text.insert(tk.END, msg + "\n")
                log_text.see(tk.END)
                self.root.update_idletasks()
            except Exception:
                pass
        # 确保有 SVG 可用：若无缓存则从当前选区生成一次
        if not self.last_svg:
            if not self.doc:
                messagebox.showinfo("提示", "请先打开 PDF")
                prog.destroy()
                return
            try:
                page = self.doc[self.page_index]
                rect = self._canvas_to_page_rect()
                tmp_doc = fitz.open()
                new_page = tmp_doc.new_page(width=rect.width, height=rect.height)
                new_page.show_pdf_page(new_page.rect, self.doc, self.page_index, clip=rect)
                svg = new_page.get_svg_image()
                try:
                    tmp_doc.close()
                except Exception:
                    pass
                self.last_svg = svg
                self.last_svg_size = (int(rect.width), int(rect.height))
                self.last_rect = rect
            except Exception as e:
                messagebox.showerror("错误", f"生成 SVG 失败: {e}")
                prog.destroy()
                return

        svg = self.last_svg
        orig_w, orig_h = self.last_svg_size
        # 选择输出文件夹
        out_dir = filedialog.askdirectory(title="选择导出文件夹")
        if not out_dir:
            try:
                prog.destroy()
            except Exception:
                pass
            return

        # CairoSVG 可选：若不可用，回退到 PyMuPDF 渲染
        cairo_available = True
        try:
            import cairosvg
        except Exception as e:
            cairo_available = False
            log(f"CairoSVG 不可用，将使用 PyMuPDF 回退渲染。错误: {e}")

        # 收集用户勾选的尺寸与格式
        sizes = [s for s, v in self.size_vars.items() if v.get()]
        # 解析自定义尺寸，合并并去重
        extra = []
        sraw = (self.custom_sizes_var.get() or "").strip()
        if sraw:
            sraw = sraw.replace("，", ",")  # 兼容中文逗号
            for part in sraw.split(","):
                part = part.strip()
                if not part:
                    continue
                try:
                    val = int(part)
                    if val > 0:
                        extra.append(val)
                except Exception:
                    pass
        sizes = sorted(set(sizes + extra))
        if not sizes:
            messagebox.showinfo("提示", "请至少选择一个导出尺寸")
            return
        formats = [fmt for fmt, v in self.export_formats.items() if v.get()]
        if not formats:
            messagebox.showinfo("提示", "请至少选择一种导出格式")
            return

        base = self.last_svg_name or "extracted"
        out_path = Path(out_dir)

        exported = []
        total = len(sizes) * len(formats)
        done = 0
        prog_bar['maximum'] = max(1, total)
        self.status_var.set("开始导出...")
        for target in sizes:
            # 保持原始宽高比，按较长边为 target 缩放
            if orig_w <= 0 or orig_h <= 0:
                w = h = target
            elif orig_w >= orig_h:
                w = target
                h = max(1, int(round(target * orig_h / orig_w)))
            else:
                h = target
                w = max(1, int(round(target * orig_w / orig_h)))

            try:
                if cairo_available:
                    png_bytes = cairosvg.svg2png(bytestring=svg.encode("utf-8"), output_width=w, output_height=h)
                    img = Image.open(io.BytesIO(png_bytes))
                    if img.size != (w, h):
                        img = img.resize((w, h), Image.LANCZOS)
                else:
                    # 回退：使用 PyMuPDF 基于原始 PDF 选区进行渲染，再按目标尺寸缩放
                    rect = self.last_rect or self._canvas_to_page_rect()
                    tmp_doc = fitz.open()
                    new_page = tmp_doc.new_page(width=rect.width, height=rect.height)
                    new_page.show_pdf_page(new_page.rect, self.doc, self.page_index, clip=rect)
                    scale_x = w / rect.width
                    scale_y = h / rect.height
                    mat = fitz.Matrix(scale_x, scale_y)
                    pix = new_page.get_pixmap(matrix=mat, alpha=True)
                    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
                    try:
                        tmp_doc.close()
                    except Exception:
                        pass
                # 去除白底（可选），然后按目标尺寸确保尺寸一致
                if self.remove_bg_var.get():
                    try:
                        img = self._remove_white_background(img)
                    except Exception:
                        pass
                if img.size != (w, h):
                    img = img.resize((w, h), Image.LANCZOS)
                if "PNG" in formats:
                    png_path = out_path / f"{base}_{w}x{h}.png"
                    buf = io.BytesIO()
                    img.save(buf, format="PNG")
                    png_path.write_bytes(buf.getvalue())
                    done += 1
                    prog_bar['value'] = done
                    log(f"PNG: {png_path.name}")
                    exported.append(str(png_path))
                if "WEBP" in formats:
                    webp_path = out_path / f"{base}_{w}x{h}.webp"
                    try:
                        img.save(webp_path, format="WEBP", lossless=True)
                    except Exception:
                        img.save(webp_path, format="WEBP")
                    done += 1
                    prog_bar['value'] = done
                    log(f"WEBP: {webp_path.name}")
                    exported.append(str(webp_path))
                if "JPG" in formats:
                    jpg_path = out_path / f"{base}_{w}x{h}.jpg"
                    rgb = Image.new("RGB", (w, h), (255, 255, 255))
                    if img.mode in ("RGBA", "LA"):
                        alpha = img.split()[-1]
                        rgb.paste(img.convert("RGB"), mask=alpha)
                    else:
                        rgb.paste(img)
                    rgb.save(jpg_path, format="JPEG", quality=95)
                    done += 1
                    prog_bar['value'] = done
                    log(f"JPG: {jpg_path.name}")
                    exported.append(str(jpg_path))
                if "ICO" in formats:
                    # ICO 通常支持最大 256x256，超出尺寸跳过
                    if w > 256 or h > 256:
                        log(f"跳过 ICO 尺寸 {w}x{h}（ICO 最大为 256）")
                    else:
                        ico_path = out_path / f"{base}_{w}x{h}.ico"
                        try:
                            img_for_ico = img
                            if img_for_ico.mode not in ("RGBA", "RGB", "P"):
                                img_for_ico = img_for_ico.convert("RGBA")
                            # 去除白底（可选）
                            if self.remove_bg_var.get():
                                try:
                                    img_for_ico = self._remove_white_background(img_for_ico)
                                except Exception:
                                    pass
                            # 明确写入目标尺寸
                            img_for_ico.save(ico_path, format="ICO", sizes=[(w, h)])
                            done += 1
                            prog_bar['value'] = done
                            log(f"ICO: {ico_path.name}")
                            exported.append(str(ico_path))
                        except Exception as e:
                            log(f"ICO 导出失败 {w}x{h}: {e}")
            except Exception as e:
                log(f"失败: 尺寸 {target} 处理失败: {e}")
                messagebox.showerror("导出失败", f"尺寸 {target} 处理失败: {e}")
                break

        # 原始尺寸的 PNG 也导出一份（若可用）
        if orig_w > 0 and orig_h > 0:
            try:
                if cairo_available:
                    png_bytes = cairosvg.svg2png(bytestring=svg.encode("utf-8"), output_width=orig_w, output_height=orig_h)
                    img = Image.open(io.BytesIO(png_bytes))
                else:
                    rect = self.last_rect or self._canvas_to_page_rect()
                    tmp_doc = fitz.open()
                    new_page = tmp_doc.new_page(width=rect.width, height=rect.height)
                    new_page.show_pdf_page(new_page.rect, self.doc, self.page_index, clip=rect)
                    pix = new_page.get_pixmap(alpha=True)
                    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
                    try:
                        tmp_doc.close()
                    except Exception:
                        pass
                if self.remove_bg_var.get():
                    try:
                        img = self._remove_white_background(img)
                    except Exception:
                        pass
                orig_png = out_path / f"{base}_{orig_w}x{orig_h}.png"
                img.save(orig_png, format="PNG")
                exported.append(str(orig_png))
                done += 1
                prog_bar['value'] = done
                log(f"原始尺寸 PNG: {orig_png.name}")
            except Exception:
                pass

        # 完成
        self.status_var.set(f"导出完成，文件数: {len(exported)}")
        log(f"完成，总计导出文件: {len(exported)}")
        prog.title("批量导出完成")
        ttk.Button(prog_frame, text="关闭", command=prog.destroy).pack(anchor="e", pady=6)


def main():
    root = tk.Tk()
    app = PdfSvgGUI(root)
    root.geometry("1200x800")
    root.mainloop()


if __name__ == "__main__":
    main()