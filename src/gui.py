import os
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from datetime import date
from PIL import Image, ImageTk
from template_reader import read_template, get_default_format
from generator import generate
from parser import parse_product_text


class OrderApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("订货单生成器")
        self.root.geometry("1100x700")
        self.root.minsize(900, 500)

        self.template = None
        self.template_path = None
        self.photo_dir = None
        self.photo_files = []
        self.thumb_refs = []

        self._build_ui()

    def _build_ui(self):
        # === Top toolbar ===
        toolbar = tk.Frame(self.root, bg="#f0f0f0", height=48)
        toolbar.pack(fill="x", padx=10, pady=(8, 4))
        toolbar.pack_propagate(False)

        tk.Label(toolbar, text="订货单生成器", font=("微软雅黑", 14, "bold"),
                 bg="#f0f0f0").pack(side="left", padx=5)

        self.btn_template = tk.Button(toolbar, text="📁 选择模板（可选）",
                                      command=self._select_template,
                                      font=("微软雅黑", 10))
        self.btn_template.pack(side="left", padx=8)

        self.btn_photos = tk.Button(toolbar, text="📷 选择照片",
                                    command=self._select_photos,
                                    font=("微软雅黑", 10))
        self.btn_photos.pack(side="left", padx=5)

        self.lbl_info = tk.Label(toolbar, text="未选择模板 | 未选择照片",
                                 bg="#f0f0f0", font=("微软雅黑", 9), fg="#666")
        self.lbl_info.pack(side="right", padx=10)

        # === Main content: left (photos) + right (text input) ===
        main_frame = tk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # -- Left: photo thumbnails --
        left_frame = tk.Frame(main_frame, width=160, bg="white")
        left_frame.pack(side="left", fill="y")
        left_frame.pack_propagate(False)

        tk.Label(left_frame, text="照片预览", font=("微软雅黑", 9, "bold"),
                 bg="#e0e0e0").pack(fill="x", pady=(0, 2))

        self.photo_canvas = tk.Canvas(left_frame, bg="white", highlightthickness=0, width=150)
        photo_scroll = tk.Scrollbar(left_frame, orient="vertical", command=self.photo_canvas.yview)
        self.photo_frame = tk.Frame(self.photo_canvas, bg="white")
        self.photo_frame.bind("<Configure>",
            lambda e: self.photo_canvas.configure(scrollregion=self.photo_canvas.bbox("all")))
        self.photo_canvas.create_window((0, 0), window=self.photo_frame, anchor="nw", width=140)
        self.photo_canvas.configure(yscrollcommand=photo_scroll.set)
        self.photo_canvas.pack(side="left", fill="both", expand=True)
        photo_scroll.pack(side="right", fill="y")

        # -- Right: text input --
        right_frame = tk.Frame(main_frame)
        right_frame.pack(side="right", fill="both", expand=True, padx=(10, 0))

        tk.Label(right_frame, text="粘贴订货信息（每行或每段一个产品，程序自动提取）：",
                 font=("微软雅黑", 10, "bold"), anchor="w").pack(fill="x")

        self.txt_input = scrolledtext.ScrolledText(
            right_frame, font=("微软雅黑", 11), wrap="word",
            height=15, padx=8, pady=8)
        self.txt_input.pack(fill="both", expand=True)

        # Placeholder behavior
        self._placeholder_text = "例如：\n货号：A-001 单价：10元 一件100个\n\n货号：B-002 单价：15.5元 一件50个\n\n货号：C-003 单价：8元 一件200个"
        self._setup_placeholder()

        self.lbl_parse_info = tk.Label(right_frame, text="", font=("微软雅黑", 9), fg="#666", anchor="w")
        self.lbl_parse_info.pack(fill="x", pady=(3, 0))

        # === Bottom: generate ===
        bottom = tk.Frame(self.root, height=60, bg="#f0f0f0")
        bottom.pack(fill="x", padx=10, pady=8)
        bottom.pack_propagate(False)

        self.btn_generate = tk.Button(bottom, text="📄 生成订单",
                                      command=self._generate,
                                      font=("微软雅黑", 12, "bold"),
                                      bg="#4CAF50", fg="white",
                                      padx=30, pady=8,
                                      state="disabled")
        self.btn_generate.pack()

        self._update_parse_preview()

    def _select_template(self):
        path = filedialog.askopenfilename(
            title="选择订货单模板",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if not path:
            return
        try:
            self.template = read_template(path)
            self.template_path = path
            self._update_info()
        except Exception as e:
            messagebox.showerror("读取模板失败", str(e))

    def _select_photos(self):
        path = filedialog.askdirectory(title="选择照片文件夹")
        if not path:
            return
        self.photo_dir = path
        exts = (".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp")
        all_files = [f for f in os.listdir(path) if f.lower().endswith(exts)]
        all_files.sort()
        if not all_files:
            messagebox.showwarning("没有照片", "所选文件夹中没有找到图片文件")
            return
        self.photo_files = all_files
        self.thumb_refs.clear()
        self._update_info()
        self._load_thumbnails()

    def _load_thumbnails(self):
        for w in self.photo_frame.winfo_children():
            w.destroy()

        for fname in self.photo_files:
            frame = tk.Frame(self.photo_frame, bg="white")
            frame.pack(fill="x", pady=2)

            img_path = os.path.join(self.photo_dir, fname)
            try:
                pil = Image.open(img_path)
                pil.thumbnail((130, 80), Image.LANCZOS)
                tk_img = ImageTk.PhotoImage(pil)
                self.thumb_refs.append(tk_img)
                lbl = tk.Label(frame, image=tk_img, bg="white")
                lbl.pack()
            except Exception:
                tk.Label(frame, text="[图片]", bg="white",
                         font=("微软雅黑", 8), fg="#999").pack()

            tk.Label(frame, text=fname[:18], bg="white",
                     font=("微软雅黑", 7), fg="#aaa").pack()

    def _update_info(self):
        parts = []
        if self.template:
            parts.append(f"模板: {os.path.basename(self.template_path)} ({self.template.col_count}列)")
        else:
            parts.append("默认格式")
        if self.photo_files:
            parts.append(f"照片: {len(self.photo_files)}张")
        else:
            parts.append("未选照片")
        self.lbl_info.config(text=" | ".join(parts))
        self.btn_generate.config(state="normal" if self.photo_files else "disabled")

    def _setup_placeholder(self):
        self.txt_input.insert("1.0", self._placeholder_text)
        self.txt_input.config(fg="gray")

        def on_focus_in(event):
            if self.txt_input.get("1.0", "end-1c").strip() == self._placeholder_text.strip():
                self.txt_input.delete("1.0", "end")
                self.txt_input.config(fg="black")

        def on_focus_out(event):
            if not self.txt_input.get("1.0", "end-1c").strip():
                self.txt_input.insert("1.0", self._placeholder_text)
                self.txt_input.config(fg="gray")

        self.txt_input.bind("<FocusIn>", on_focus_in)
        self.txt_input.bind("<FocusOut>", on_focus_out)

    def _update_parse_preview(self, event=None):
        text = self.txt_input.get("1.0", "end-1c").strip()
        if not text or text == self._placeholder_text.strip():
            self.lbl_parse_info.config(text="")
            return
        parsed = parse_product_text(text)
        if parsed:
            keys = list(parsed[0].keys())
            self.lbl_parse_info.config(
                text=f"识别到 {len(parsed)} 个产品，字段：{' / '.join(keys)}"
            )

    def _generate(self):
        if not self.photo_files or not self.photo_dir:
            messagebox.showwarning("缺少照片", "请先选择照片文件夹")
            return

        text = self.txt_input.get("1.0", "end-1c").strip()
        if not text or text == self._placeholder_text.strip():
            messagebox.showwarning("缺少信息", "请粘贴订货信息")
            return

        parsed = parse_product_text(text)
        if not parsed:
            messagebox.showwarning("解析失败", "未能从文本中识别出产品信息\n请参考示例格式：货号：xxx 单价：xx元")
            return

        if len(parsed) != len(self.photo_files):
            ok = messagebox.askyesno(
                "数量不匹配",
                f"照片有 {len(self.photo_files)} 张，但识别到 {len(parsed)} 个产品。\n"
                f"继续的话会按 {min(len(parsed), len(self.photo_files))} 个生成。\n\n确定继续？"
            )
            if not ok:
                return

        # Merge parsed data with photo filenames
        data = []
        for i, product in enumerate(parsed):
            if i >= len(self.photo_files):
                break
            row = {"filename": self.photo_files[i]}
            row.update(product)
            # Ensure defaults for missing fields
            row.setdefault("件数", 1)
            row.setdefault("每件数量", 240)
            data.append(row)

        tmpl = self.template if self.template else get_default_format()

        desktop = os.path.join(os.path.expanduser("~"), "Desktop")
        today = date.today()
        out_name = f"订货单 {today.year}年{today.month}月{today.day}日.xlsx"
        out_path = os.path.join(desktop, out_name)

        try:
            generate(tmpl, data, self.photo_dir, out_path)
            count = len(data)
            msg = f"成功生成 {count} 个产品\n已保存到桌面：{out_name}"
            # Detail summary
            details = "\n".join(
                f"  {d.get('品名','?')}  ¥{d.get('单价',0)}  {d.get('件数',1)}件×{d.get('每件数量',240)}个"
                for d in data
            )
            messagebox.showinfo("生成成功", f"{msg}\n\n{details}")
        except Exception as e:
            messagebox.showerror("生成失败", str(e))

    def run(self):
        self.root.mainloop()
