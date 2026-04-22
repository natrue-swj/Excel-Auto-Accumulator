import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

class ExcelMultiAttachProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("多附表时长累加工具（最终版）")
        self.root.geometry("700x600")
        self.root.resizable(False, False)

        # 主表数据
        self.main_file = tk.StringVar()
        self.main_df = None
        self.main_columns = []

        # 附表列表 [ {"file":路径, "column":列名} ]
        self.attach_list = []

        # ========== 界面 ==========
        # 主表
        ttk.Label(root, text="主表 Excel：").place(x=20, y=20)
        ttk.Entry(root, textvariable=self.main_file, width=50, state="readonly").place(x=120, y=20)
        ttk.Button(root, text="选择主表", command=self.load_main).place(x=550, y=18)

        # 附表区域标题
        ttk.Label(root, text="↓ 可添加多个附表，每个附表选择要累加的列", font=("微软雅黑", 10, "bold")).place(x=20, y=70)

        # 附表容器
        self.attach_frame = ttk.Frame(root)
        self.attach_frame.place(x=20, y=100, width=650, height=350)

        # 添加附表按钮
        self.add_btn = ttk.Button(root, text="➕ 添加附表", command=self.add_attach_row, state=tk.DISABLED)
        self.add_btn.place(x=20, y=460)

        # 执行按钮
        self.run_btn = ttk.Button(root, text="✅ 开始处理并导出总表", command=self.run_all, state=tk.DISABLED)
        self.run_btn.place(x=200, y=460, width=200, height=40)

        # 状态
        self.status_label = ttk.Label(root, text="状态：请选择主表", font=("微软雅黑", 10))
        self.status_label.place(x=20, y=520)

    # 加载主表
    def load_main(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return

        try:
            df = pd.read_excel(path, engine="openpyxl")

            if "姓名" not in df.columns:
                messagebox.showerror("错误", "主表必须包含【姓名】列")
                return

            self.main_file.set(path)
            self.main_df = df.copy()
            self.main_columns = [c for c in df.columns if c != "姓名"]

            self.status_label.config(text="状态：主表加载成功")
            self.add_btn.config(state=tk.NORMAL)
            self.refresh_all_attach_columns()

        except Exception as e:
            messagebox.showerror("错误", f"读取失败：{str(e)}")

    # 添加一行附表
    def add_attach_row(self):
        row_idx = len(self.attach_list)
        var_file = tk.StringVar()
        var_col = tk.StringVar()

        frame = ttk.Frame(self.attach_frame)
        frame.pack(fill=tk.X, pady=3)

        ttk.Label(frame, text=f"附表 {row_idx+1}：").pack(side=tk.LEFT, padx=5)
        entry = ttk.Entry(frame, textvariable=var_file, width=30, state="readonly")
        entry.pack(side=tk.LEFT, padx=5)

        ttk.Button(frame, text="选择文件", command=lambda: self.select_attach(var_file)).pack(side=tk.LEFT, padx=5)
        combo = ttk.Combobox(frame, textvariable=var_col, values=self.main_columns, width=12, state="readonly")
        combo.pack(side=tk.LEFT, padx=5)

        if self.main_columns:
            combo.current(0)

        self.attach_list.append({
            "frame": frame,
            "file": var_file,
            "column": var_col
        })

        self.check_run_button()

    # 选择附表
    def select_attach(self, var):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if f:
            var.set(f)
            self.check_run_button()

    # 刷新所有附表的列选择
    def refresh_all_attach_columns(self):
        for item in self.attach_list:
            combo = item["frame"].winfo_children()[4]
            combo["values"] = self.main_columns
            if self.main_columns:
                combo.current(0)

    # 检查是否可以运行
    def check_run_button(self):
        if self.main_df is None:
            self.run_btn.config(state=tk.DISABLED)
            return

        has_all = True
        for item in self.attach_list:
            if not item["file"].get() or not item["column"].get():
                has_all = False

        if len(self.attach_list) > 0 and has_all:
            self.run_btn.config(state=tk.NORMAL)
        else:
            self.run_btn.config(state=tk.DISABLED)

    # 执行所有累加
    def run_all(self):
        try:
            self.status_label.config(text="状态：处理中...")
            self.root.update()

            # 复制主表
            result = self.main_df.copy()

            # 所有数字列先转数值 + 空值变0
            for col in self.main_columns:
                result[col] = pd.to_numeric(result[col], errors="coerce").fillna(0)

            # 逐个处理附表
            for item in self.attach_list:
                file = item["file"].get()
                target_col = item["column"].get()

                add_df = pd.read_excel(file, engine="openpyxl")

                if "姓名" not in add_df.columns or "时长" not in add_df.columns:
                    messagebox.showerror("错误", f"附表缺少 姓名/时长 列")
                    return

                add_df["时长"] = pd.to_numeric(add_df["时长"], errors="coerce").fillna(0)

                # 按姓名累加
                for _, r in add_df.iterrows():
                    name = str(r["姓名"]).strip()
                    val = r["时长"]
                    mask = result["姓名"].astype(str).str.strip() == name
                    if mask.any():
                        result.loc[mask, target_col] += val

            # ========== 关键：把所有 0 变成空白 ==========
            for col in self.main_columns:
                result[col] = result[col].replace(0, "")

            # 保存
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel", "*.xlsx")],
                initialfile="总累加结果.xlsx"
            )

            if not save_path:
                self.status_label.config(text="状态：已取消")
                return

            result.to_excel(save_path, index=False, engine="openpyxl")
            self.status_label.config(text="状态：导出完成！0值已自动清空")
            messagebox.showinfo("成功", f"总表已导出！\n{save_path}")

        except Exception as e:
            messagebox.showerror("失败", f"错误：{str(e)}")
            self.status_label.config(text="状态：出错")

if __name__ == "__main__":
    root = tk.Tk()
    ExcelMultiAttachProcessor(root)
    root.mainloop()