import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

class ExcelQueryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel多Sheet查询工具")
        
        # 初始化变量
        self.file_path = ""
        self.all_data = {}
        
        # 创建界面组件
        self.create_widgets()
        
    def create_widgets(self):
        """创建GUI界面组件"""
        # 文件选择区域
        file_frame = tk.LabelFrame(self.root, text="文件操作", padx=5, pady=5)
        file_frame.pack(fill="x", padx=10, pady=5)
        
        self.file_entry = tk.Entry(file_frame, width=50)
        self.file_entry.pack(side="left", padx=5)
        
        browse_btn = tk.Button(file_frame, text="选择Excel文件", command=self.load_file)
        browse_btn.pack(side="left", padx=5)
        
        load_all_btn = tk.Button(file_frame, text="加载Sheet名称", command=self.display_all_sheets)
        load_all_btn.pack(side="left", padx=5)
        
        # 搜索区域
        search_frame = tk.LabelFrame(self.root, text="内容搜索", padx=5, pady=5)
        search_frame.pack(fill="x", padx=10, pady=5)
        
        self.search_entry = tk.Entry(search_frame, width=50)
        self.search_entry.pack(side="left", padx=5)
        
        search_btn = tk.Button(search_frame, text="搜索", command=self.search_content)
        search_btn.pack(side="left", padx=5)
        
        export_btn = tk.Button(search_frame, text="导出结果", command=self.export_results)
        export_btn.pack(side="left", padx=5)
        
        # 结果显示区域
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=5)
        
        # 搜索结果显示标签页(放在前面)
        self.search_results_tab = tk.Frame(self.notebook)
        self.notebook.add(self.search_results_tab, text="搜索结果")
        
        self.search_results_text = tk.Text(self.search_results_tab, wrap="none")
        self.search_results_scroll_y = tk.Scrollbar(self.search_results_tab, orient="vertical", command=self.search_results_text.yview)
        self.search_results_scroll_x = tk.Scrollbar(self.search_results_tab, orient="horizontal", command=self.search_results_text.xview)
        self.search_results_text.configure(yscrollcommand=self.search_results_scroll_y.set, xscrollcommand=self.search_results_scroll_x.set)
        
        self.search_results_text.grid(row=0, column=0, sticky="nsew")
        self.search_results_scroll_y.grid(row=0, column=1, sticky="ns")
        self.search_results_scroll_x.grid(row=1, column=0, sticky="ew")
        
        # Sheet名称显示标签页
        self.all_sheets_tab = tk.Frame(self.notebook)
        self.notebook.add(self.all_sheets_tab, text="所有Sheet名称")
        
        self.all_sheets_text = tk.Text(self.all_sheets_tab, wrap="none")
        self.all_sheets_scroll_y = tk.Scrollbar(self.all_sheets_tab, orient="vertical", command=self.all_sheets_text.yview)
        self.all_sheets_scroll_x = tk.Scrollbar(self.all_sheets_tab, orient="horizontal", command=self.all_sheets_text.xview)
        self.all_sheets_text.configure(yscrollcommand=self.all_sheets_scroll_y.set, xscrollcommand=self.all_sheets_scroll_x.set)
        
        self.all_sheets_text.grid(row=0, column=0, sticky="nsew")
        self.all_sheets_scroll_y.grid(row=0, column=1, sticky="ns")
        self.all_sheets_scroll_x.grid(row=1, column=0, sticky="ew")
        
        # 配置网格权重
        self.all_sheets_tab.grid_rowconfigure(0, weight=1)
        self.all_sheets_tab.grid_columnconfigure(0, weight=1)
        self.search_results_tab.grid_rowconfigure(0, weight=1)
        self.search_results_tab.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        
    def load_file(self):
        """加载Excel文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if file_path:
            self.file_path = file_path
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)
            self.all_data = self.read_all_sheets(file_path)
            
    def read_all_sheets(self, file_path):
        """读取Excel所有sheet内容"""
        try:
            xls = pd.ExcelFile(file_path)
            return {sheet_name: pd.read_excel(xls, sheet_name=sheet_name) for sheet_name in xls.sheet_names}
        except Exception as e:
            messagebox.showerror("错误", f"读取文件失败: {str(e)}")
            return {}
            
    def display_all_sheets(self):
        """显示所有sheet名称"""
        if not self.all_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        self.all_sheets_text.delete(1.0, tk.END)
        
        for sheet_name in self.all_data.keys():
            self.all_sheets_text.insert(tk.END, f"{sheet_name}\n")
                
    def search_content(self):
        """搜索内容并显示结果"""
        if not self.all_data:
            messagebox.showwarning("警告", "请先选择Excel文件")
            return
            
        search_term = self.search_entry.get().strip()
        if not search_term:
            messagebox.showwarning("警告", "请输入搜索内容")
            return
            
        self.search_results_text.delete(1.0, tk.END)
        found_results = False
        
        for sheet_name, df in self.all_data.items():
            # 搜索所有列
            mask = df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, regex=False)).any(axis=1)
            results = df[mask]
            
            if not results.empty:
                found_results = True
                # 添加sheet名称标签
                self.search_results_text.insert(tk.END, f"\n\n=== Sheet: {sheet_name} ===\n\n")
                
                # 添加表头
                headers = "\t".join(df.columns) + "\n"
                self.search_results_text.insert(tk.END, headers)
                
                # 添加匹配行
                for _, row in results.iterrows():
                    row_str = "\t".join([str(cell) for cell in row]) + "\n"
                    self.search_results_text.insert(tk.END, row_str)
                    
        if not found_results:
            self.search_results_text.insert(tk.END, f"未找到包含 '{search_term}' 的内容")
    
    def export_results(self):
        """导出搜索结果到Excel文件，保持与界面相同的格式"""
        if not hasattr(self, 'search_results_text') or not self.search_results_text.get(1.0, tk.END).strip():
            messagebox.showwarning("警告", "没有可导出的搜索结果")
            return
            
        # 获取当前时间作为文件名
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile=f"搜索结果_{timestamp}.xlsx"
        )
        
        if not save_path:
            return
            
        try:
            # 创建Excel写入对象
            writer = pd.ExcelWriter(save_path, engine='xlsxwriter')
            current_sheet = ""
            current_data = []
            headers = []
            
            for line in self.search_results_text.get(1.0, tk.END).split('\n'):
                if line.startswith("=== Sheet: ") and line.endswith(" ==="):
                    # 保存前一个Sheet的数据
                    if current_sheet and current_data:
                        df = pd.DataFrame(current_data)
                        df.to_excel(writer, sheet_name=current_sheet, index=False)
                    
                    # 开始新Sheet
                    current_sheet = line[11:-4]
                    current_data = []
                    headers = []
                elif line.strip() and not line.startswith("=== "):
                    if not headers and current_sheet:
                        # 第一行是表头
                        headers = line.split('\t')
                    elif headers and current_sheet:
                        # 数据行
                        values = line.split('\t')
                        if len(values) == len(headers):
                            current_data.append(dict(zip(headers, values)))
            
            # 保存最后一个Sheet的数据
            if current_sheet and current_data:
                df = pd.DataFrame(current_data)
                df.to_excel(writer, sheet_name=current_sheet, index=False)
            
            # 保存Excel文件
            writer.close()
            messagebox.showinfo("成功", f"搜索结果已导出到:\n{save_path}")
        except Exception as e:
            messagebox.showerror("错误", f"导出失败: {str(e)}")

def main():
    root = tk.Tk()
    root.geometry("800x600")
    app = ExcelQueryApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
