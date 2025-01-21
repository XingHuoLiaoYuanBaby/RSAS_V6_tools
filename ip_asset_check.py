import ipaddress
import csv
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path

class IPRangeCounter(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # 设置窗口
        self.title("IP地址段统计工具")
        self.geometry("500x600")
        
        # 使窗口居中显示
        self.center_window()
        
        # 创建主框架
        self.main_frame = ttk.Frame(self, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 创建样式
        style = ttk.Style()
        style.configure('Custom.TButton', padding=5)
        
        # 输入文件选择
        self.input_frame = ttk.LabelFrame(self.main_frame, text="输入文件", padding="5")
        self.input_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.input_path = tk.StringVar()
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path, width=50)
        self.input_entry.grid(row=0, column=0, padx=5)
        
        self.input_button = ttk.Button(self.input_frame, text="选择文件", 
                                     command=self.select_input_file, style='Custom.TButton')
        self.input_button.grid(row=0, column=1, padx=5)
        
        # 输出文件选择
        self.output_frame = ttk.LabelFrame(self.main_frame, text="输出文件", padding="5")
        self.output_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.output_path = tk.StringVar()
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path, width=50)
        self.output_entry.grid(row=0, column=0, padx=5)
        
        self.output_button = ttk.Button(self.output_frame, text="选择文件", 
                                      command=self.select_output_file, style='Custom.TButton')
        self.output_button.grid(row=0, column=1, padx=5)
        
        # 进度显示
        self.progress_frame = ttk.LabelFrame(self.main_frame, text="处理进度", padding="5")
        self.progress_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        self.progress_var = tk.StringVar(value="等待开始...")
        self.progress_label = ttk.Label(self.progress_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        self.progressbar = ttk.Progressbar(self.progress_frame, mode='determinate')
        self.progressbar.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        # 结果显示
        self.result_frame = ttk.LabelFrame(self.main_frame, text="处理结果", padding="5")
        self.result_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        
        # 创建Treeview来显示结果
        self.tree = ttk.Treeview(self.result_frame, columns=('Network', 'Count'), 
                                show='headings', height=10)
        self.tree.heading('Network', text='网段')
        self.tree.heading('Count', text='IP数量')
        self.tree.column('Network', width=300)
        self.tree.column('Count', width=100)
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.result_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 开始处理按钮
        self.process_button = ttk.Button(self.main_frame, text="开始处理", 
                                       command=self.process_file, style='Custom.TButton')
        self.process_button.grid(row=4, column=0, columnspan=2, pady=10)

    def center_window(self):
        """
        使窗口在屏幕中居中显示
        """
        # 获取屏幕宽度和高度
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        
        # 获取窗口宽度和高度
        window_width = 500
        window_height = 600
        
        # 计算居中位置
        center_x = int((screen_width - window_width) / 2)
        center_y = int((screen_height - window_height) / 2)
        
        # 设置窗口位置
        self.geometry(f"{window_width}x{window_height}+{center_x}+{center_y}")

    def select_input_file(self):
        """
        选择输入文件
        默认打开当前目录，支持txt文件
        自动设置输出文件名为同目录下的ip_ranges.csv
        """
        filename = filedialog.askopenfilename(
            title="选择输入文件",
            initialdir=Path.cwd(),  # 设置初始目录为当前目录
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if filename:
            self.input_path.set(filename)
            # 自动设置输出文件名
            output_path = Path(filename).parent / "ip段统计.csv"
            self.output_path.set(str(output_path))

    def select_output_file(self):
        """
        选择输出文件
        默认打开当前目录，默认文件名为ip_ranges.csv
        """
        filename = filedialog.asksaveasfilename(
            title="选择保存位置",
            initialdir=Path.cwd(),  # 设置初始目录为当前目录
            initialfile="ip_ranges.csv",  # 设置默认文件名
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.output_path.set(filename)

    def read_ips_from_file(self, file_path):
        """
        读取IP列表并去重
        Args:
            file_path: IP文件路径
        Returns:
            list: 去重后的IP列表
        """
        try:
            with open(file_path, 'r') as file:
                # 读取所有行，去除空白字符，并转换为集合去重
                ips = set(ip.strip() for ip in file.readlines() if ip.strip())
                
                # 转回列表并排序，方便查看
                unique_ips = sorted(list(ips))
                
                # 显示去重信息
                original_count = sum(1 for line in open(file_path) if line.strip())
                removed_count = original_count - len(unique_ips)
                if removed_count > 0:
                    messagebox.showinfo("去重结果", 
                        f"原始IP数量: {original_count}\n"
                        f"去重后数量: {len(unique_ips)}\n"
                        f"重复IP数量: {removed_count}")
                
                return unique_ips
                
        except Exception as e:
            messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
            return []

    def count_ip_ranges(self, ips):
        """
        统计IP地址段，同时统计/24和/16网段
        """
        ip_ranges_24 = {}  # 存储/24网段统计
        ip_ranges_16 = {}  # 存储/16网段统计
        total = len(ips)
        self.progressbar['maximum'] = total
        
        for i, ip in enumerate(ips, 1):
            try:
                # 统计/24网段
                network_24 = ipaddress.ip_network(f"{ip}/24", strict=False)
                network_24_str = str(network_24)
                if network_24_str in ip_ranges_24:
                    ip_ranges_24[network_24_str] += 1
                else:
                    ip_ranges_24[network_24_str] = 1
                
                # 统计/16网段
                network_16 = ipaddress.ip_network(f"{ip}/16", strict=False)
                network_16_str = str(network_16)
                if network_16_str in ip_ranges_16:
                    ip_ranges_16[network_16_str] += 1
                else:
                    ip_ranges_16[network_16_str] = 1
                
                # 更新进度
                self.progressbar['value'] = i
                self.progress_var.set(f"处理进度: {i}/{total}")
                self.update_idletasks()
                
            except ValueError:
                messagebox.showwarning("警告", f"发现无效IP地址: {ip}")
        
        return ip_ranges_24, ip_ranges_16

    def write_to_csv(self, ip_ranges_24, ip_ranges_16, output_file):
        """
        将/24和/16网段的统计结果写入CSV文件
        """
        with open(output_file, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)
            # 写入/24网段统计
            writer.writerow(['C段统计（/24）', '存活ip数量'])
            for network, count in ip_ranges_24.items():
                writer.writerow([network, count])
            
            # 添加空行
            writer.writerow([])
            writer.writerow([])
            
            # 写入/16网段统计
            writer.writerow(['B段统计（/16）', '存活ip数量'])
            for network, count in ip_ranges_16.items():
                writer.writerow([network, count])

    def update_result_tree(self, ip_ranges_24, ip_ranges_16):
        """
        更新树形视图，显示/24和/16网段的统计结果
        """
        # 清除现有内容
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加/24网段数据
        c_section = self.tree.insert('', 'end', text='C段统计（/24）', open=True)
        for network, count in ip_ranges_24.items():
            self.tree.insert(c_section, 'end', values=(network, count))
        
        # 添加/16网段数据
        b_section = self.tree.insert('', 'end', text='B段统计（/16）', open=True)
        for network, count in ip_ranges_16.items():
            self.tree.insert(b_section, 'end', values=(network, count))

    def process_file(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        
        if not input_file or not output_file:
            messagebox.showerror("错误", "请选择输入和输出文件")
            return
        
        try:
            # 重置进度条
            self.progressbar['value'] = 0
            self.progress_var.set("正在读取文件...")
            self.update_idletasks()
            
            # 处理文件
            ips = self.read_ips_from_file(input_file)
            ip_ranges_24, ip_ranges_16 = self.count_ip_ranges(ips)
            self.write_to_csv(ip_ranges_24, ip_ranges_16, output_file)
            
            # 更新显示
            self.update_result_tree(ip_ranges_24, ip_ranges_16)
            self.progress_var.set("处理完成!")
            messagebox.showinfo("成功", f"结果已保存到: {output_file}")
            
        except Exception as e:
            messagebox.showerror("错误", f"处理过程中出错: {str(e)}")

if __name__ == "__main__":
    app = IPRangeCounter()
    app.mainloop()
