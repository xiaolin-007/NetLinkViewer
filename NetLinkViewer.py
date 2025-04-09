import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import socket
import psutil
import requests
import csv
from datetime import datetime
import threading

class SortedNetworkViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("NetLinkViewer")
        self.root.geometry("1200x700")
        
        # 排序相关变量
        self.sort_column = None
        self.sort_reverse = False
        
        # 创建主界面
        self.create_widgets()
        
        # 初始化数据
        self.connections = []
        self.ip_location_cache = {}
        self.is_loading = False
        
        # 首次加载数据
        self.refresh_data()
    
    def create_widgets(self):
        # 顶部按钮区域
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        # 状态栏放在左侧
        self.status_var = tk.StringVar()
        self.status_var.set("就绪")
        status_bar = ttk.Label(top_frame, textvariable=self.status_var, relief=tk.SUNKEN, width=50)
        status_bar.pack(side=tk.LEFT, padx=5)
        
        # 动态提醒标签
        self.notification_var = tk.StringVar()
        self.notification_label = ttk.Label(top_frame, textvariable=self.notification_var, foreground="blue")
        self.notification_label.pack(side=tk.LEFT, padx=10, expand=True)
        
        # 按钮放在右侧
        button_frame = ttk.Frame(top_frame)
        button_frame.pack(side=tk.RIGHT)
        
        self.refresh_btn = ttk.Button(button_frame, text="刷新", command=self.refresh_data)
        self.refresh_btn.pack(side=tk.LEFT, padx=5)
        
        self.export_btn = ttk.Button(button_frame, text="导出CSV", command=self.export_to_csv)
        self.export_btn.pack(side=tk.LEFT, padx=5)
        
        # 连接表格
        self.tree = ttk.Treeview(self.root, columns=(
            'pid', 'name', 'protocol', 'local_ip', 'local_port', 
            'remote_ip', 'remote_port', 'status', 'location'
        ), show='headings')
        
        # 设置列宽和标题
        columns = {
            'pid': ('PID', 50),
            'name': ('程序名称', 150),
            'protocol': ('协议', 60),
            'local_ip': ('本地IP', 120),
            'local_port': ('本地端口', 80),
            'remote_ip': ('远程IP', 120),
            'remote_port': ('远程端口', 80),
            'status': ('状态', 80),
            'location': ('归属地', 200)
        }
        
        for col, (text, width) in columns.items():
            self.tree.heading(col, text=text, command=lambda c=col: self.treeview_sort_column(c))
            self.tree.column(col, width=width, anchor=tk.CENTER)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(self.root, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # 布局
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    
    def treeview_sort_column(self, col):
        """点击列标题排序"""
        # 获取当前排序方向
        if self.sort_column == col:
            # 同一列点击，切换排序方向
            self.sort_reverse = not self.sort_reverse
        else:
            # 不同列点击，默认升序
            self.sort_column = col
            self.sort_reverse = False
        
        # 更新列标题显示排序方向
        for c in self.tree['columns']:
            if c == col:
                text = self.tree.heading(c)['text']
                if self.sort_reverse:
                    self.tree.heading(c, text=text + ' ↓')
                else:
                    self.tree.heading(c, text=text + ' ↑')
            else:
                # 移除其他列的排序指示
                text = self.tree.heading(c)['text']
                self.tree.heading(c, text=text.rstrip(' ↓↑'))
        
        # 对数据进行排序
        self.sort_data()
    
    def sort_data(self):
        """根据当前排序设置对数据进行排序"""
        if not self.sort_column or not self.connections:
            return
        
        # 获取排序键函数
        def get_key(item):
            value = item[self.sort_column]
            
            # 处理数字类型的列
            if self.sort_column in ('pid', 'local_port', 'remote_port'):
                try:
                    return int(value) if value else 0
                except ValueError:
                    return 0
            
            # 处理IP地址
            if self.sort_column in ('local_ip', 'remote_ip'):
                try:
                    return tuple(map(int, value.split('.'))) if value else (0, 0, 0, 0)
                except ValueError:
                    return (0, 0, 0, 0)
            
            # 默认按字符串处理
            return str(value).lower()
        
        # 执行排序
        self.connections.sort(key=get_key, reverse=self.sort_reverse)
        
        # 更新表格显示
        self.update_treeview()
    
    def refresh_data(self):
        if self.is_loading:
            return
            
        self.is_loading = True
        self.refresh_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        
        # 显示动态提醒
        self.show_loading_notification()
        
        # 在新线程中获取数据
        threading.Thread(target=self._get_network_data, daemon=True).start()
    
    def show_loading_notification(self):
        self.notification_var.set("正在获取网络连接数据，请稍候...")
        self.status_var.set("数据加载中...")
        self.root.update()
    
    def show_complete_notification(self):
        self.notification_var.set("网络连接数据更新已完成！")
        self.root.after(3000, lambda: self.notification_var.set(""))  # 3秒后自动清除
    
    def _get_network_data(self):
        try:
            # 清空现有数据
            self.connections = []
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # 获取所有网络连接
            connections = []
            for conn in psutil.net_connections(kind='inet'):
                if conn.status == 'NONE' and conn.type not in (socket.SOCK_STREAM, socket.SOCK_DGRAM):
                    continue
                
                # 获取进程信息
                try:
                    process = psutil.Process(conn.pid)
                    name = process.name()
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    name = "N/A"
                
                # 格式化连接信息
                protocol = 'TCP' if conn.type == socket.SOCK_STREAM else 'UDP'
                
                local_ip = conn.laddr.ip if conn.laddr else ""
                local_port = conn.laddr.port if conn.laddr else ""
                
                remote_ip = conn.raddr.ip if conn.raddr else ""
                remote_port = conn.raddr.port if conn.raddr else ""
                
                status = conn.status
                
                # 获取IP归属地
                location = self.get_ip_location(remote_ip) if remote_ip else ""
                
                connections.append({
                    'pid': str(conn.pid),
                    'name': name,
                    'protocol': protocol,
                    'local_ip': local_ip,
                    'local_port': str(local_port),
                    'remote_ip': remote_ip,
                    'remote_port': str(remote_port),
                    'status': status,
                    'location': location
                })
            
            # 更新数据
            self.connections = connections
            
            # 如果有排序设置，应用排序
            if self.sort_column:
                self.sort_data()
            else:
                self.update_treeview()
            
            # 显示完成通知
            self.show_complete_notification()
            self.status_var.set(f"数据已更新 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            
        except Exception as e:
            self.notification_var.set(f"错误: {str(e)}")
            self.status_var.set("数据加载失败")
            messagebox.showerror("错误", f"获取网络连接数据时出错:\n{str(e)}")
        finally:
            self.is_loading = False
            self.refresh_btn.config(state=tk.NORMAL)
            self.export_btn.config(state=tk.NORMAL)
    
    def update_treeview(self):
        # 先清空表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 添加新数据
        for conn in self.connections:
            self.tree.insert('', tk.END, values=(
                conn['pid'],
                conn['name'],
                conn['protocol'],
                conn['local_ip'],
                conn['local_port'],
                conn['remote_ip'],
                conn['remote_port'],
                conn['status'],
                conn['location']
            ))
    
    def get_ip_location(self, ip):
        if not ip or ip.startswith(('127.', '192.168.', '10.', '172.')):
            return "局域网"
        
        # 使用缓存
        if ip in self.ip_location_cache:
            return self.ip_location_cache[ip]
        
        try:
            # 使用ip-api.com查询IP归属地
            response = requests.get(f"http://ip-api.com/json/{ip}?lang=zh-CN", timeout=3)
            data = response.json()
            
            if data['status'] == 'success':
                location = f"{data['country']} {data['regionName']} {data['city']} ({data['isp']})"
                self.ip_location_cache[ip] = location
                return location
            return "未知"
        except Exception:
            return "查询失败"
    
    def export_to_csv(self):
        if not self.connections:
            messagebox.showwarning("警告", "没有可导出的数据")
            return
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV 文件", "*.csv")],
            title="保存为CSV文件"
        )
        
        if not file_path:
            return
        
        try:
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.writer(f)
                # 写入标题行
                writer.writerow([
                    'PID', '程序名称', '协议', '本地IP', '本地端口',
                    '远程IP', '远程端口', '状态', '归属地'
                ])
                # 写入数据
                for conn in self.connections:
                    writer.writerow([
                        conn['pid'], conn['name'], conn['protocol'],
                        conn['local_ip'], conn['local_port'],
                        conn['remote_ip'], conn['remote_port'],
                        conn['status'], conn['location']
                    ])
            
            self.notification_var.set(f"数据已成功导出到: {file_path}")
            self.root.after(3000, lambda: self.notification_var.set(""))
        except Exception as e:
            messagebox.showerror("错误", f"导出CSV文件时出错:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = SortedNetworkViewer(root)
    root.mainloop()
