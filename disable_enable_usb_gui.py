import subprocess
import re
import ctypes
import sys
import os
import time
import json
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading

# 设置Windows任务栏图标
try:
    from ctypes import windll
    windll.shell32.SetCurrentProcessExplicitAppUserModelID("USB_Device_Controller")
except:
    pass

def is_admin():
    """检查脚本是否以管理员权限运行"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def elevate_privileges():
    """提升权限到管理员"""
    # 获取当前脚本的完整路径
    script = os.path.abspath(sys.argv[0])
    # 使用 Python 解释器重新运行该脚本
    params = ' '.join([f'"{item}"' for item in sys.argv[1:]])
    # 使用 ShellExecute 以管理员身份启动程序
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)

def get_all_devices():
    """获取所有设备的列表"""
    cmd = 'pnputil /enum-devices'
    try:
        result = subprocess.check_output(cmd, shell=True, text=True)
        return result
    except subprocess.CalledProcessError:
        return ""

def load_config():
    """从config.json加载配置"""
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    
    # 默认配置
    default_config = {
        "device_id": "USB\\VID_174C&PID_1153",
        "use_full_id": False,
        "full_device_id": "USB\\VID_174C&PID_1153\\MSFT3023456789013B"
    }
    
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
            return config
        except Exception as e:
            print(f"读取配置文件失败: {e}")
    else:
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
        except Exception as e:
            print(f"创建配置文件失败: {e}")
    
    return default_config

def save_config(config):
    """保存配置到config.json"""
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
    try:
        with open(config_path, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"保存配置失败: {e}")
        return False

def update_config_with_device_id(device_id):
    """根据找到的设备ID更新配置文件"""
    config = load_config()
    
    # 更新配置文件中的设备ID
    config["use_full_id"] = True
    config["full_device_id"] = device_id
    
    # 提取设备ID的部分ID (VID和PID部分)
    vid_pid_match = re.search(r'(USB\\VID_[0-9A-F]{4}.*?PID_[0-9A-F]{4})', device_id, re.IGNORECASE)
    if vid_pid_match:
        partial_id = vid_pid_match.group(1)
        config["device_id"] = partial_id
    
    # 保存配置文件
    save_config(config)
    
    return config

def find_devices_by_partial_id(device_id):
    """根据部分ID找到设备的完整ID列表"""
    all_devices = get_all_devices()
    
    # 提取VID和PID部分
    vid_pid_match = re.search(r'VID_([0-9A-F]{4}).*?PID_([0-9A-F]{4})', device_id, re.IGNORECASE)
    if not vid_pid_match:
        return []
    
    vid = vid_pid_match.group(1)
    pid = vid_pid_match.group(2)
    
    # 查找设备块
    device_blocks = re.split(r'(?:实例 ID|Instance ID):', all_devices)
    
    # 存储匹配的设备ID
    matched_devices = []
    
    # 搜索包含VID和PID的设备
    pattern = rf'USB.*?VID_{vid}.*?PID_{pid}'
    for block in device_blocks:
        device_match = re.search(pattern, block, re.IGNORECASE | re.DOTALL)
        if device_match:
            # 尝试找到完整的设备ID
            id_match = re.search(r'(USB\\VID_.*?)(?:\r|\n|$)', block, re.IGNORECASE | re.DOTALL)
            if id_match:
                device_full_id = id_match.group(1).strip()
                matched_devices.append(device_full_id)
    
    return matched_devices

def get_device_status(device_id):
    """获取指定设备ID的设备状态，返回是否被禁用"""
    # 确保设备ID格式正确（去除可能的引号和空格）
    device_id = device_id.strip('"\'').strip()
    
    # 使用正确的参数格式尝试获取设备状态
    commands = [
        f'pnputil /enum-devices /instanceid "{device_id}"',
        f'pnputil /enum-devices /deviceid "{device_id}"'
    ]
    
    for cmd in commands:
        try:
            result = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.STDOUT)
            
            # 检查设备是否存在
            if "找不到" in result or "not found" in result.lower():
                continue
                
            # 更精确地检查设备状态
            status_lines = [line for line in result.split('\n') if "状态" in line or "Status" in line]
            if status_lines:
                status_line = status_lines[0].lower()
                if "已禁用" in status_line or "disabled" in status_line:
                    return True  # 设备已禁用
                elif "已启用" in status_line or "enabled" in status_line or "started" in status_line or "已启动" in status_line:
                    return False  # 设备已启用
            
            # 尝试查找其他状态信息
            if "已禁用" in result or "disabled" in result.lower():
                return True  # 设备已禁用
            elif "正常工作" in result or "working properly" in result.lower():
                return False  # 设备已启用
            
            # 如果没有找到明确的状态信息，假设设备已启用
            return False
            
        except subprocess.CalledProcessError:
            continue
    
    # 如果上述方法都失败，尝试使用devcon命令（如果可用）
    try:
        cmd = f'devcon status "@{device_id}"'
        result = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.STDOUT)
        
        if "已禁用" in result or "disabled" in result.lower():
            return True
        else:
            return False
    except subprocess.CalledProcessError:
        pass
    
    return False  # 默认为已启用

def disable_device(device_id):
    """禁用指定设备ID的设备"""
    device_id = device_id.strip('"\'').strip()
    
    # 尝试不同的命令格式
    commands = [
        f'pnputil /disable-device /instanceid "{device_id}"',
        f'devcon disable "@{device_id}"'
    ]
    
    for cmd in commands:
        try:
            result = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.STDOUT)
            return True
        except subprocess.CalledProcessError:
            continue
    
    return False

def enable_device(device_id):
    """启用指定设备ID的设备"""
    device_id = device_id.strip('"\'').strip()
    
    # 尝试不同的命令格式
    commands = [
        f'pnputil /enable-device /instanceid "{device_id}"',
        f'devcon enable "@{device_id}"'
    ]
    
    for cmd in commands:
        try:
            result = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.STDOUT)
            return True
        except subprocess.CalledProcessError:
            continue
    
    return False

def device_exists(device_id):
    """检查设备是否存在"""
    device_id = device_id.strip('"\'').strip()
    
    # 尝试不同的命令格式
    commands = [
        f'pnputil /enum-devices /instanceid "{device_id}"',
        f'pnputil /enum-devices /deviceid "{device_id}"'
    ]
    
    for cmd in commands:
        try:
            result = subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.PIPE)
            # 如果命令执行成功并且结果不包含"找不到"
            if "找不到" not in result and "not found" not in result.lower():
                return True
        except subprocess.CalledProcessError:
            continue
    
    # 尝试使用devcon
    try:
        cmd = f'devcon status "@{device_id}"'
        subprocess.check_output(cmd, shell=True, text=True, stderr=subprocess.PIPE)
        return True
    except subprocess.CalledProcessError:
        pass
    
    return False

def list_all_usb_devices():
    """列出所有USB设备以帮助用户找到正确的设备ID"""
    devices_info = []
    
    try:
        # 使用pnputil列出所有USB设备
        cmd = 'pnputil /enum-devices /deviceid "USB*" /connected'
        result = subprocess.check_output(cmd, shell=True, text=True)
        
        # 尝试提取所有USB设备ID
        device_blocks = re.split(r'(?:实例 ID|Instance ID):', result)
        
        for block in device_blocks:
            if not block.strip():
                continue
                
            # 提取设备ID
            id_match = re.search(r'(USB\\VID_.*?)(?:\r|\n|$)', block, re.IGNORECASE | re.DOTALL)
            if id_match:
                device_id = id_match.group(1).strip()
                
                # 尝试提取设备名称
                name_match = re.search(r'(?:设备描述|Device Description):\s*(.*?)(?:\r|\n|$)', block, re.IGNORECASE | re.DOTALL)
                device_name = name_match.group(1).strip() if name_match else "未知设备"
                
                devices_info.append({"id": device_id, "name": device_name})
    
    except subprocess.CalledProcessError:
        pass
    
    # 如果pnputil失败，尝试使用devcon
    if not devices_info:
        try:
            cmd = 'devcon findall *usb*'
            result = subprocess.check_output(cmd, shell=True, text=True)
            
            # 解析devcon输出
            lines = result.splitlines()
            for i in range(0, len(lines), 2):
                if i+1 < len(lines):
                    device_id = lines[i].strip()
                    device_name = lines[i+1].strip()
                    if "VID_" in device_id and "PID_" in device_id:
                        devices_info.append({"id": device_id, "name": device_name})
        
        except subprocess.CalledProcessError:
            pass
    
    return devices_info

class DeviceControllerGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("USB设备控制器")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)
        
        # 配置信息
        self.config = load_config()
        
        # 创建设备ID变量
        self.device_id_var = tk.StringVar(value=self.config.get("full_device_id", ""))
        self.current_device_id = self.config.get("full_device_id", "")
        
        # 设置界面布局
        self.setup_ui()
        
        # 加载设备
        self.load_current_device()
        
        # 添加定时器，定期刷新设备状态
        self.status_timer = None
        self.start_status_timer()
    
    def setup_ui(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建设备选择区域
        device_frame = ttk.LabelFrame(main_frame, text="设备选择", padding="10")
        device_frame.pack(fill=tk.X, pady=5)
        
        ttk.Label(device_frame, text="当前设备ID:").grid(row=0, column=0, sticky=tk.W, pady=5)
        
        device_id_entry = ttk.Entry(device_frame, textvariable=self.device_id_var, width=50)
        device_id_entry.grid(row=0, column=1, sticky=tk.W+tk.E, padx=5, pady=5)
        
        scan_button = ttk.Button(device_frame, text="扫描设备", command=self.scan_devices)
        scan_button.grid(row=0, column=2, padx=5, pady=5)
        
        device_frame.columnconfigure(1, weight=1)
        
        # 创建设备状态区域
        status_frame = ttk.LabelFrame(main_frame, text="设备状态", padding="10")
        status_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # 设备信息文本区域
        self.status_text = scrolledtext.ScrolledText(status_frame, wrap=tk.WORD, height=10, width=70)
        self.status_text.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.status_text.config(state=tk.DISABLED)
        
        # 创建操作按钮区域
        button_frame = ttk.Frame(main_frame, padding="10")
        button_frame.pack(fill=tk.X, pady=5)
        
        self.enable_button = ttk.Button(button_frame, text="启用设备", command=self.enable_current_device)
        self.enable_button.pack(side=tk.LEFT, padx=5)
        
        self.disable_button = ttk.Button(button_frame, text="禁用设备", command=self.disable_current_device)
        self.disable_button.pack(side=tk.LEFT, padx=5)
        
        refresh_button = ttk.Button(button_frame, text="刷新状态", command=self.refresh_device_status)
        refresh_button.pack(side=tk.LEFT, padx=5)
        
        # 添加状态栏
        self.status_bar = ttk.Label(main_frame, text="就绪", relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def load_current_device(self):
        """加载当前配置中的设备"""
        if self.config.get("use_full_id", False) and self.config.get("full_device_id"):
            self.current_device_id = self.config.get("full_device_id")
            self.device_id_var.set(self.current_device_id)
            self.refresh_device_status()
        else:
            self.log_message("未配置设备，请扫描并选择一个设备。")
    
    def scan_devices(self):
        """扫描并选择设备"""
        self.log_message("正在扫描USB设备...")
        
        # 在后台线程中扫描设备，避免界面卡顿
        threading.Thread(target=self._scan_devices_thread, daemon=True).start()
    
    def _scan_devices_thread(self):
        # 获取所有USB设备
        devices = list_all_usb_devices()
        
        if not devices:
            self.log_message("未找到任何USB设备。")
            return
        
        # 创建设备选择对话框
        self.root.after(0, self._show_device_selection_dialog, devices)
    
    def _show_device_selection_dialog(self, devices):
        dialog = tk.Toplevel(self.root)
        dialog.title("选择设备")
        dialog.geometry("600x400")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # 创建设备列表
        ttk.Label(dialog, text="请选择要控制的USB设备:").pack(pady=10)
        
        # 设备列表框
        device_frame = ttk.Frame(dialog)
        device_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # 设备列表
        columns = ("name", "id")
        device_tree = ttk.Treeview(device_frame, columns=columns, show="headings")
        device_tree.heading("name", text="设备名称")
        device_tree.heading("id", text="设备ID")
        device_tree.column("name", width=250)
        device_tree.column("id", width=300)
        
        # 添加设备到列表
        for device in devices:
            device_tree.insert("", tk.END, values=(device["name"], device["id"]))
        
        device_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # 添加滚动条
        scrollbar = ttk.Scrollbar(device_frame, orient=tk.VERTICAL, command=device_tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        device_tree.configure(yscrollcommand=scrollbar.set)
        
        # 按钮区域
        button_frame = ttk.Frame(dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def on_select():
            selection = device_tree.selection()
            if selection:
                item = device_tree.item(selection[0])
                device_id = item["values"][1]
                self.select_device(device_id)
                dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        select_button = ttk.Button(button_frame, text="选择", command=on_select)
        select_button.pack(side=tk.RIGHT, padx=5)
        
        cancel_button = ttk.Button(button_frame, text="取消", command=on_cancel)
        cancel_button.pack(side=tk.RIGHT, padx=5)
    
    def select_device(self, device_id):
        """选择并保存设备ID"""
        self.current_device_id = device_id
        self.device_id_var.set(device_id)
        
        # 更新配置
        self.config = update_config_with_device_id(device_id)
        
        self.log_message(f"已选择设备: {device_id}")
        self.refresh_device_status()
    
    def enable_current_device(self):
        """启用当前设备"""
        if not self.current_device_id:
            messagebox.showwarning("警告", "未选择设备")
            return
        
        # 禁用按钮，避免重复点击
        self.enable_button.config(state=tk.DISABLED)
        self.disable_button.config(state=tk.DISABLED)
        
        self.status_bar.config(text="正在启用设备...")
        self.log_message(f"正在启用设备: {self.current_device_id}")
        
        # 在后台线程中执行设备操作
        threading.Thread(target=self._enable_device_thread, daemon=True).start()
    
    def _enable_device_thread(self):
        result = enable_device(self.current_device_id)
        
        # 在主线程中更新UI
        self.root.after(0, self._handle_enable_result, result)
    
    def _handle_enable_result(self, result):
        if result:
            self.log_message("设备已成功启用")
            self.status_bar.config(text="设备已启用")
        else:
            self.log_message("启用设备失败")
            self.status_bar.config(text="操作失败")
            messagebox.showerror("错误", "启用设备失败")
        
        # 延迟刷新设备状态，给设备一些时间来改变状态
        self.root.after(2000, self.refresh_device_status)
        
        # 重新启用按钮
        self.enable_button.config(state=tk.NORMAL)
        self.disable_button.config(state=tk.NORMAL)
    
    def disable_current_device(self):
        """禁用当前设备"""
        if not self.current_device_id:
            messagebox.showwarning("警告", "未选择设备")
            return
        
        # 禁用按钮，避免重复点击
        self.enable_button.config(state=tk.DISABLED)
        self.disable_button.config(state=tk.DISABLED)
        
        self.status_bar.config(text="正在禁用设备...")
        self.log_message(f"正在禁用设备: {self.current_device_id}")
        
        # 在后台线程中执行设备操作
        threading.Thread(target=self._disable_device_thread, daemon=True).start()
    
    def _disable_device_thread(self):
        result = disable_device(self.current_device_id)
        
        # 在主线程中更新UI
        self.root.after(0, self._handle_disable_result, result)
    
    def _handle_disable_result(self, result):
        if result:
            self.log_message("设备已成功禁用")
            self.status_bar.config(text="设备已禁用")
        else:
            self.log_message("禁用设备失败")
            self.status_bar.config(text="操作失败")
            messagebox.showerror("错误", "禁用设备失败")
        
        # 延迟刷新设备状态，给设备一些时间来改变状态
        self.root.after(2000, self.refresh_device_status)
        
        # 重新启用按钮
        self.enable_button.config(state=tk.NORMAL)
        self.disable_button.config(state=tk.NORMAL)
    
    def refresh_device_status(self):
        """刷新设备状态"""
        if not self.current_device_id:
            return
        
        # 在后台线程中获取设备状态，避免界面卡顿
        threading.Thread(target=self._refresh_device_status_thread, daemon=True).start()
    
    def _refresh_device_status_thread(self):
        try:
            # 检查设备是否存在
            if not device_exists(self.current_device_id):
                self.root.after(0, self._update_status_not_found)
                return
            
            # 获取设备状态
            is_disabled = get_device_status(self.current_device_id)
            
            # 在主线程中更新UI
            self.root.after(0, self._update_status_ui, is_disabled)
            
        except Exception as e:
            self.root.after(0, self._update_status_error, str(e))
    
    def _update_status_ui(self, is_disabled):
        status_text = "已禁用" if is_disabled else "已启用"
        
        self.log_message(f"设备当前状态: {status_text}")
        self.status_bar.config(text=f"设备状态: {status_text}")
        
        # 更新按钮状态
        if is_disabled:
            self.enable_button.config(state=tk.NORMAL)
            self.disable_button.config(state=tk.DISABLED)
        else:
            self.enable_button.config(state=tk.DISABLED)
            self.disable_button.config(state=tk.NORMAL)
    
    def _update_status_not_found(self):
        self.log_message("设备未找到，请检查设备是否已连接")
        self.status_bar.config(text="设备未找到")
        
        # 禁用所有操作按钮
        self.enable_button.config(state=tk.DISABLED)
        self.disable_button.config(state=tk.DISABLED)
    
    def _update_status_error(self, error_message):
        self.log_message(f"获取设备状态出错: {error_message}")
        self.status_bar.config(text="获取状态出错")
    
    def log_message(self, message):
        """在状态文本框中添加消息"""
        self.status_text.config(state=tk.NORMAL)
        
        # 添加时间戳
        timestamp = time.strftime("%H:%M:%S")
        
        # 添加消息
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        
        # 滚动到底部
        self.status_text.see(tk.END)
        
        self.status_text.config(state=tk.DISABLED)
    
    def start_status_timer(self):
        """启动状态更新定时器"""
        self.refresh_device_status()
        # 每30秒刷新一次状态
        self.status_timer = self.root.after(30000, self.start_status_timer)
    
    def on_closing(self):
        """关闭窗口时清理资源"""
        if self.status_timer:
            self.root.after_cancel(self.status_timer)
        self.root.destroy()

def main():
    # 检查管理员权限，如果不是管理员，则自动提权
    if not is_admin():
        print("正在请求管理员权限...")
        elevate_privileges()
        # 退出当前非管理员进程
        sys.exit(0)
    
    # 创建GUI
    root = tk.Tk()
    app = DeviceControllerGUI(root)
    
    # 设置窗口关闭处理
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # 运行主循环
    root.mainloop()

if __name__ == "__main__":
    main() 