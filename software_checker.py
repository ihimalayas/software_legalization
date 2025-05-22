import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import psutil
import winreg
import re
import tkinter.messagebox
import csv
import os

# 从文件读取正版软件清单
import os

# 从文件读取系统软件清单
SYSTEM_SOFTWARE_FILE = os.path.join(os.path.dirname(__file__), 'system_software_list.txt')


def read_system_software():
    try:
        with open(SYSTEM_SOFTWARE_FILE, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            return [line.strip() for line in lines if line.strip() and not line.strip().startswith('#')]
    except FileNotFoundError:
        print('系统软件清单文件未找到，请检查文件路径。')
        return []

LICENSED_SOFTWARE_FILE = os.path.join(os.path.dirname(__file__), 'licensed_software_list.txt')

def read_licensed_software():
    try:
        with open(LICENSED_SOFTWARE_FILE, 'r', encoding='utf-8') as file:
            lines = file.readlines()
            return [line.strip() for line in lines if line.strip() and not line.strip().startswith('#')]
    except FileNotFoundError:
        print('正版软件清单文件未找到，请检查文件路径。')
        return []

LICENSED_SOFTWARE = read_licensed_software()

class SoftwareCheckerApp:
    def __init__(self, root):
        # 确保每次启动时重新读取清单
        global LICENSED_SOFTWARE
        LICENSED_SOFTWARE = read_licensed_software()
        global SYSTEM_SOFTWARE
        SYSTEM_SOFTWARE = read_system_software()
        self.root = root
        self.root.title("软件正版化检查工具")
        self.root.geometry("800x900")
        self.root.resizable(False, False)

        # 创建执行检查按钮
        self.check_button = ttk.Button(self.root, text="执行检查", command=self.check_software)
        self.check_button.pack(pady=10)

        # 创建导出按钮
        self.export_button = ttk.Button(self.root, text="导出结果", command=self.export_results, state='disabled')
        self.export_button.pack(pady=10)

        # 创建Treeview
        self.tree = ttk.Treeview(self.root, columns=("序号", "软件名称", "是否为正版化软件", "处置建议"), show="headings")
        self.tree.heading("序号", text="序号", anchor=tk.CENTER)
        self.tree.column("序号", anchor=tk.CENTER)
        self.tree.heading("软件名称", text="软件名称", anchor=tk.W)
        self.tree.column("软件名称", anchor=tk.W)
        self.tree.heading("是否为正版化软件", text="是否为正版化软件", anchor=tk.CENTER)
        self.tree.column("是否为正版化软件", anchor=tk.CENTER)
        self.tree.heading("处置建议", text="处置建议", anchor=tk.CENTER)
        self.tree.column("处置建议", anchor=tk.CENTER)

        # 设置列宽自动调整
        self.tree.column("序号", width=50, stretch=tk.NO)
        self.tree.column("软件名称", stretch=tk.YES)
        self.tree.column("是否为正版化软件", width=110, stretch=tk.NO)
        self.tree.column("处置建议", width=130, stretch=tk.NO)

        # 创建滚动条
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # 布局
        # 修改布局，为 Treeview 左侧增加 5 个单位的距离
        self.tree.pack(side="left", fill="both", expand=True, pady=(0, 5), padx=(5, 0))
        scrollbar.pack(side="right", fill="y", pady=(0, 5))

        # # 添加版本信息和开发者信息
        # self.version_label = ttk.Label(self.root, text="版本号: 1.0.0, 开发者: 示例团队")
        # self.version_label.pack(side="bottom", pady=10)

    def check_software(self):
        # 添加加载提示
        self.check_button.config(state='disabled', text='检查中...')
        self.root.update_idletasks()

        # 清空Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 从Windows注册表读取控制面板程序
        installed_software = []
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                    installed_software.append(display_name)
                except OSError:
                    continue
                finally:
                    subkey.Close()
            key.Close()
            # 32位程序在64位系统的注册表位置
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                    installed_software.append(display_name)
                except OSError:
                    continue
                finally:
                    subkey.Close()
            key.Close()
        except Exception as e:
            print(f'读取注册表时出错: {e}')

        # 从系统软件清单中过滤系统程序
        SYSTEM_SOFTWARE = read_system_software()
        processes = [name for name in installed_software if not any(re.search(system_sw, name, re.IGNORECASE) for system_sw in SYSTEM_SOFTWARE)]

        index = 1
        for process in processes:
            
            try:
                software_name = process
                is_licensed = any(re.search(licensed, software_name, re.IGNORECASE) for licensed in LICENSED_SOFTWARE)
                suggestion = '保留' if is_licensed else '请立即卸载'
 
                if suggestion == '保留':
                    is_licensed_display = '✓'
                else:
                    is_licensed_display = '×'
                if suggestion == '请立即卸载':
                    action = '卸载'
                    iid = self.tree.insert('', 'end', values=(index, software_name, is_licensed_display, suggestion, action))
                    self.tree.item(iid, tags=('button',))
                else:
                    self.tree.insert('', 'end', values=(index, software_name, is_licensed_display, suggestion, ''))
                if not is_licensed:
                    self.tree.tag_configure('unlicensed', foreground='red')
                    self.tree.item(self.tree.get_children()[-1], tags=('unlicensed',))
                    index += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        # 检查完成后恢复按钮状态，并启用导出按钮
        self.check_button.config(state='normal', text='执行检查')
        self.export_button.config(state='normal')

    def export_results(self):
        try:
            # 导出 Treeview 结果到 CSV 文件
            file_path = "software_check_result.csv"
            with open(file_path, 'w', newline='', encoding='utf-8-sig') as file:
                writer = csv.writer(file)
                writer.writerow(["序号", "软件名称", "是否为正版化软件", "处置建议"])
                for item in self.tree.get_children(): 
                    values = self.tree.item(item, 'values')
                    writer.writerow(values[:4])
            tkinter.messagebox.showinfo("成功", "导出成功！")
        except Exception as e:
            tkinter.messagebox.showerror("错误", f"导出失败: {str(e)}")

def read_installed_software():
    installed_software = []
    try:
        keys = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]
        for key_path in keys:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                    installed_software.append(display_name)
                except OSError:
                    continue
                finally:
                    subkey.Close()
            key.Close()
    except Exception as e:
        logging.error('读取注册表时出错: %s', e)
    return installed_software

# 在 check_software 方法中调用
    def check_software(self):
        SYSTEM_SOFTWARE = read_software_list(SYSTEM_SOFTWARE_FILE)
        installed_software = read_installed_software()
        # 添加加载提示
        self.check_button.config(state='disabled', text='检查中...')
        self.root.update_idletasks()

        # 清空Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # 从Windows注册表读取控制面板程序
        installed_software = []
        try:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                    installed_software.append(display_name)
                except OSError:
                    continue
                finally:
                    subkey.Close()
            key.Close()
            # 32位程序在64位系统的注册表位置
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
            for i in range(0, winreg.QueryInfoKey(key)[0]):
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                try:
                    display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                    installed_software.append(display_name)
                except OSError:
                    continue
                finally:
                    subkey.Close()
            key.Close()
        except Exception as e:
            print(f'读取注册表时出错: {e}')

        # 从系统软件清单中过滤系统程序
        SYSTEM_SOFTWARE = read_system_software()
        processes = [name for name in installed_software if not any(re.search(system_sw, name, re.IGNORECASE) for system_sw in SYSTEM_SOFTWARE)]

        index = 1
        for process in processes:
            
            try:
                software_name = process
                is_licensed = any(re.search(licensed, software_name, re.IGNORECASE) for licensed in LICENSED_SOFTWARE)
                suggestion = '保留' if is_licensed else '请立即卸载'

                # 修改显示逻辑，确保处置建议为保留时显示 √
                is_licensed_display = '✓' if suggestion == '保留' else '×'

                if suggestion == '请立即卸载':
                    action = '卸载'
                    iid = self.tree.insert('', 'end', values=(index, software_name, is_licensed_display, suggestion, action))
                    self.tree.item(iid, tags=('button',))
                else:
                    self.tree.insert('', 'end', values=(index, software_name, is_licensed_display, suggestion, ''))
                if not is_licensed:
                    self.tree.tag_configure('unlicensed', foreground='red')
                    self.tree.item(self.tree.get_children()[-1], tags=('unlicensed',))
                    index += 1
            except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
                continue

        # 检查完成后恢复按钮状态
        self.check_button.config(state='normal', text='执行检查')

if __name__ == "__main__":
    root = ThemedTk(theme='arc')
    app = SoftwareCheckerApp(root)
    root.mainloop()

logging.basicConfig(level=logging.ERROR, format='%(asctime)s - %(levelname)s - %(message)s')