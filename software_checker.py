import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import psutil
import winreg
import re
import tkinter.messagebox
import csv
import os
import logging

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
        # 获取屏幕宽度和高度
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        # 计算窗口的 x 和 y 坐标
        x = (screen_width - 800) // 2
        y = (screen_height - 900) // 2
        # 设置窗口位置
        self.root.geometry(f"800x900+{x}+{y}")
        self.root.resizable(False, False)

        # 创建菜单栏
        self.menu_bar = tk.Menu(self.root)
        self.root.config(menu=self.menu_bar)

        # 直接添加关于菜单项（顶层菜单）
        self.menu_bar.add_command(label="关于", command=self.show_about)

        # 导入PIL库（用于图片显示）


        # 创建按钮框架（整体居中）
        self.button_frame = ttk.Frame(self.root)
        self.button_frame.pack(pady=10, fill='x', anchor='center')  # 父框架整体居中

        # 执行检查按钮（框架内顶部）
        self.check_button = ttk.Button(self.button_frame, text="执行检查", command=self.check_software)
        self.check_button.pack(pady=5)

        # 导出/批量删除按钮子框架（水平排列并居中）
        self.export_batch_frame = ttk.Frame(self.button_frame)
        self.export_batch_frame.pack(pady=5, anchor='center')  # 子框架居中

        # 导出结果按钮
        self.export_button = ttk.Button(self.export_batch_frame, text="导出结果", command=self.export_results, state='disabled')
        self.export_button.pack(side='left', padx=(0, 20))

        # 打开卸载页面按钮
        self.open_uninstall_button = ttk.Button(self.export_batch_frame, text="打开软件卸载页面", command=self.open_control_panel_uninstall)
        self.open_uninstall_button.pack(side='left', padx=(0, 20))
        # 自定义工具提示
        def show_tooltip(widget, text):
            def enter(event):
                tooltip = tkinter.Toplevel(widget)
                tooltip.wm_overrideredirect(True)
                tooltip.wm_geometry(f"+{event.x_root+10}+{event.y_root+10}")
                ttk.Label(tooltip, text=text, background="#ffffe0", relief="solid", borderwidth=1).pack()
                widget.tooltip = tooltip
            def leave(event):
                if hasattr(widget, 'tooltip'):
                    widget.tooltip.destroy()
            widget.bind("<Enter>", enter)
            widget.bind("<Leave>", leave)
        show_tooltip(self.open_uninstall_button, "点击打开系统控制面板的卸载程序页面")



        # 统计信息框架（Treeview正上方）
        self.stats_frame = ttk.Frame(self.root)
        self.stats_frame.pack(pady=5, fill='x', anchor='center')
        self.stats_label = ttk.Label(self.stats_frame, text="")
        self.stats_label.pack()

        # 创建Treeview（按钮框架下方）
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

        # 计算统计信息
        total_software = len(processes)
        licensed_software = sum(1 for p in processes if any(re.search(licensed, p, re.IGNORECASE) for licensed in LICENSED_SOFTWARE))
        unlicensed_software = total_software - licensed_software
        # 系统软件和正版化清单内正版数（假设为正版软件数，因系统软件已被过滤）
        system_licensed_software = licensed_software
        self.stats_label.config(text=f"软件总数：{total_software} | 正版软件数：{licensed_software} | 非法软件数：{unlicensed_software}")

        # 检查完成后恢复按钮状态，并启用导出按钮
        self.check_button.config(state='normal', text='执行检查')
        self.export_button.config(state='normal')

        # 检查是否有需要卸载的软件以启用卸载按钮
        has_uninstall_items = any(
            len(self.tree.item(item, 'values')) >=4 and self.tree.item(item, 'values')[3] == '请立即卸载'
            for item in self.tree.get_children()
        )

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

    def show_about(self):
        # 创建关于对话框
        about_window = tk.Toplevel(self.root)
        about_window.title("关于")
        about_window.resizable(False, False)

        # 计算对话框居中位置
        main_x = self.root.winfo_x()
        main_y = self.root.winfo_y()
        main_width = self.root.winfo_width()
        main_height = self.root.winfo_height()
        dialog_width = 400
        dialog_height = 300
        x = main_x + (main_width - dialog_width) // 2
        y = main_y + (main_height - dialog_height) // 2
        about_window.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        # 版本和开发者信息
        info_label = ttk.Label(about_window, text="软件版本: 1.0.0\n\n开发者: 甘肃分行金融科技部", justify="center")
        info_label.pack(pady=10)

    def open_control_panel_uninstall(self):
        # 打开控制面板的程序和功能（卸载）页面
        import subprocess
        # 使用start命令调用控制面板的appwiz.cpl（程序和功能）
        subprocess.run(['start', 'control', 'appwiz.cpl'], shell=True, check=True)

    # 原批量卸载功能已移除

    def get_uninstall_command(self, software_name):
        # 从注册表查找软件对应的卸载命令（优化匹配逻辑）
        reg_paths = [
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
            r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        ]

        for path in reg_paths:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
                for i in range(winreg.QueryInfoKey(key)[0]):
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        display_name = winreg.QueryValueEx(subkey, 'DisplayName')[0].strip().lower()
                        # 修复：使用精确小写匹配代替正则（避免误匹配）
                        if display_name == software_name.strip().lower():
                            return winreg.QueryValueEx(subkey, 'UninstallString')[0]
                    except OSError:
                        continue
                    finally:
                        subkey.Close()
                key.Close()
            except Exception as e:
                continue
        return None


    def _get_uninstall_path(self, software_name):
        # 从注册表读取卸载路径（简化示例）
        try:
            keys = [r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall", r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall"]
            for key_path in keys:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                for i in range(0, winreg.QueryInfoKey(key)[0]):
                    subkey_name = winreg.EnumKey(key, i)
                    subkey = winreg.OpenKey(key, subkey_name)
                    try:
                        display_name, _ = winreg.QueryValueEx(subkey, 'DisplayName')
                        if display_name == software_name:
                            uninstall_string, _ = winreg.QueryValueEx(subkey, 'UninstallString')
                            return uninstall_string
                    except OSError:
                        continue
                    finally:
                        subkey.Close()
                key.Close()
        except Exception as e:
            logging.error(f"读取注册表失败：{e}")
        return None

        # 示例：模拟卸载（实际需调用系统卸载命令）
        for software in uninstall_items:
            # 实际应从注册表获取卸载命令（示例伪代码）
            # uninstall_cmd = get_uninstall_command(software)
            # subprocess.run(uninstall_cmd, shell=True)
            tkinter.messagebox.showinfo("卸载提示", f"正在模拟卸载：{software}")

        tkinter.messagebox.showinfo("完成", "批量卸载操作已完成（模拟）")

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