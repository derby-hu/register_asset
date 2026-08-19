#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
君问公司资产自动登记工具
功能：
1. 通过钉钉进行用户身份验证
2. 自动扫描本机物理网卡MAC地址
3. 将MAC地址登记信息生成Excel文件
4. 通过钉钉将Excel文件发送给管理员进行处理
"""

import sys
import time
import random
import string
import re
import os
import json
import subprocess
import tempfile
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import psutil
import requests
import openpyxl

# 设置默认编码为UTF-8（解决Windows控制台GBK编码不支持emoji的问题）
if sys.platform == 'win32':
    import locale
    try:
        locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
    except Exception:
        pass
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

# ======================= 配置导入 =======================
# 从外部配置文件导入敏感信息
try:
    from config import (
        DINGTALK_APP_KEY,
        DINGTALK_APP_SECRET,
        DINGTALK_AGENT_ID,
        ADMIN_USER_ID,
    )
except ImportError:
    try:
        _rt = tk.Tk(); _rt.withdraw()
        messagebox.showerror("配置缺失", "未找到配置文件 config.py\n请复制 config.example.py 并填写配置。")
        _rt.destroy()
    except Exception:
        pass
    exit(1)
# =============================================================

class DingTalkClient:
    """钉钉客户端类，用于处理用户验证和通知"""
    
    def __init__(self, app_key, app_secret):
        """初始化钉钉客户端"""
        self.app_key = app_key
        self.app_secret = app_secret
        self.access_token = None
        self.token_expires_at = 0

    def get_access_token(self):
        """获取访问令牌"""
        if self.access_token and time.time() < self.token_expires_at:
            return self.access_token
        
        url = "https://oapi.dingtalk.com/gettoken"
        params = {"appkey": self.app_key, "appsecret": self.app_secret}
        try:
            resp = requests.get(url, params=params, timeout=10)
            data = resp.json()
            if data.get("errcode") == 0:
                self.access_token = data["access_token"]
                self.token_expires_at = time.time() + data["expires_in"] - 60
                return self.access_token
            else:
                raise Exception(f"Token 获取失败：{data}")
        except Exception as e:
            raise Exception(f"网络请求错误：{e}")

    def get_department_list(self, parent_id=1):
        """获取部门列表"""
        token = self.get_access_token()
        url = "https://oapi.dingtalk.com/department/list"
        params = {"access_token": token, "id": parent_id}
        try:
            r = requests.get(url, params=params, timeout=10)
            d = r.json()
            if d.get("errcode") != 0:
                raise Exception(f"获取部门列表失败: {d.get('errmsg')}")
            return d.get("department", [])
        except Exception as e:
            raise Exception(f"获取部门列表异常: {e}")

    def get_all_users(self):
        """获取所有部门的用户"""
        all_users = []
        
        try:
            departments = self.get_department_list(1)
            if not departments:
                raise Exception("无法获取部门列表")
            
            for dept in departments:
                dept_id = dept.get('id')
                try:
                    users = self.get_user_list(dept_id)
                    if users:
                        all_users.extend(users)
                except Exception:
                    continue
            
            unique_users = {u.get('userid'): u for u in all_users}
            return list(unique_users.values())
            
        except Exception as e:
            raise Exception(f"获取所有用户失败: {e}")

    def get_user_list(self, department_id=1):
        """获取指定部门的员工"""
        token = self.get_access_token()
        url = "https://oapi.dingtalk.com/user/listbypage"
        all_users = []
        offset = 0
        page_size = 50
        
        while True:
            params = {"access_token": token, "department_id": department_id, "offset": offset, "size": page_size}
            try:
                r = requests.get(url, params=params, timeout=10)
                d = r.json()
                if d.get("errcode") != 0:
                    raise Exception(f"获取通讯录失败: {d.get('errmsg')}")
                users = d.get("userlist", [])
                all_users.extend(users)
                if len(users) < page_size:
                    break
                offset += page_size
            except Exception as e:
                raise Exception(f"获取通讯录异常: {e}")
        
        return [u for u in all_users if u.get('mobile')]

    def _send_work_msg(self, userid, msg):
        """统一封装发送工作通知，失败抛异常"""
        token = self.get_access_token()
        url = "https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2"
        params = {"access_token": token}
        data = {
            "agent_id": DINGTALK_AGENT_ID,
            "userid_list": userid,
            "msg": msg,
        }
        resp = requests.post(url, json=data, params=params, timeout=10)
        res = resp.json()
        if res.get("errcode") != 0 and res.get("code") != 0:
            raise Exception(f"{res.get('errmsg') or res}")
        return True

    def send_verify_code(self, userid, code):
        """发送验证码"""
        self._send_work_msg(userid, {
            "msgtype": "text",
            "text": {"content": f"【资产登记】您的验证码是：{code}\n有效期 5 分钟。"}
        })
        return True

    def upload_media(self, file_path):
        """上传媒体文件到钉钉，返回media_id"""
        token = self.get_access_token()
        url = "https://oapi.dingtalk.com/media/upload"
        params = {"access_token": token, "type": "file"}
        
        if not os.path.exists(file_path):
            raise Exception(f"文件不存在: {file_path}")
        
        with open(file_path, 'rb') as f:
            files = {"media": (os.path.basename(file_path), f, "application/octet-stream")}
            resp = requests.post(url, params=params, files=files, timeout=60)
        
        res = resp.json()
        if res.get("errcode") == 0:
            return res.get("media_id")
        raise Exception(f"上传失败: {res}")

    def send_file_to_admin(self, file_path):
        """将Excel文件发送给管理员（仅文件）"""
        if not ADMIN_USER_ID:
            raise Exception("未配置管理员ID")

        media_id = self.upload_media(file_path)
        if not media_id:
            raise Exception("文件上传失败，未获取到 media_id")

        self._send_work_msg(ADMIN_USER_ID, {
            "msgtype": "file",
            "file": {"media_id": media_id}
        })
        return True


def generate_mac_excel(user_name, macs):
    """生成MAC登记信息的Excel文件，返回临时文件路径"""
    try:
        ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_name = re.sub(r"[^\w\u4e00-\u9fa5]", "_", user_name)
        file_name = f"MAC登记_{safe_name}_{ts}.xlsx"
        
        tmp_dir = tempfile.gettempdir()
        file_path = os.path.join(tmp_dir, file_name)
        
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "MAC登记"
        
        # 数据表头
        ws.cell(row=1, column=1, value='MAC')
        ws.cell(row=1, column=2, value='有效日期')
        ws.cell(row=1, column=3, value='描述（请输入0-128位字符）')
        
        for idx, mac in enumerate(macs, start=1):
            row_no = 1 + idx
            ws.cell(row=row_no, column=1, value=mac)
            ws.cell(row=row_no, column=2, value=0)
            ws.cell(row=row_no, column=3, value=user_name)
        
        # 简单调整列宽
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 40
        
        wb.save(file_path)
        return file_path
    except Exception as e:
        raise Exception(f"生成Excel失败：{e}")


def generate_code():
    """生成6位数字验证码"""
    return "".join(random.choices(string.digits, k=6))

def get_physical_macs():
    """获取本机所有物理网卡 MAC

    通过 Windows Get-NetAdapter 的 InterfaceDescription 字段过滤虚拟网卡。
    描述信息由驱动厂商定义（如 "Hyper-V Virtual Ethernet Adapter"），不受用户改网卡名影响，
    比基于网卡名关键字过滤更稳定可靠。
    """
    # 获取网卡名→描述信息映射（描述由驱动厂商定义）
    ps_script = (
        "$ErrorActionPreference='Stop'; "
        "Get-NetAdapter -IncludeHidden | ForEach-Object { "
        "[PSCustomObject]@{Name=$_.Name;Desc=$_.InterfaceDescription} "
        "} | ConvertTo-Json -Compress"
    )
    r = subprocess.run(
        ['powershell', '-NoProfile', '-Command', ps_script],
        capture_output=True, timeout=15,
        creationflags=subprocess.CREATE_NO_WINDOW,  # 避免在 GUI 程序中弹出控制台窗口
    )
    # ConvertTo-Json 对非 ASCII 字符转成 \uXXXX 转义，优先 utf-8
    raw = None
    for enc in ('utf-8', 'gbk', 'gb18030'):
        try:
            decoded = r.stdout.decode(enc)
            json.loads(decoded)
            raw = decoded
            break
        except Exception:
            continue
    if raw is None:
        raise Exception("无法解析 Get-NetAdapter 输出")
    data = json.loads(raw)
    if isinstance(data, dict):
        data = [data]
    desc_map = {item.get('Name', ''): item.get('Desc', '') or '' for item in data}

    # 描述信息中的虚拟网卡 / 伪接口关键字（厂商定义，稳定可靠）
    skip_desc_keywords = [
        'virtual',          # Hyper-V Virtual Ethernet Adapter / Wi-Fi Direct Virtual Adapter / Virtual Switch
        'miniport',         # WAN Miniport (SSTP/IP/IPv6/...)
        'tunnel',           # Teredo Tunneling Pseudo-Interface
        'pseudo-interface', # 各类伪接口
        'teredo', '6to4', 'isatap',  # 隧道协议适配器
        'kernel debug',     # Microsoft Kernel Debug Network Adapter
        'bluetooth',        # Bluetooth Device (PAN)
    ]

    mac_list = []
    addrs = psutil.net_if_addrs()
    for iface, addr_list in addrs.items():
        desc = desc_map.get(iface, '').lower()
        if any(k in desc for k in skip_desc_keywords):
            continue

        for addr in addr_list:
            if addr.family == psutil.AF_LINK and addr.address:
                # 格式化MAC地址
                mac = addr.address.replace(':', '-').upper()
                if re.match(r"^([0-9A-F]{2}-){5}[0-9A-F]{2}$", mac):
                    mac_list.append(mac)

    return mac_list

class RegDialog(tk.Toplevel):
    """用户注册对话框"""

    def __init__(self, parent, users, on_confirmed=None):
        """初始化对话框"""
        super().__init__(parent)
        self.title("君问公司资产自动登记")
        self.geometry("400x360")
        self.resizable(False, False)
        self.attributes('-topmost', True)

        self.users = users
        self.selected_user = None
        self.verify_code_sent = None
        self._status_after_id = None
        self._on_confirmed = on_confirmed

        self.configure(bg="#f5f5f5")
        self._setup_ui()

    def _setup_ui(self):
        """设置界面"""
        # 样式设置
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background='#f5f5f5')
        style.configure('TLabel', background='#f5f5f5', font=('Microsoft YaHei', 10))
        style.configure('Header.TLabel', background='#f5f5f5', font=('Microsoft YaHei', 14, 'bold'), foreground='#333333')
        style.configure('TButton', font=('Microsoft YaHei', 10), padding=8)
        style.configure('Primary.TButton', font=('Microsoft YaHei', 11, 'bold'), foreground='white', background='#4a90e2')
        style.map('Primary.TButton', foreground=[('active', 'white')], background=[('active', '#357abd')])
        style.configure('TEntry', fieldbackground='white', padding=6)
        style.configure('TCombobox', fieldbackground='white', padding=6)
        style.configure('Status.TLabel', background='#ffffff', relief='solid', borderwidth=1,
                        font=('Microsoft YaHei', 10, 'bold'), padding=8)

        # 标题
        header_frame = ttk.Frame(self, style='TFrame')
        header_frame.pack(fill='x', padx=30, pady=(20, 10))
        ttk.Label(header_frame, text="君问公司资产自动登记工具", style='Header.TLabel').pack()

        # 提示文字
        ttk.Label(self, text="请选择本人并进行身份验证", style='TLabel', foreground='#666666').pack(pady=(5, 15))

        # 主内容区
        main_frame = ttk.Frame(self, style='TFrame')
        main_frame.pack(fill='both', expand=True, padx=30)

        # 第一行：姓名选择 + 获取验证码
        f1 = ttk.Frame(main_frame, style='TFrame')
        f1.pack(fill='x', pady=10)
        ttk.Label(f1, text="选择您的姓名", width=10, style='TLabel').pack(side='left')
        self.combo = ttk.Combobox(f1, values=[], state="normal", width=15)
        self.combo.pack(side='left', padx=5)
        self.combo.bind("<KeyRelease>", self._on_search)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self._on_select())
        self.btn_send = ttk.Button(f1, text="获取验证码", command=self._send_code, width=12)
        self.btn_send.pack(side='left', padx=10)

        # 第二行：验证码输入 + 确定登记
        f3 = ttk.Frame(main_frame, style='TFrame')
        f3.pack(fill='x', pady=10)
        ttk.Label(f3, text="请输入验证码", width=10, style='TLabel').pack(side='left')
        self.entry_code = ttk.Entry(f3, width=15)
        self.entry_code.pack(side='left', padx=5)
        ttk.Button(f3, text="确 定 登 记", command=self._ok, style='Primary.TButton', width=12).pack(side='left', padx=10)

        # 状态提示（原生tk.Label确保颜色生效，非阻塞）
        self.status_label = tk.Label(self, text="", anchor='center',
                                     font=('Microsoft YaHei', 10, 'bold'),
                                     bg="#f5f5f5", fg="#333333",
                                     relief='flat', borderwidth=0,
                                     pady=10, padx=10, height=2, wraplength=340,
                                     justify='center')
        self.status_label.pack(fill='x', padx=30, pady=(10, 15))

        # 初始化数据
        self._init_search_data()

    def _init_search_data(self):
        """初始化搜索数据"""
        self.user_map = {}
        self.all_names = []
        for u in self.users:
            key = u['name']
            self.user_map[key] = u
            self.all_names.append(key)
        self.all_names.sort()
        self.combo['values'] = self.all_names

    def _on_search(self, event):
        """搜索姓名"""
        search_text = self.combo.get().lower()
        if not search_text:
            self.combo['values'] = self.all_names
            return
        
        filtered_names = [name for name in self.all_names if search_text in name.lower()]
        self.combo['values'] = filtered_names
        if filtered_names:
            self.combo.event_generate('<Down>')

    def _on_select(self):
        """选择姓名"""
        key = self.combo.get()
        self.selected_user = self.user_map.get(key)
        if self.selected_user:
            self.entry_code.delete(0, tk.END)
            self.verify_code_sent = None

    def _set_status(self, text, level='info', duration=5000):
        """设置非阻塞状态提示文字，指定毫秒数后自动清空。
        level: info(蓝)/success(绿)/warn(黄)/error(红)
        """
        # (前景色, 背景色, 边框色) —— 前景用深色确保高对比度可读
        palette = {
            'info':    ('#1a4b8c', '#e7f0fb', '#357abd'),
            'success': ('#1e5a33', '#e6f5ec', '#2d8a4e'),
            'warn':    ('#7a5200', '#fdf4dc', '#d9a22e'),
            'error':   ('#8a0000', '#fdeaea', '#cc0000'),
        }
        fg, bg, bd = palette.get(level, palette['info'])
        # 取消上一次未触发的清空定时器
        if self._status_after_id is not None:
            try:
                self.status_label.after_cancel(self._status_after_id)
            except Exception:
                pass
            self._status_after_id = None

        if not text:
            self._clear_status()
            return

        self.status_label.config(text=text, fg=fg, bg=bg, relief='solid',
                                 borderwidth=2, highlightbackground=bd,
                                 highlightthickness=1)
        self.status_label.update_idletasks()
        if duration and duration > 0:
            self._status_after_id = self.status_label.after(
                duration, self._clear_status)

    def _clear_status(self):
        self.status_label.config(text="", bg="#f5f5f5", fg="#333333",
                                 relief='flat', borderwidth=0,
                                 highlightthickness=0)
        self._status_after_id = None

    def _send_code(self):
        """发送验证码"""
        if not self.selected_user:
            self._set_status("请先选择您的姓名！", level='warn')
            return

        self.btn_send.config(state="disabled", text="发送中...")
        self._set_status("正在发送验证码，请稍候…", level='info', duration=0)
        try:
            code = generate_code()
            client.send_verify_code(self.selected_user['userid'], code)
            self.verify_code_sent = code
            self.btn_send.config(state="normal", text="重新发送")
            self._set_status("✓ 验证码已发送至您的钉钉！", level='success', duration=6000)
        except Exception as e:
            self.btn_send.config(state="normal", text="获取验证码")
            self._set_status(f"发送失败：{e}", level='error', duration=10000)

    def _ok(self):
        """确认登记"""
        if not self.selected_user:
            self._set_status("请先选择您的姓名！", level='warn')
            return
        if not self.verify_code_sent:
            self._set_status("请先点击【获取验证码】获取验证码", level='warn')
            return
        if self.entry_code.get().strip() != self.verify_code_sent:
            self._set_status("验证码错误，请重新输入", level='error')
            self.entry_code.delete(0, tk.END)
            return

        self.btn_send.config(state="disabled")
        self.entry_code.config(state="disabled")
        # 通过 after 调用回调，避免阻塞当前事件处理
        if self._on_confirmed:
            self.after(100, self._on_confirmed, self)

    def show_progress(self, text):
        """在状态栏显示处理进度"""
        self._set_status(text, level='info', duration=0)
        self.update_idletasks()
        self.update()

def _create_prompt_win(parent, initial_text=""):
    """创建提示信息窗口（统一样式和居中位置）"""
    win = tk.Toplevel(parent)
    win.title("提示")
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - 300) // 2
    y = (sh - 150) // 2
    win.geometry(f"300x150+{x}+{y}")
    win.resizable(False, False)
    win.attributes('-topmost', True)
    win.configure(bg="#f5f5f5")
    label = ttk.Label(win, text=initial_text, style='Info.TLabel', wraplength=280)
    label.pack(expand=True)
    win.update_idletasks()
    win.update()
    win.lift()
    win.focus_force()
    return win, label


def _show_error_in_win(root, win, label, message):
    """在提示窗口中显示错误信息并等待用户确认"""
    label.config(text=message, style='Error.TLabel')
    ttk.Button(win, text="确定", command=win.destroy).pack(pady=10)
    win.update()
    root.wait_window(win)
    root.destroy()


def _show_result_in_win(dialog, message, success=True):
    """关闭主对话框，弹出提示信息窗口显示登记结果并等待用户确认。

    成功用 Info 样式（灰色），失败用 Error 样式（红色）。
    不销毁 root，由 main 末尾统一销毁。
    """
    root = dialog.master
    dialog.destroy()

    win, label = _create_prompt_win(root, message)
    label.config(style='Info.TLabel' if success else 'Error.TLabel')
    ttk.Button(win, text="确定", command=win.destroy).pack(pady=10)
    win.update()
    root.wait_window(win)


def main():
    """主函数"""
    root = tk.Tk()
    root.withdraw()

    # 设置统一样式
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('Info.TLabel', background='#f5f5f5', font=('Microsoft YaHei', 10), foreground='#666666')
    style.configure('Error.TLabel', background='#f5f5f5', font=('Microsoft YaHei', 10), foreground='#cc0000')
    style.configure('TButton', font=('Microsoft YaHei', 10), padding=8)

    # === 初始化阶段：弹出提示窗口显示进度 ===
    info_win, info_label = _create_prompt_win(root, "程序正在启动，请稍等……")

    global client

    try:
        info_label.config(text="正在初始化钉钉客户端...")
        info_win.update_idletasks()
        info_win.update()
        client = DingTalkClient(DINGTALK_APP_KEY, DINGTALK_APP_SECRET)

        info_label.config(text="正在同步通讯录...")
        info_win.update_idletasks()
        info_win.update()
        users = client.get_all_users()

        if not users:
            _show_error_in_win(root, info_win, info_label, "无法获取通讯录，请检查配置及权限。")
            return

    except Exception as e:
        _show_error_in_win(root, info_win, info_label, f"初始化失败：{e}")
        return

    # 成功：关闭提示窗口，弹出主程序窗口
    info_win.destroy()

    def _process_registration(dialog):
        """验证通过后的处理流程，在 dialog 状态栏显示进度"""
        user = dialog.selected_user
        try:
            dialog.show_progress("正在扫描本地网卡...")
            macs = get_physical_macs()

            if not macs:
                _show_result_in_win(dialog, "未检测到物理网卡，请检查网络连接。", success=False)
                return

            dialog.show_progress("正在生成登记文件...")
            excel_path = generate_mac_excel(user['name'], macs)

            dialog.show_progress("正在发送给管理员...")
            client.send_file_to_admin(excel_path)

            _show_result_in_win(dialog, "资产登记完成！请联系网络管理员。", success=True)

        except Exception as e:
            _show_result_in_win(dialog, f"登记失败：{e}", success=False)

    dialog = RegDialog(root, users, on_confirmed=_process_registration)
    root.wait_window(dialog)
    root.destroy()

if __name__ == "__main__":
    main()