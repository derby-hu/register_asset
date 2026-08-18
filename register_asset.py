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
import tkinter as tk
from tkinter import ttk, messagebox
import psutil
import requests
import re
import os
import tempfile
import datetime
import openpyxl

# 设置默认编码为UTF-8（解决Windows控制台GBK编码不支持emoji的问题）
if sys.platform == 'win32':
    import locale
    try:
        locale.setlocale(locale.LC_ALL, 'zh_CN.UTF-8')
    except:
        pass
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except:
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
    print("❌ 未找到配置文件 config.py，请复制 config.example.py 并填写配置")
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
                dept_name = dept.get('name')
                try:
                    users = self.get_user_list(dept_id)
                    if users:
                        all_users.extend(users)
                except Exception as e:
                    print(f"获取部门 '{dept_name}' (ID:{dept_id}) 用户失败: {e}")
                    continue
            
            # 去重
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
        
        # 过滤没有手机号的用户
        return [u for u in all_users if u.get('mobile')]

    def send_verify_code(self, userid, code):
        """发送验证码"""
        token = self.get_access_token()
        url = "https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2"
        params = {"access_token": token}
        data = {
            "agent_id": DINGTALK_AGENT_ID,
            "userid_list": userid,
            "msg": {
                "msgtype": "text",
                "text": {"content": f"【资产登记】您的验证码是：{code}\n有效期 5 分钟。"}
            }
        }
        
        resp = requests.post(url, json=data, params=params, timeout=10)
        res = resp.json()
        if res.get("errcode") == 0 or res.get("code") == 0:
            return True
        raise Exception(f"发送失败：{res}")

    def notify_admin(self, name, count):
        """通知网管（纯文本提示）"""
        if not ADMIN_USER_ID: 
            print("⚠️ 未配置管理员ID，跳过通知")
            return
        
        try:
            token = self.get_access_token()
            url = "https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2"
            params = {"access_token": token}
            content = f"新资产登记提醒\n员工：{name}\n数量：{count} 个 MAC 地址\n登记信息Excel文件已通过附件方式发送，请查收。"
            data = {
                "agent_id": DINGTALK_AGENT_ID,
                "userid_list": ADMIN_USER_ID,
                "msg": {
                    "msgtype": "text",
                    "text": {"content": content}
                }
            }
            
            resp = requests.post(url, json=data, params=params, timeout=10)
            res = resp.json()
            
            if res.get("errcode") == 0:
                print(f"📧 管理员文本通知发送成功")
            else:
                print(f"⚠️ 管理员文本通知发送失败：{res.get('errmsg')}")
        except Exception as e:
            print(f"⚠️ 管理员文本通知发送异常：{e}")

    def upload_media(self, file_path):
        """上传媒体文件到钉钉，返回media_id"""
        try:
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
                media_id = res.get("media_id")
                print(f"✅ 媒体文件上传成功，media_id: {str(media_id)[:20]}...")
                return media_id
            else:
                raise Exception(f"上传失败: {res}")
        except Exception as e:
            print(f"⚠️ 媒体文件上传异常：{e}")
            return None

    def send_file_to_admin(self, file_path, name, count, macs):
        """将Excel文件发送给管理员"""
        if not ADMIN_USER_ID:
            print("⚠️ 未配置管理员ID，跳过发送文件")
            return False
        
        media_id = self.upload_media(file_path)
        if not media_id:
            return False
        
        try:
            token = self.get_access_token()
            url = "https://oapi.dingtalk.com/topapi/message/corpconversation/asyncsend_v2"
            params = {"access_token": token}
            
            data = {
                "agent_id": DINGTALK_AGENT_ID,
                "userid_list": ADMIN_USER_ID,
                "msg": {
                    "msgtype": "file",
                    "file": {"media_id": media_id}
                }
            }
            
            resp = requests.post(url, json=data, params=params, timeout=10)
            res = resp.json()
            
            if res.get("errcode") == 0:
                print(f"📎 Excel文件已通过钉钉发送给管理员")
                return True
            else:
                print(f"⚠️ 文件发送失败：{res.get('errmsg')}")
                return False
        except Exception as e:
            print(f"⚠️ 文件发送异常：{e}")
            return False


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
        print(f"📄 Excel已生成：{file_path} ({os.path.getsize(file_path)} bytes)")
        return file_path
    except Exception as e:
        print(f"❌ 生成Excel失败：{e}")
        return None


def generate_code():
    """生成6位数字验证码"""
    return "".join(random.choices(string.digits, k=6))

def get_physical_macs():
    """获取本机所有物理网卡 MAC"""
    mac_list = []
    addrs = psutil.net_if_addrs()
    
    # 跳过虚拟网卡
    skip_keywords = ['virtual', 'vmware', 'virtualbox', 'hyper-v', 'loopback', 'bluetooth', 'wsl', 'docker']
    
    for iface, addr_list in addrs.items():
        if any(k in iface.lower() for k in skip_keywords):
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
    
    def __init__(self, parent, users):
        """初始化对话框"""
        super().__init__(parent)
        self.title("君问公司资产自动登记")
        self.geometry("400x320")
        self.resizable(False, False)
        self.attributes('-topmost', True)
        
        self.users = users
        self.selected_user = None
        self.verify_code_sent = None
        self.confirmed = False
        
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
        
        # 标题
        header_frame = ttk.Frame(self, style='TFrame')
        header_frame.pack(fill='x', padx=30, pady=(20, 10))
        ttk.Label(header_frame, text="君问公司资产自动登记工具", style='Header.TLabel').pack()
        
        # 提示文字
        ttk.Label(self, text="请选择本人并进行身份验证", style='TLabel', foreground='#666666').pack(pady=(5, 20))
        
        # 主内容区
        main_frame = ttk.Frame(self, style='TFrame')
        main_frame.pack(fill='both', expand=True, padx=30)
        
        # 第一行：姓名选择 + 获取验证码
        f1 = ttk.Frame(main_frame, style='TFrame')
        f1.pack(fill='x', pady=15)
        ttk.Label(f1, text="选择您的姓名", width=10, style='TLabel').pack(side='left')
        self.combo = ttk.Combobox(f1, values=[], state="normal", width=15)
        self.combo.pack(side='left', padx=5)
        self.combo.bind("<KeyRelease>", self._on_search)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self._on_select())
        self.btn_send = ttk.Button(f1, text="获取验证码", command=self._send_code, width=12)
        self.btn_send.pack(side='left', padx=10)

        # 第二行：验证码输入 + 确定登记
        f3 = ttk.Frame(main_frame, style='TFrame')
        f3.pack(fill='x', pady=15)
        ttk.Label(f3, text="请输入验证码", width=10, style='TLabel').pack(side='left')
        self.entry_code = ttk.Entry(f3, width=15)
        self.entry_code.pack(side='left', padx=5)
        ttk.Button(f3, text="确 定 登 记", command=self._ok, style='Primary.TButton', width=12).pack(side='left', padx=10)
        
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

    def _send_code(self):
        """发送验证码"""
        if not self.selected_user:
            messagebox.showwarning("提示", "请先选择您的姓名！")
            return
        
        self.btn_send.config(state="disabled", text="发送中...")
        try:
            code = generate_code()
            client.send_verify_code(self.selected_user['userid'], code)
            self.verify_code_sent = code
            self.btn_send.config(state="normal", text="重新发送")
            messagebox.showinfo("成功", "验证码已发送至您的钉钉！")
        except Exception as e:
            self.btn_send.config(state="normal", text="获取验证码")
            messagebox.showerror("发送失败", str(e))

    def _ok(self):
        """确认登记"""
        if not self.selected_user:
            messagebox.showwarning("提示", "未选择姓名")
            return
        if not self.verify_code_sent:
            messagebox.showwarning("提示", "请先点击【获取验证码】")
            return
        if self.entry_code.get().strip() != self.verify_code_sent:
            messagebox.showerror("验证失败", "验证码错误，请重试")
            return
        
        self.confirmed = True
        self.destroy()

def main():
    """主函数"""
    root = tk.Tk()
    
    # 设置主窗口样式（与操作界面一致）
    root.title("君问公司资产自动登记")
    root.geometry("400x200")
    root.resizable(False, False)
    root.attributes('-topmost', True)
    root.configure(bg="#f5f5f5")
    
    # 设置样式（与操作界面一致）
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TFrame', background='#f5f5f5')
    style.configure('TLabel', background='#f5f5f5', font=('Microsoft YaHei', 10), foreground='#666666')
    style.configure('Header.TLabel', background='#f5f5f5', font=('Microsoft YaHei', 14, 'bold'), foreground='#333333')
    
    # 标题区域（与操作界面一致）
    header_frame = ttk.Frame(root, style='TFrame')
    header_frame.pack(fill='x', padx=30, pady=(30, 10))
    ttk.Label(header_frame, text="君问公司资产自动登记工具", style='Header.TLabel').pack()
    
    # 提示文字（与操作界面一致）
    ttk.Label(root, text="程序正在启动，请稍等……", style='TLabel').pack(pady=20)
    
    root.update()
    
    print("🚀 正在启动资产登记助手...")
    
    global client
    
    try:
        print("📱 初始化钉钉客户端...")
        client = DingTalkClient(DINGTALK_APP_KEY, DINGTALK_APP_SECRET)
        
        print("📞 正在同步通讯录...")
        users = client.get_all_users()
        
        if not users:
            messagebox.showerror("错误", "无法获取通讯录，请检查 AppKey/Secret 及权限。")
            root.destroy()
            return
            
        print(f"✅ 成功获取 {len(users)} 名员工信息")
            
    except Exception as e:
        messagebox.showerror("初始化失败", str(e))
        root.destroy()
        return
    
    # 关闭启动界面，显示操作窗口
    root.destroy()
    
    # 创建新的主窗口
    root = tk.Tk()
    root.withdraw()

    # 2. 弹出交互窗口
    dialog = RegDialog(root, users)
    root.wait_window(dialog)
    
    if not dialog.confirmed:
        print("用户取消操作。")
        return

    user = dialog.selected_user
    print(f"✅ 用户 {user['name']} 验证通过。")
    
    # 3. 扫描物理网卡
    print("🔍 正在扫描本地物理网卡...")
    macs = get_physical_macs()
    
    if not macs:
        messagebox.showerror("错误", "未检测到任何物理网卡地址。\n请检查网络连接或联系管理员。")
        return

    # 4. 显示扫描到的网卡
    print(f"✅ 扫描到 {len(macs)} 个物理网卡：")
    for m in macs:
        print(f"   - {m}")
    
    # 5. 生成Excel并发送给管理员
    print("📝 正在生成登记信息Excel文件...")
    
    # 显示进度
    progress_win = tk.Toplevel(root)
    progress_win.title("处理中")
    progress_win.geometry("320x110")
    progress_win.attributes('-topmost', True)
    ttk.Label(progress_win, text=f"发现 {len(macs)} 个网卡\n正在生成Excel并发送给管理员...").pack(pady=20)
    root.update()

    excel_path = None
    try:
        # 生成Excel
        excel_path = generate_mac_excel(user['name'], macs)
        if not excel_path:
            raise Exception("Excel文件生成失败")
        
        print("📤 正在通过钉钉发送Excel给管理员...")
        
        # 发送文本通知 + Excel附件
        file_sent = client.send_file_to_admin(excel_path, user['name'], len(macs), macs)
        client.notify_admin(user['name'], len(macs))
        
        success_count = len(macs)
        
        # 清理进度窗口
        progress_win.destroy()
        
        # 6. 结果反馈
        msg = f"登记完成！\n\n员工：{user['name']}\n成功登记：{success_count} 个 MAC 地址\n"
        if file_sent:
            msg += "登记Excel已发送给管理员。"
        else:
            msg += "Excel发送失败，请手动将文件交给管理员。\n文件位置：\n" + excel_path
        
        messagebox.showinfo("完成", msg)
        
        # 最后销毁主窗口
        root.destroy()
            
    except Exception as e:
        # 错误处理
        progress_win.destroy()
        # 保留生成的临时文件便于排错
        root.destroy()
        extra = f"\n生成的文件：{excel_path}" if excel_path else ""
        messagebox.showerror("错误", f"登记失败：{e}{extra}")
        print(f"❌ 登记失败：{e}")

if __name__ == "__main__":
    try:
        import psutil, requests, tkinter
        import openpyxl
    except ImportError as e:
        print(f"❌ 缺少依赖库：{e}")
        print("请运行命令安装：pip install psutil requests openpyxl")
        input("按回车退出...")
        sys.exit(1)
    
    # 运行主函数
    main()