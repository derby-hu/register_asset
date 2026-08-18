# 君问公司资产自动登记工具

> 基于钉钉认证的企业资产自动登记系统，实现员工身份验证和MAC地址自动采集登记。

## 📋 功能特点

- **钉钉身份验证**：通过钉钉企业应用进行员工身份验证
- **验证码验证**：发送验证码到员工钉钉，确保身份真实性
- **MAC地址自动采集**：自动扫描本机所有物理网卡MAC地址
- **Excel自动生成**：自动生成MAC地址登记Excel文件
- **钉钉文件发送**：通过钉钉工作通知将Excel附件发送给管理员
- **管理员通知**：登记完成后自动通知网管

## 🛠️ 技术栈

| 分类 | 技术 |
|------|------|
| 语言 | Python 3.8+ |
| GUI框架 | Tkinter |
| 网络请求 | requests |
| Excel处理 | openpyxl |
| 系统信息 | psutil |
| 打包工具 | PyInstaller |

## 🚀 快速开始

### 1. 安装依赖

```bash
pip install -r requirements.txt
```

### 2. 配置钉钉应用

1. 打开 [钉钉开放平台](https://open.dingtalk.com/)
2. 进入「应用开发」→ 创建应用（企业内部应用）
3. 命名为「资产自动登记」
4. 配置权限：
   - `employee.basicinfo` - 获取员工基本信息
   - `message` - 发送工作通知消息
   - `qyapi_get_department_member` - 获取部门成员列表
   - `qyapi_get_department_list` - 获取部门信息

### 3. 配置应用

```bash
# 复制配置模板
cp config.example.py config.py

# 编辑配置文件
notepad config.py
```

配置说明：

```python
# 钉钉应用凭证
DINGTALK_APP_KEY = "你的APP_KEY"
DINGTALK_APP_SECRET = "你的APP_SECRET"
DINGTALK_AGENT_ID = "你的AGENT_ID"

# 管理员通知（必填：登记Excel会发给该管理员）
ADMIN_USER_ID = "管理员的userid"
```

### 4. 运行程序

```bash
# 开发模式运行
python register_asset.py

# 或者运行编译后的EXE
dist/register_asset.exe
```

## 📊 登记流程

```
┌─────────────────────────────────────────────────────────────┐
│                    资产自动登记流程                          │
├─────────────────────────────────────────────────────────────┤
│  1. 启动程序                                               │
│     ↓                                                      │
│  2. 显示加载窗口，同步钉钉通讯录                             │
│     ↓                                                      │
│  3. 选择员工姓名                                            │
│     ↓                                                      │
│  4. 获取并输入钉钉验证码                                     │
│     ↓                                                      │
│  5. 自动扫描物理网卡MAC地址                                  │
│     ↓                                                      │
│  6. 生成MAC登记Excel文件                                    │
│     ↓                                                      │
│  7. 通过钉钉发送Excel给管理员 + 文本通知                    │
│     ↓                                                      │
│  8. 显示登记成功提示                                         │
└─────────────────────────────────────────────────────────────┘
```

## 📁 项目结构

```
register_asset/
├── register_asset.py    # 主程序
├── config.py            # 配置文件（不提交）
├── config.example.py    # 配置模板
├── requirements.txt     # 依赖列表
├── register_asset.spec  # PyInstaller配置
└── README.md            # 项目说明
```

## 🔒 安全注意事项

1. **敏感信息保护**：`config.py` 包含敏感信息，已加入 `.gitignore`
2. **权限最小化**：钉钉应用只申请必要的权限
3. **HTTPS通信**：所有网络请求使用HTTPS协议
4. **验证码机制**：确保操作人为真实员工

## 📝 更新日志

| 版本 | 日期 | 更新内容 |
|------|------|----------|
| v1.0 | 2024-01-XX | 初始版本，实现基本功能 |
| v1.1 | 2024-01-XX | 添加启动加载窗口，优化用户体验 |
| v1.2 | 2024-01-XX | 分离敏感配置，增强安全性 |
| v2.0 | 2026-08-XX | 移除NAS依赖，改为生成Excel通过钉钉发送给管理员 |

## 📄 许可证

MIT License

---

**注意**：本工具仅供君问公司内部使用，未经授权请勿用于其他用途。
