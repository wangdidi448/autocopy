# 快捷工作台 autocopy

一个悬浮在 Windows 桌面的快捷信息工作台：把常用信息（姓名、手机号、地址、邮编、密码等）拆成一条条字段，**点击即复制**；支持分组管理，并能自动获取邮箱 / 短信验证码。

## 功能

- 无边框圆角悬浮窗，可拖动、置顶、折叠
- 字段点击即复制；字段 key 可改名（组内唯一校验）、着色、加粗
- 一键隐藏 / 显示所有 value（只看 key）
- 分组管理（如「求职信息」「应用密码」），可新建 / 改名 / 删除组
- 邮箱验证码：IMAP 绑定后自动识别新邮件中的验证码并复制（已适配 163/126/QQ/Gmail/Outlook 等）
- 短信验证码：安卓手机通过 USB（ADB）读取新短信验证码，ADB 可一键下载安装
- 数据全部本地保存（JSON），授权码经 Windows DPAPI 加密

## 运行环境

- Windows 10/11
- Python 3.10+
- 依赖：`pip install PySide6`

## 启动

```bash
python qt_workdesk.py
```

或双击 `启动快捷复制.bat`（无控制台窗口）。

首次启动会在程序目录自动生成 `quick_copy_data.json`（字段数据）与 `verify_config.json`（验证码绑定配置，已被 .gitignore 忽略）。

## 目录说明

| 文件 | 说明 |
|---|---|
| `qt_workdesk.py` | 主程序（PySide6 / Qt，推荐） |
| `quick_copy_widget.py` | 早期 tkinter 版本，备份 |
| `启动快捷复制.bat` | Windows 双击启动器 |
