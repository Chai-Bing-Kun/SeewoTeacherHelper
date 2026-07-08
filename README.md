# Seevo Teacher Helper - 触屏教学助手

专为 Windows 触屏教学场景设计的辅助工具，帮助教师快速打开 U 盘中的 PPT 课件，并自动检测外置设备驱动程序的运行状态。

## ✨ 功能特点

### 📖 一键打开 PPT
- 点击按钮自动扫描 U 盘（可移动磁盘）中的 PPT/PPTX/PPS 文件
- 自动列出所有找到的 PPT 文件，双击即可打开
- 支持手动浏览选择任意位置的 PPT 文件
- 使用系统默认程序（PowerPoint/WPS）打开

### 🔧 外置程序管理
- 添加需要配合使用的程序（如翻页笔驱动、音响控制台、手写板驱动等）
- 自动检测程序是否正在运行
- 一键启动未运行的程序
- 支持自定义启动参数

### ⚙ 智能检测
- **开机检测**：程序启动时自动检测所有外置程序运行状态
- **打开PPT前检测**：点击打开PPT时自动检测外置程序是否就绪
- 检测结果清晰展示，未运行的程序可一键启动

### 🖥 系统托盘
- 关闭窗口时自动最小化到系统托盘
- 托盘菜单快速访问主要功能
- 后台运行，不干扰教学

## 📦 安装使用

### 方法一：用安装包安装（推荐）
1. 下载 `installer/SeevoTeacherHelper_Setup.exe`
2. 双击运行，按提示选择安装位置
3. 安装完成后程序会自动启动，桌面也会生成快捷方式

### 方法二：直接运行 exe
1. 从 `dist/` 文件夹获取 `SeevoTeacherHelper.exe`
2. 双击运行即可（适合临时使用，不用安装）

### 方法三：源码运行（适合开发者）
```bash
# 安装依赖
pip install -r requirements.txt

# 运行程序
python app/main.py
```

### 打包为 exe
```bash
# 安装打包工具
pip install pyinstaller

# 执行打包
python build_exe.py

# 打包后的 exe 在 dist/ 目录下
```

## 🎯 使用场景

1. **上课前**：双击运行本程序，自动检测翻页笔、音响等驱动是否正常
2. **打开课件**：插入 U 盘，点击「一键打开PPT」，自动扫描并列出课件
3. **自定义程序**：在「外置程序」选项卡中添加教学常用的辅助工具

## 🛠 技术栈

- Python 3.8+
- PyQt5（GUI 界面）
- PyInstaller（打包为 exe）

## 📄 配置文件

程序配置保存在 `app/config.json` 中，包括：
- 外置程序列表
- 检测开关设置
- 窗口大小

## � 项目文件说明

下面是这个项目里每个文件夹和文件是干什么用的，方便大家理解和修改：

```
SeevoTeacherHelper/          ← 项目根目录（整个项目的家）
│
├── app/                     ← 📂 程序源代码（核心代码都在这）
│   ├── main.py              ← 📄 主程序（程序的"大脑"，界面、搜索、检测等功能都在这里）
│   └── error_helper.py      ← 📄 错误处理工具（程序出错时弹出友好提示，告诉你怎么解决）
│
├── installer/               ← 📂 安装程序制作（用来生成安装包）
│   ├── setup_script.iss     ← 📄 安装脚本（告诉安装工具怎么打包成安装程序）
│   └── SeevoTeacherHelper_Setup.exe  ← 📦 安装包（双击它就能把程序装到电脑上）
│
├── tests/                   ← 📂 测试文件（用来测试程序好不好用）
│   ├── test_app.py          ← 📄 测试用模拟程序（假装是一个外置程序，用来测试检测功能）
│   └── test_files/          ← 📂 测试课件（里面有几个空的PPT和测试程序）
│       ├── 数学课件.pptx
│       ├── 英语课件.pptx
│       ├── 语文课件.pptx
│       ├── 测试课件1.pptx
│       ├── 测试程序1.exe
│       └── 测试程序2.exe
│
├── dist/                    ← 📂 打包好的程序（编译后的成品）
│   └── SeevoTeacherHelper.exe  ← 📦 单文件程序（下载这个就能直接运行，不用安装Python）
│
├── build/                   ← 📂 打包中间文件（打包时自动生成的临时文件，不用管它）
│
├── .venv/                   ← 📂 Python虚拟环境（用来隔离项目依赖，避免和其他程序冲突）
│
├── .gitignore               ← 📄 Git忽略规则（告诉Git哪些文件不需要上传到GitHub）
├── build_exe.py             ← 📄 打包脚本（运行它就能把源代码打包成exe文件）
├── requirements.txt         ← 📄 依赖清单（记录了这个项目用到了哪些Python库）
├── SeevoTeacherHelper.spec  ← 📄 打包配置文件（PyInstaller打包时用的配置文件）
└── README.md                ← 📄 本文件（项目介绍和使用说明）
```

### 简单来说

| 如果你想... | 请关注这个文件/文件夹 |
|------------|---------------------|
| **直接使用程序** | `dist/SeevoTeacherHelper.exe` 或 `installer/SeevoTeacherHelper_Setup.exe` |
| **修改程序代码** | `app/main.py`（主要功能）和 `app/error_helper.py`（错误提示） |
| **重新打包成exe** | 运行 `build_exe.py` |
| **制作安装包** | `installer/setup_script.iss` |
| **测试功能** | `tests/` 文件夹里的内容 |
| **查看用了哪些库** | `requirements.txt` |

## ⚠ 注意事项

- 本程序仅支持 Windows 系统
- 需要安装 PowerPoint 或 WPS 来打开 PPT 文件
- U 盘扫描仅搜索可移动磁盘，如需扫描其他磁盘请使用「浏览其他位置」
- 配置文件 `config.json` 会在程序启动时自动生成，**不会上传到 GitHub**（已在 `.gitignore` 中忽略）
