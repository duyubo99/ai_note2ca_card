# AI 笔录解析系统

这是一个基于 Flask 的 Web 应用程序，用于处理 Word 文档（.docx）并生成 Excel 和 PowerPoint 格式的输出文件。

## 功能特点

- 支持多个 Word 文档（.docx）的上传和处理
- 可选择生成 Excel 和/或 PowerPoint 格式的输出
- 直观的用户界面，支持拖放文件上传
- 实时处理状态显示
- 处理完成后可直接下载生成的文件

## 技术栈

- **后端**：Python, Flask
- **前端**：HTML, CSS, JavaScript, Bootstrap 5
- **数据处理**：
  - python-docx（Word 文档处理）
  - openpyxl（Excel 生成）
  - python-pptx（PowerPoint 生成）

## 安装与运行

### 前提条件

- Python 3.6+
- pip（Python 包管理器）

### 安装步骤

1. 克隆或下载本仓库
2. 安装依赖包：

```bash
pip install -r requirements.txt
```

### 运行应用

在 Windows 系统上，可以直接运行批处理文件：

```bash
run_web.bat
```

或者手动启动：

```bash
python app.py
```

应用将在 http://localhost:5000 上运行。

## 使用说明

1. 打开浏览器访问 http://localhost:5000
2. 点击上传区域或拖放文件到上传区域，选择要处理的 Word 文档（.docx）
3. 选择需要的输出格式（Excel 和/或 PowerPoint）
4. 点击"开始处理"按钮
5. 等待处理完成后，点击下载按钮获取生成的文件

## 项目结构

```
ai_note2ca_card/
├── app.py                  # Flask应用主文件
├── requirements.txt        # 依赖包列表
├── run_web.bat             # Windows启动脚本
├── static/                 # 静态资源
│   ├── css/                # CSS样式文件
│   ├── js/                 # JavaScript文件
│   └── favicon.ico         # 网站图标
├── templates/              # HTML模板
│   ├── index.html          # 首页模板
│   └── results.html        # 结果页模板
├── core/                   # 核心处理模块
│   ├── ai_core.py          # AI处理核心
│   ├── excel_generator.py  # Excel生成器
│   ├── flatten_aijson.py   # JSON数据处理
│   ├── genText.py          # 文本生成
│   ├── ppt_generator.py    # PPT生成器
│   └── prompt_manager.py   # 提示管理
└── data/                   # 数据目录
    ├── input/              # 输入数据
    │   ├── docx/           # Word文档
    │   ├── json/           # JSON数据
    │   └── temp_upload/    # 临时上传文件
    └── output/             # 输出数据
```

## 注意事项

- 上传的 Word 文档必须是.docx 格式
- 处理大文件可能需要较长时间，请耐心等待
- 生成的文件将保存在 data/output 目录中
