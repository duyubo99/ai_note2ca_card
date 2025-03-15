# 文件处理系统

这是一个基于 Flask 的 Web 应用，用于上传 JSON 文件并通过 gen_main.py 的逻辑处理后生成 Excel 和 PPT 文件。

## 功能特点

- 提供简洁的 Web 界面上传 JSON 文件
- 自动处理上传的文件并生成 Excel 和 PPT
- 显示处理进度和结果
- 提供已生成文件的下载链接
- 支持查看和刷新已生成的文件列表

## 安装步骤

1. 确保已安装 Python 3.7+
2. 安装依赖包：

```bash
pip install -r requirements.txt
```

## 运行应用

```bash
python app.py
```

启动后，在浏览器中访问 http://localhost:5000 即可使用应用。

## 使用说明

1. 点击"选择文件"按钮，选择要上传的 JSON 文件
2. 点击"上传并处理"按钮开始处理
3. 等待处理完成，系统会显示处理结果和生成的文件
4. 点击文件名可以下载生成的文件
5. "已生成的文件"区域显示所有已生成的文件，可以点击"刷新"按钮更新列表

## 文件结构

- `app.py`: Flask 应用主文件
- `templates/`: HTML 模板目录
- `static/`: 静态资源目录（CSS、JavaScript）
- `data/`: 数据目录
  - `input/`: 输入文件目录
  - `output/`: 输出文件目录
- `core/`: 核心处理模块

## 注意事项

- 上传的文件必须是 JSON 格式
- 文件大小限制为 16MB
- 生成的文件保存在 data/output 目录中
