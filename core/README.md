# AI Core 模块优化说明

## 优化内容

`ai_core.py` 模块已经进行了以下优化：

1. **模块化设计**：将原有功能拆分为多个独立函数，每个函数负责特定的任务
2. **外部调用接口**：新增 `process_file()` 函数，可以直接传入文件路径进行处理
3. **参数灵活性**：支持自定义输入/输出目录和文件名
4. **返回处理结果**：函数返回处理后的结果字典，方便进一步处理
5. **完善的文档**：为每个函数添加了详细的文档字符串，说明参数和返回值

## 主要函数说明

### `process_file(file_path, output_dir='data/input/json/temp', output_filename='ai_json.json')`

这是主要的外部调用接口，用于处理单个 Word 文档并生成结构化 JSON 数据。

**参数**：

- `file_path`：要处理的 Word 文档路径（必需）
- `output_dir`：输出目录（可选，默认为'data/input/json/temp'）
- `output_filename`：输出文件名（可选，默认为'ai_json.json'）

**返回值**：

- 处理结果字典，包含从文档中提取的结构化数据

### `extract_word_tables(file_path=None, input_dir='data/input/docx')`

从 Word 文档中提取表格内容。

**参数**：

- `file_path`：单个文档路径，如果提供则只处理该文档
- `input_dir`：文档目录，当 file_path 为 None 时使用

**返回值**：

- 包含(文档路径, 表格内容)元组的列表

### `process_docx(file_path=None, input_dir='data/input/docx', output_dir='data/input/json/temp', output_filename='ai_json.json')`

处理 Word 文档并生成结构化 JSON 数据（异步函数）。

**参数**：

- `file_path`：单个文档路径，如果提供则只处理该文档
- `input_dir`：文档目录，当 file_path 为 None 时使用
- `output_dir`：输出目录
- `output_filename`：输出文件名

**返回值**：

- 处理结果字典

## 使用示例

### 基本用法

```python
from core.ai_core import process_file

# 处理单个文档
result = process_file(
    file_path='data/input/docx/example.docx',
    output_dir='data/input/json/ai',
    output_filename='result.json'
)

# 使用处理结果
print(result.keys())  # 查看处理的文档路径
```

### 高级用法

请参考项目根目录下的 `example_usage.py` 和 `advanced_usage.py` 文件，这些文件展示了如何在简单和复杂场景中使用优化后的 `ai_core.py` 模块。

- `example_usage.py`：演示基本用法
- `advanced_usage.py`：演示如何在完整工作流中集成使用

## 注意事项

1. 确保已安装所有必要的依赖项（docx, tqdm 等）
2. 处理大型文档时可能需要较长时间，请耐心等待
3. 输出目录会自动创建，无需手动创建
