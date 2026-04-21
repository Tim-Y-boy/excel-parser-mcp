# Excel Parser MCP Server

[English](README.md) | 中文

通用 Excel 转文本 MCP（模型上下文协议）服务器，专为 AI 智能体设计。无需任何配置，即可将 `.xlsx` 或 `.xls` 文件接入你的 LLM 工作流。

## 解决什么问题

| 痛点 | 方案 |
|------|------|
| LLM 无法直接读取 Excel | 将 Excel 转为结构化的、LLM 友好的文本 |
| 不同项目 Excel 格式各异 | 格式无关 —— 无需适配器或配置文件 |
| 合并单元格导致数据解析错误 | 自动解析合并单元格，继承左上角值 |
| 无关 Sheet 太多浪费 Token | 智能关键词过滤，减少 Token 消耗 |
| `.xlsx` 和 `.xls` 两种格式并存 | 统一读取器，透明支持两种格式 |

## 快速开始

### 安装

```bash
git clone https://github.com/Tim-Y-boy/excel-parser-mcp.git
cd excel-parser-mcp
pip install -e .
```

### 在 Claude Code 中配置

在项目的 `.claude/settings.json` 中添加：

```json
{
  "mcpServers": {
    "excel-parser": {
      "command": "python",
      "args": ["你本地的路径/excel-parser-mcp/mcp_server.py"]
    }
  }
}
```

重启 Claude Code，即可使用三个新工具。

### 在项目根目录创建 .mcp.json


{
  "mcpServers": {
    "excel-parser": {
      "command": "python",
      "args": ["你本地的路径/excel-parser-mcp/server.py"]
    }
  }
}

> **注意**：MCP Server 是本地运行的，`args` 需要指向你自己机器上 clone 的路径。

## 工具说明

### 1. `list_sheets`

列出所有 Sheet 及其元数据（行数、列数、合并单元格数）。

```
> 用 list_sheets 工具查看 "report.xlsx" 里有什么
```

返回：
```json
[
  {"name": "Sales Data", "rows": 150, "cols": 12, "merged_cells": 5},
  {"name": "Summary", "rows": 30, "cols": 8, "merged_cells": 2},
  {"name": "Change History", "rows": 20, "cols": 4, "merged_cells": 0}
]
```

### 2. `read_excel`

读取整个工作簿并转为 LLM 友好的文本。自动过滤无关 Sheet。

```
> 读取 "report.xlsx" 并总结销售数据
```

特性：
- **智能过滤** — 自动跳过 "Change History"、"Cover"、"Instructions" 等 Sheet
- **合并单元格解析** — 子单元格自动继承左上角的值
- **自动截断** — 单 Sheet 默认最多 200 行，控制 Token 消耗
- **哨兵检测** — 遇到 `#endofdata` 标记自动停止读取

可自定义过滤关键词：

```json
{
  "file_path": "report.xlsx",
  "relevant_keywords": ["sales", "revenue", "quarterly"],
  "exclude_keywords": ["draft", "internal"],
  "max_rows": 100
}
```

### 3. `read_single_sheet`

只读取一个指定的 Sheet，适用于不需要整个文件的场景。

```
> 读取 "report.xlsx" 中的 "Sales Data" Sheet
```

## 架构

```
excel-parser-mcp/
├── mcp_server.py              # MCP Server 入口（3 个工具）
├── excel_parser/
│   ├── reader.py              # 统一 .xlsx/.xls 读取器（~100 行）
│   ├── text_converter.py      # Sheet 转文本 + 智能过滤（~80 行）
│   └── value_normalizer.py    # 布尔/十六进制/时间标准化（~40 行）
└── pyproject.toml
```

### 工作原理

```
┌──────────────┐     ┌─────────────────┐     ┌──────────────┐
│  ExcelReader  │────>│ SheetTextConverter│────>│  LLM 智能体   │
│  (.xlsx/.xls)│     │ (过滤 + 格式化)   │     │ (Claude/GPT) │
└──────────────┘     └─────────────────┘     └──────────────┘
       │                      │
       v                      v
    SheetData              纯文本
   (行数据+合并信息)    "Row 1: A | B | C"
```

**核心设计决策：**

1. **格式无关** — 不做格式适配器或 YAML 配置。新格式零配置即可使用，格式理解交给 LLM。

2. **合并单元格解析** — 预计算 `{(row,col): 左上角坐标}` 映射。当 Service ID 合并了 3 行，所有 3 行都能看到同一个值。

3. **哨兵检测** — 识别 `#endofdata` / `#EndOfData` 标记（大小写不敏感），自动停止读取。

4. **智能过滤** — 两组关键词：`relevant`（保留）和 `exclude`（跳过）。都不匹配的 Sheet 默认保留（宁可多读也不漏数据）。

## 作为 Python 库使用

不使用 MCP 也可以直接调用：

```python
from excel_parser import ExcelReader, SheetTextConverter

# 读取任意 Excel 文件
reader = ExcelReader("any_file.xlsx")
sheets = reader.read_all_sheets()

# 转为文本（带过滤）
converter = SheetTextConverter(
    relevant_keywords=["你的", "关键词"],
    exclude_keywords=["跳过", "这些"],
)
text = converter.convert_workbook(sheets)

# 传给任意 LLM
print(text)
```

## 系统要求

- Python 3.12+
- openpyxl >= 3.1.0（用于 .xlsx）
- xlrd >= 2.0.0（用于 .xls）
- mcp >= 1.0.0（用于 MCP Server 模式）

## License

MIT
