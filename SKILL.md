---
name: docx2md
description: Convert Office documents (doc/docx/xlsx/xls) to AI-readable Markdown. TRIGGER when: user asks to read doc/docx/xlsx/xls files; user asks to convert/transform these formats; AI determines it needs to read content from these file types; user mentions analyzing Word/Excel documents. This skill makes document content accessible for AI analysis.
metadata:
  short-description: Convert doc/docx/xlsx to markdown
---

# docx2md

将Office文档转换为AI可读的Markdown格式。

## 触发场景

- 用户要求读取doc/docx/xlsx/xls文件内容
- 用户主动要求转换/处理这些格式
- AI判断需要读取这些文件进行分析
- 用户提到分析Word/Excel文档

## 支持格式

| 格式 | 处理方式 |
|------|----------|
| doc | Word COM转docx + pandoc转换 + OLE附件提取 |
| docx | pandoc转换 + OLE附件提取 |
| xlsx/xls | openpyxl读取转为markdown表格 |
| pdf | MarkItDown提取 + 正则后处理识别标题结构 |

## 附件处理

docx中嵌入的OLE附件同样支持转换：

| 附件类型 | 处理方式 |
|----------|----------|
| Excel (.xlsx/.xls) | 提取原文件 + 自动生成.md |
| Word (.docx) | 提取原文件 + pandoc转.md |
| Word (.doc) | 提取原文件 + Word COM转docx + pandoc转.md |
| 图片 | 放入images目录 |
| 其他 | 原样保存到attachments目录 |

## 用法

```bash
python <skill_dir>/scripts/docx2md.py <input_file>
```

## 输出结构

**doc/docx**:
```
{filename}.md
{filename}_files/
  ├── images/image_1.png, ...
  └── attachments/{name}.xlsx, {name}.md, ...
```

**xlsx**: `{filename}.md` (包含所有工作表)
