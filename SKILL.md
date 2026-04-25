---
name: docx2md
description: Convert Office documents (docx/xlsx/xls) to AI-readable Markdown. TRIGGER when: user asks to read docx/xlsx/xls files; user asks to convert/transform these formats; AI determines it needs to read content from these file types; user mentions analyzing Word/Excel documents. This skill makes document content accessible for AI analysis.
metadata:
  short-description: Convert docx/xlsx to markdown
---

# docx2md

将Office文档转换为AI可读的Markdown格式。

## 触发场景

- 用户要求读取docx/xlsx/xls文件内容
- 用户主动要求转换/处理这些格式
- AI判断需要读取这些文件进行分析
- 用户提到分析Word/Excel文档

## 支持格式

| 格式 | 处理方式 |
|------|----------|
| docx | pandoc转换 + OLE附件提取 |
| xlsx/xls | openpyxl读取转为markdown表格 |

## 工作流程

1. 转换文档为markdown格式
2. AI读取生成的.md文件获取内容
3. docx中的Excel附件自动生成同名.md，AI可直接读取表格数据

## 用法

```bash
python ~/.claude/skills/docx2md/scripts/docx2md.py <input_file>
```

## 输出结构

**docx**:
```
{filename}.md
{filename}_files/
  ├── images/{filename}_image_1.png, ...
  └── attachments/{name}.xlsx, {name}.md, ...
```

**xlsx**: `{filename}.md` (包含所有工作表)