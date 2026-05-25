# Paper Format Skill

用于分析论文模板或格式要求，并把目标 `.docx` 论文调整到对应格式的 Codex Skill。

## 用途

- 从 PDF、DOCX 或文字版格式要求中提取论文排版规则
- 生成结构化 `format_rules.json`
- 按确认后的规则修改目标 `.docx`
- 核查页边距、标题层级、正文字体、行距、页眉页脚、图表题注和参考文献格式
- 输出格式核查报告，方便复查修改是否完成

## 目录结构

```text
paper-format-skill/
├── SKILL.md
├── README.md
├── paper-format.skill
├── examples/
│   ├── format_rules_template.json
│   └── sample_output_report.txt
└── LICENSE
```

## 安装方法

直接使用根目录下的 `paper-format.skill` 安装包导入 Codex。

如需从源码查看或修改 skill 主体，请编辑 `SKILL.md`。

## 使用方法

在 Codex 中提出类似请求，并同时提供模板和待排版论文：

```text
请按这个 PDF 模板格式，帮我排版我的论文 docx，并输出核查报告。
```

或：

```text
提取这个论文模板的格式规则，先给我确认，再套用到我的论文。
```

Skill 会按三个阶段执行：

1. 分析模板并提取规则
2. 等用户确认后应用格式
3. 重新核查并输出报告

## 示例截图

当前仓库未附带真实运行截图。建议在完成一次格式核查后，把核查报告或最终文档预览截图放到 `examples/`，再在这里引用：

```markdown
![格式核查报告示例](examples/report-screenshot.png)
```

## 示例文件

- `examples/format_rules_template.json`：格式规则 JSON 模板
- `examples/sample_output_report.txt`：示例核查报告
