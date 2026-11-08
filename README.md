# AMA Reference Formatting Skill

A Codex skill for checking and formatting Word reference lists in **American Marketing Association journal style**. This is the marketing-journal style, not the American Medical Association citation style.

## What it does

- Inventories every source before editing and checks the final count.
- Verifies uncertain metadata against publishers, DOI records, and institutional sources.
- Normalizes author names, title capitalization, alphabetization, page ranges, and container italics.
- Builds a Word reference list with 12-point Times New Roman, one-inch margins, single spacing, and a 0.3-inch hanging indent.
- Requires a rendered-page review before delivering the final Word file.

## Install in Codex

Ask Codex:

```text
Use $skill-installer to install this skill:
https://github.com/PonnyR/ama-reference-skill/tree/main/skills/format-ama-references
```

If the skill does not appear after installation, restart Codex.

## Use

Attach your reference-list `.docx` and ask:

```text
Use $format-ama-references to verify and format this reference list
according to current American Marketing Association journal style.
Preserve every source and report unresolved bibliographic details.
```

## Requirements

The full agent workflow needs web browsing, local file access, Python 3.10 or later, `python-docx`, and a compatible **documents skill** that can render DOCX files and inspect every output page. The documents skill and its renderer are environment-provided dependencies and are not bundled in this repository. Enable them in your agent environment before running the complete workflow.

The included formatter can also run separately:

```bash
python -m pip install -r requirements.txt
python skills/format-ama-references/scripts/format_reference_docx.py \
  --input source.docx \
  --references finalized_references.json \
  --output formatted_references.docx
```

`finalized_references.json` must contain already verified and ordered entries. See the [skill instructions](skills/format-ama-references/SKILL.md) for the JSON format. The Python formatter applies layout; bibliographic research and sorting are performed by the agent. Inspect the generated document before using it.

## 中文说明

这是用于 **美国市场营销协会期刊参考文献格式** 的 Codex skill，适用于 Journal of Marketing、Journal of Marketing Research 等营销期刊。

安装后，附上 Word 参考文献列表，输入：

```text
请使用 $format-ama-references，核对并按当前 AMA 营销期刊格式整理这份参考文献。
保留所有来源，说明元数据修正和未能核实的信息。
```

完整流程需要联网检索、Python、`python-docx`，以及环境中可用的 documents skill 和 Word 页面渲染功能。本仓库没有附带 documents skill。单独运行排版脚本时，需要先提供核实并排序后的 JSON 文献列表，再人工检查输出。

## Style sources

- [AMA journal submission guidelines](https://www.ama.org/submission-guidelines-american-marketing-association-journals/)
- [AMA reference-style examples](https://www.ama.org/american-marketing-association-journals-reference-style-examples/)

The current official AMA guidance controls reference style. This is an independent community skill.

## License

[MIT](LICENSE). Copyright 2026 Zeming Ren. The license covers the repository files; linked publications and external resources retain their own terms.
