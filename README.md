# 高物讲义转课件 Skill

这是一个面向高中物理讲义转课件的 Codex skill，用于把高一/高二/高三物理讲义基于指定 PPT 模板拆解生成课堂课件。

它的重点不是把讲义整页截图塞进 PPT，而是把讲义内容转成适合授课的课件页面：目录、知识点、例题、练习、题图、公式和总结都按模板版式重新组织。

## 适用场景

- 高中物理讲义转 PPT 课件
- 基于已有模板生成课件
- 参考「讲义 + 成品课件」案例复刻转换规则
- 修正课件中的题号、目录、图片、公式、字号和版式问题
- 检查是否漏题、是否插入整页截图、公式是否和讲义原式一致

## 核心规则

- 复用模板原题号框、正文框、背景和字号，不自造白底板或新框。
- 题号按模板格式生成，例如 `P7-例4`、`P5-练2-1`。
- 目录页存在时默认保留，并映射本讲模块。
- 题图从讲义内嵌图片或局部题图提取，不用整页截图代替内容拆解。
- 公式必须保持讲义原式，不把 `2√3mg/3` 改成等价的 `2mg/√3`。
- 根号要带横线，分式要保持讲义中的上下分式形态。
- 完成前必须检查题目覆盖、风险文本和公式形态。

## 目录结构

```text
gaowu-ppt-courseware/
├── SKILL.md
├── agents/
│   └── openai.yaml
├── references/
│   ├── formula-rendering.md
│   └── qa-checklist.md
└── scripts/
    └── audit_pptx.py
```

## 安装到 Codex

把仓库克隆到本地 skills 目录：

```bash
cd ~/.codex/skills
git clone https://github.com/488589095-alt/gaowu-ppt-courseware.git
```

如果本地没有 `~/.codex/skills`，先创建：

```bash
mkdir -p ~/.codex/skills
```

安装后，在 Codex 中请求高中物理讲义转课件时即可触发这个 skill。

## 使用示例

```text
请使用 gaowu-ppt-courseware，把这个高二物理讲义第一讲基于这个 PPT 模板转成课件。
```

```text
用这个模板和参考案例，把讲义第二讲转成课件，注意不要漏练习题，公式要保持讲义原式。
```

## PPTX 反查脚本

`scripts/audit_pptx.py` 可用于检查生成后的 PPT 是否漏题号、是否残留风险文本。

准备一个题号清单 `labels.txt`：

```text
P4-例1
P5-练2-1
P5-练2-2
P7-例4
```

运行：

```bash
python3 scripts/audit_pptx.py 输出.pptx --expect-labels labels.txt
```

脚本会检查：

- PPT 页数
- 预期题号是否全部出现
- 是否残留 `sqrt`、`L1-`、`表达式A/B/C/D`、`...` 等风险文本
- PPT 中媒体和图片引用数量

注意：脚本只能检查文本结构。根号是否带横线、上下分式是否和讲义一致，仍需要结合截图或公式源清单复核。

## 关键文档

- [SKILL.md](SKILL.md)：skill 主流程和触发规则
- [公式渲染规则](references/formula-rendering.md)：根号、分式、原式一致性规则
- [验收清单](references/qa-checklist.md)：生成后检查项
