# 裁员维权助手

`yuanye-layoff-rights` 是一个面向中国大陆劳动合同制员工及劳务派遣员工的 Agent Skill。通过对话梳理劳动争议、提供具体法律依据、计算相关费用，并按需要生成可编辑的 Word 诉求文书。

机关、事业单位中签订劳动合同的员工属于分析范围；公务员、事业编制人事争议暂未覆盖。

## 根据你的需要开始

Skill 会先了解你想解决什么，已经说清楚的内容不重复询问。

| 需要 | 交付 |
| --- | --- |
| 单位已提出解聘或已经离职，想争取权益 | 争议分析、费用明细、给单位的 Word 文书 |
| 尚未被裁，想提前了解风险 | 风险及已发生问题的分析、法律依据与沟通方向 |
| 暂不考虑离职，想检查单位侵犯权益的问题 | 按相关话题排查并给出判断 |
| 只想估算离职补偿 | 按明确情景给出公式和金额，默认不生成 Word |

需要分析时，你可以先讲一段自己的经历，不必填完一张问卷。Agent 会提取信息，再按话题补问关键缺口；宿主支持点选控件时使用控件，否则使用普通对话。

## 在 Codex 中安装

仓库根目录就是 Skill 目录。在终端执行：

```bash
git clone https://github.com/yuanye1024/yuanye-layoff-rights.git ~/.codex/skills/yuanye-layoff-rights
```

如果目标目录已经存在，先检查已有版本，不要直接覆盖。安装后在新对话中使用 `$yuanye-layoff-rights`，例如：

> $yuanye-layoff-rights 我还没有收到裁员通知，只想按年底离职估算一下补偿，不需要生成文档。

> $yuanye-layoff-rights 单位说我考核不合格，让我签个人辞职申请。我想看看有没有违法，并准备一份给单位的诉求文书。

其他支持本地 Skill 的 Agent 可将整个目录放入其 Skill 目录，或读取 [SKILL.md](SKILL.md) 后按引用加载需要的资料。

## 文件与法律来源

```text
SKILL.md                 入口、分流和交互约束
agents/openai.yaml       界面名称与调用提示
references/              提问、判断、计算、法源和文书格式
assets/laws/             三份指定法律原文快照
assets/report-template.docx  通用 Word 格式模板
scripts/build_report.py  将已审核正文排版为 DOCX
```

三份原文分别是《中华人民共和国劳动合同法》《中华人民共和国劳动合同法实施条例》和《最高人民法院关于审理劳动争议案件适用法律问题的解释（一）》。版本、快照哈希、效力注意事项及允许使用的官方补充来源见 [法律来源](references/legal-sources.md)。先读包内原文，缺规则或核验效力时再查官方来源；地区标准按对应地区和期间查询。

Word 模板仅保留格式和段落角色，不含作者的个人案例、姓名、证据或图片。完整文书按“适用法律条例／相关争议问题及条例／本人诉求及相关条例依据”组织，先写问题或诉求，再附对应具体条款原文。法律引用使用浅灰底色及斜体，与正文区分。

## Word 生成环境

对话分析及估算不需要执行 Word 脚本。生成 DOCX 时需要 Python 3.10+ 和 `lxml`；宿主已有相应运行环境时直接复用。需要安装依赖时可在仓库内的虚拟环境中执行：

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python scripts/build_report.py --content /path/to/reviewed-content.json --output /path/to/report.docx
```

脚本默认读取包内模板，只负责排版，不负责法律判断或金额计算。输入正文应由 Agent 依据 [文书规范](references/report-format.md) 审核；交付前使用宿主可用的 Word/PDF 渲染工具逐页检查。输出文书和用户材料保存在仓库之外。

## 当前验证范围

已完成 Skill 基础结构、包内文件依赖、法律快照一致性和模板生成校验。通用行为回归规格见 [回归案例](references/regression-cases.md)；独立多案例对话回归尚未完成，地区政策和法律效力仍需在个案中核验。

本仓库没有指定开源许可证。公开可访问不等于自动授予任意修改、再分发或商业使用许可；需要授权时请联系作者。
