# dsh-skill-project-docs

给 DeepSeek Harness 用的**软件研发全流程项目管理过程文档**技能包：安装后在会话技能目录中提供 `software-project-docs`，用于生成与修订 18 类项目管理过程文档（DOCX）。

## 它做什么

- **先确认再生成**：每次任务前确认要生成哪些文档、正式版还是模板、新建还是修订。
- **统一排版**：A4、宋体小四正文、黑体分级标题、1.5 倍行距、首行缩进 2 字符、三线/全框表格、表上图下、分节页码、自动目录、正式修订记录表。
- **可复现构建**：内容写在受控 Markdown 里，由包内 `docx_kit.py` 渲染为 DOCX；同一内容源重复构建得到同一排版。
- **模板模式**：每章配楷体斜体的【填写说明】与示例，便于他人填写。
- **版本化修订**：结构调整升主版本、内容修订升次版本，每次产出新文件并追加修订记录，旧版保留。

## 安装

装到 DSH 的插件商城（Discover）后点安装，或在 **Plugins → Add plugin** 选择本包目录。

安装后无需启用任何设置：技能随包注册，新开会话即可在技能目录中看到 `software-project-docs`。

## 使用

直接用自然语言触发，例如：

> 用软件研发项目管理文档技能，给「智慧园区综合管理平台」生成需求规格说明书和测试计划。

技能会先问你文档清单、模式（正式/模板）、版本与项目信息，然后逐份生成并自检。

## 包含的文件

| 路径 | 用途 |
|---|---|
| `skills/software-project-docs/SKILL.md` | 技能主流程与硬性规则 |
| `skills/software-project-docs/references/` | 排版规范、内容边界、版本规则、18 类文档章节骨架 |
| `skills/software-project-docs/assets/` | DOCX 构建库 `docx_kit.py`、内容源语法说明与样例 |
| `skills/software-project-docs/scripts/` | `build.py` 构建命令、`check_docx.py` 规范自检 |
| `index.js` | 把技能注册到 `ctx.skills` 的 Host 插件 |
| `cordis.patch.yml` | bundle 的 Loader 补丁 |

## 运行要求

- 技能生成 DOCX 需要 Python 与 `python-docx`；DSH 运行时自带的 Python 已包含，无需另装。
- 全流程说明以中文撰写，输出文档默认中文排版。

## 限制

- 不编造项目事实：资料缺失处写成 `【待确认】` 并登记到附录。
- 不做商务报价、法务条款与涉密等级判定。
- 图表使用你提供的图片或明确数据，不虚构数据。
