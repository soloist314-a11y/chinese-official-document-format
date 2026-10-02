# 公文格式处理

将用户提供的内容排成可编辑的 Word 公文，调整已有 DOCX 的格式，或检查公文格式。技能按相关请求自动选择，也可明确使用 `$chinese-official-document-format`。

## 格式依据与范围

规则来源于用户提供的《公文处理格式要求》第 19–22 页，已提取为匿名化格式规则。它是一套常用公文配置，不宣称等同于所有机关的现行统一标准。

支持日常公文、红头文件、会议纪要和函件的格式指导，包括页面、字体、标题层级、正文、主送机关、表格、附件、落款、日期和版记。当次用户要求与当次原生模板优先；仅要求排版时保留正文、顺序和业务数据。

公开包不包含原始 PPTX、截图、字体文件、来源单位、人员、地址、电话、网站、商标或内部审批流程。正式文稿中的实际业务信息由本次用户输入，不套用示例名称，也不默认匿名化本次正文。

## 安装与自动调用


在支持本地插件的 ChatGPT/Codex 环境中，可通过此仓库的个人 marketplace 安装：

```sh
codex plugin marketplace add soloist314-a11y/chinese-official-document-format
codex plugin add chinese-official-document-format@soloist314-skills
```

安装并成功加载后，开始新聊天，发送“按我的公文格式排版”“检查这个 Word 公文格式”等相关请求。技能配置已开启 `allow_implicit_invocation: true`，模型可根据请求自动选择；需要明确调用时可写 `$chinese-official-document-format`。

安装完成须在目标应用的新聊天中核对技能实际可用。仓库、本地安装路径或 GitHub 发布状态本身不证明 ChatGPT 已加载技能。本地插件不会仅因安装在 Mac 上就自动同步到云端、其他设备或手机。

技能路径：

```text
plugins/chinese-official-document-format/skills/chinese-official-document-format/
```

Marketplace 清单：`.agents/plugins/marketplace.json`。

## 日常公文脚本

`scripts/format_docx.py` 使用 Python 3.8+ 标准库，无需安装 `python-docx`。它从明确标注段落角色的 UTF-8 JSON 新建日常公文，保留手工编号、标点、换行和制表符，不覆盖已有输出文件。

进入技能目录后：

```sh
python3 scripts/format_docx.py --help
python3 scripts/format_docx.py create examples/daily-document.json document.docx --report document-check.json
python3 scripts/format_docx.py check document.docx --report document-check.json
python3 scripts/test_format_docx.py
```

输入字段见 [结构化输入说明](plugins/chinese-official-document-format/skills/chinese-official-document-format/examples/input-schema.md)，完整能力边界见 [脚本说明](plugins/chinese-official-document-format/skills/chinese-official-document-format/references/script-usage.md)，格式细节见 [格式规则](plugins/chinese-official-document-format/skills/chinese-official-document-format/references/format-rules.md)。

生成器仅新建日常公文；红头文件、会议纪要和固定函件应采用当次原生模板或其他能够实现其要求的编辑方式。已有 DOCX 应另存编辑，不通过生成器重新导入覆盖。

## 验证与交付边界

现有 7 项回归检查已通过，包括文本与结构保留、错误格式识别、元数据诊断、已有输出保护、正式文种拒绝、不合法输入拒绝和未知语义角色拒绝。示例 DOCX 已生成并重新打开做结构检查，结果通过。

这些结果仅证明 OOXML 结构。脚本不验证字体是否安装、字体替代、裁切、Word/WPS 实际分页、附件与落款日期是否独立成页或跨页表头表现。正式交付仍需用可用的 Word、WPS 或 LibreOffice 渲染，并逐页检查；未完成渲染时应说明“已检查结构，尚未验证分页”。本次发布未把视觉分页验收列为已完成。

`check --anonymous-metadata` 只将作者、单位、批注和修订等元数据列为错误，不删除信息，也不检查正文中的单位名称或个人信息。匿名副本须由用户明确要求并另行检查。
