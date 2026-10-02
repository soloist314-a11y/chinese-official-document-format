# 日常公文 JSON 输入

执行：

```sh
python3 scripts/format_docx.py create examples/daily-document.json /tmp/daily-document.docx --report /tmp/daily-document-check.json
python3 scripts/format_docx.py check /tmp/daily-document.docx
```

根对象包含 `mode: "daily"` 和非空 `blocks` 数组。文本段落必须明确提供 `role` 与完整 `text`。脚本不改写文字，不补写发文机关、标题序号、日期、主送机关或业务事实。日期由用户提供，不默认使用当前日期。

| role | 用途 | 文本要求 |
|---|---|---|
| title | 主标题 | 完整标题 |
| body | 正文 | 完整正文 |
| h1–h5 | 五级标题 | 在 `text` 中手动写出序号；不使用 Word 自动编号 |
| recipient | 主送机关 | 实际机关名称，最后使用全角冒号；不可虚构 |
| table_title | 表格外标题 | 完整标题 |
| attachment_item | 附件列表一项 | 完整文字和手工序号；附件名称后不加标点 |
| appendix_marker | 附件页标识 | `附件` 或手工输入 `附件1` 等；独立起页 |
| signature | 发文机关署名 | 仅在用户提供真实署名时使用 |
| date | 成文日期 | 仅在用户提供日期时使用 |
| blank | 空白段落 | 不要求 `text`；一行29磅 |
| page_break | 手动分页 | 不要求 `text`；一般优先使用 appendix_marker |
| table | 表格 | 矩形字符串二维数组 `rows`，至少一行重复表头 |
| attachment_list | 附件列表分组 | 非空 `items`，每项字符串或附件项对象 |

附件项可提供 `continuation_indent_chars`：表示回行文字距离版心左侧的字符数，须按实际标签宽度设定，例如 `附件：工作清单` 为5字符（2字符缩进＋3字符标签）。多项附件各行的手工标签应由编写者明确输入，不自动生成序号。

署名可提供 `align_with_date`，其值为紧随其后的日期段文字，用于估算署名与日期的中心对齐。此估算按汉字和半角字符宽度计算，需在 Word/WPS 中核对。也可提供 `right_indent_chars` 明确设定署名右缩进。日期右空4字。

表格可设定 `font_size_pt: 12` 或 `10.5`、`line_pt: 14` 至 `18`、`header_rows: 1` 或更多。表格使用自适应布局，不设固定行高，表格上下各空半行。跨页效果仍须渲染检查。

脚本仅创建“其他日常公文”模式。红头文件、会议纪要、对外函件使用当次用户提供或认可的原生模板，不能把首页专用上边距伪装为全篇页边距或用空行代替。当次业务模板的实际单位信息可以保留，不写入公开技能包。日常公文不包含版记，因此脚本不接受 `imprint`。

结构检查仅校验 OOXML。它不会确认字体已安装、Word/WPS 实际分页、字形替代、附件列表或署名是否独立成页；这些必须渲染为 PDF/页面图像后逐页核查。当前环境无法渲染时，应报告“结构检查通过，视觉校验未完成”。
