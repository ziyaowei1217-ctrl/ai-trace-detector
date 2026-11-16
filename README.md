# AI 痕迹检测 · ai-trace-detector

检查 Word、PowerPoint、Excel、PDF 文件留下的技术痕迹，判断它是否由程序或 AI 工具生成、加工。看的是文件，不是文风。Claude 和 Codex / ChatGPT 通用。

A skill for Claude and Codex / ChatGPT that scans Office and PDF files for AI-generation traces in metadata and structure (not writing style). The skill instructions and reports are in Chinese.

## 能做什么

- **审核**别人交来的文件：列出疑点、证据强弱、正常解释，以及建议核实的问题。
- **自查**自己要对外发的文件：列出会随文件带出去的名字、路径、批注、隐藏内容，以及收件人能看到的疑点。

## 检查的 8 类

| 类别 | 例子 |
|---|---|
| 作者 | 作者是 `python-docx`、`Un-named`、`openpyxl` 等库名；与预期作者或机构对不上 |
| 路径 | 残留 `/mnt/data/…` 等 AI 沙盒路径；本机用户名路径 |
| 时间 | 库模板的固定时间戳；创建晚于修改；零编辑时间；属性时间与真实打包时间不符 |
| 批注与修订 | 批注早于创建或晚于保存；多条批注时间完全相同；署名是 AI 产品名 |
| 文件结构 | 压缩包时间戳；没有编辑会话标识；Excel 公式无缓存结果；PDF 生成库指纹 |
| 正文残留 | `utm_source=chatgpt.com`；引用占位符；AI 对话套话；未渲染的 Markdown |
| 官方标识 | 国标 AIGC 隐式标识；C2PA 内容凭证 |
| 嵌入对象 | 嵌入文件自带的作者；AI 绘图工具的图片文件名；备注和隐藏内容 |

每项的原理、可靠程度和常见误报见 [checks.md](plugins/ai-trace-detector/skills/ai-trace-check/references/checks.md)。

## 安装

**Claude Code**

```
/plugin marketplace add ziyaowei1217-ctrl/ai-trace-detector
/plugin install ai-trace-detector@ai-trace-detector
```

**Codex**

```
codex plugin marketplace add ziyaowei1217-ctrl/ai-trace-detector
```

然后在插件目录里安装 AI Trace Detector。

**其他支持 Agent Skills 的工具**

```
npx skills add ziyaowei1217-ctrl/ai-trace-detector --skill ai-trace-check
```

或者手动把 `plugins/ai-trace-detector/skills/ai-trace-check` 文件夹复制到工具的技能目录。

## 用法

装好后上传文件并说：

- "帮我查一下这个文件有没有 AI 生成的痕迹，作者应该是张敏"
- "这份报告要发给客户，先帮我自查一遍"

也可以单独运行脚本（Python 3.8 以上，只用标准库）：

```bash
S=plugins/ai-trace-detector/skills/ai-trace-check/scripts/ai_trace_scan.py
python $S 文件或文件夹 --mode review --expect-author "张敏" --expect-org "某某资本"
python $S 报告.docx --mode self
python $S 收件箱/ --format json --out 结果.json
```

PDF 的批注和正文检查需要 `pip install pypdf`。

## 局限

- 强证据说明文件或其中内容经过了程序或 AI 工具，不等于全文由 AI 撰写。
- 没有痕迹不等于没用 AI：把 AI 文字手动粘进正常的 Word 文件，基本查不出。
- 另存为、WPS、在线文档导出都会造成和程序生成相似的痕迹，报告里每条疑点都附了正常解释。
- Claude 等模型的统计型文本水印无法用本工具检测；C2PA 只做标记识别，不验证签名。
- 目前只用代码库生成的样本测试过，真实 Word、WPS、Google 文档导出文件上的误报率还没有评估。欢迎提 issue 反馈误报。
- 脚本只读，不修改文件，也不提供清除痕迹的功能。
- 指纹库会过时：AI 工具换了生成库或默认值，对应检查就会失效。

## 许可证

MIT
