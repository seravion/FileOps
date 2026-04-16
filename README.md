# FileOps

FileOps 是一个面向 Windows 的文档处理工具，提供 GUI 一键操作与 CLI 批处理两种模式。

当前版本重点解决三类场景：
- 文档按结构拆分（避免章节内容串段）
- 文档严格套用模板排版
- 文档与模板格式对照并输出可执行整改报告

---

## 1. 核心能力

### GUI（图形界面）
GUI 当前提供 4 个主功能：

1. `按大小拆分`
   - 按指定 MB 大小切分文件
   - 支持常规文件，PDF / DOCX 会走更合适的分片策略

2. `文档拆分`
   - 支持 `DOCX / Markdown / TXT / PDF`
   - 支持按 `一级标题 / 二级标题 / 一级+二级` 拆分
   - 优化章节边界，尽量保证每个分片只包含目标章节内容
   - 可选 OCR（提取图片文字）

3. `文档一键排版`
   - 输入：模板 `.docx` + 待处理 `.docx`
   - 输出：`*_formatted.docx`
   - 重点对齐模板中的标题、正文、目录、表格等样式

4. `文档模板对照`
   - 输入：模板 `.docx` + 待检查 `.docx`
   - 输出：
     - `*_compare_report.docx`
     - `*_compare_report.json`
     - `*_adjusted.docx`
   - 检查项包括：段落样式、缩进、行间距、公式编号、图编号、参考文献格式等

### AI 辅助（已覆盖全部 GUI 功能）
- AI 辅助不再只限“模板对照”，可用于所有 GUI 功能
- 界面中只需选择服务商与模型，不需要手填接口地址（已内置）
- 当前内置服务商：
  - ChatGPT（OpenAI）
  - DeepSeek
  - 智谱 GLM
  - Claude（Anthropic）
  - Kimi（Moonshot）
- 开启后会生成 AI 建议文件（Markdown）：
  - 模板对照：`*_ai_assist.md`
  - 其他操作：`*_{operation}_ai_assist.md`

> 说明：使用 AI 功能仍需填写对应服务商可用的 `API Key`。

### CLI（命令行）
CLI 主要覆盖通用文件操作：
- `copy`
- `move`
- `rename`
- `delete`

---

## 2. 快速开始

### 环境要求
- Python `>= 3.11`
- Windows（GUI 基于 PySide6）

### 安装依赖
```powershell
python -m venv .venv
.\.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements-dev.txt
pip install .
```

### 启动 GUI
```powershell
python scripts/entrypoint.py
```

---

## 3. GUI 使用说明

1. 选择 `操作类型`
2. 选择 `工作区`（安全范围）
3. 添加源文件
4. 配置目标目录 / 模板 / 拆分规则等参数
5. 如需 AI，勾选“启用AI辅助（全功能）”，选择服务商和模型，填写 API Key
6. 点击“开始执行”

执行过程中可查看：
- 进度条
- 执行日志
- 报告输出路径

---

## 4. 报告与输出

### 模板对照
- 结构化 JSON 报告
- 可读 DOCX 报告（差异 + 建议）
- 自动调整后的文档
- 可选 AI 修订建议（Markdown）

### 其他操作
- 常规运行报告（通过界面“报告文件”配置）
- 可选 AI 复盘建议（Markdown）

---

## 5. 测试

```powershell
pytest
```

---

## 6. 打包

### 构建 EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```
输出：`dist/fileops.exe`

### 构建安装包
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```
输出：`dist/FileOps-Setup.exe`

---

## 7. 补充说明

### OCR
- OCR 依赖：`pytesseract` + 系统安装的 Tesseract OCR
- 未安装 Tesseract 时，不影响主流程，但图片文字提取能力受限

### PDF
- 优先按目录（书签）进行分段
- 无目录时回退为文本标题规则
- 加密 PDF 需先解除限制再处理

---

## 8. 项目结构（简要）

```text
src/fileops/
  gui.py                # 图形界面与执行调度
  document_split.py     # 文档拆分
  word_template.py      # 模板排版
  document_compare.py   # 模板对照
  ai_assistant.py       # AI服务商、模型与建议生成
  operations.py         # 通用文件操作
  cli.py                # 命令行入口
```

