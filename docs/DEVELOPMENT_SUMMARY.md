# FileOps 开发总结（便于后续迭代）

> 更新时间：2026-03-24  
> 基线提交：`d697e76`（`main`）

## 1. 项目现状

FileOps 目前是一个以 GUI 为主的桌面文件处理工具，核心能力已覆盖：

- 文件操作：复制 / 移动 / 重命名 / 删除
- 文件按大小拆分
- 文档按标题拆分（一级、二级、一级+二级）
- 文档拆分导入格式可选（自动 / DOCX / Markdown / TXT）
- 文档拆分导出格式可选（原格式 / DOCX / Markdown / TXT）
- OCR 图片文字提取（可选）
- Dry Run（预演模式）
- 进度条、日志、JSON 报告
- Windows `exe` 与安装包脚本

## 2. 文档拆分能力矩阵（当前）

### 2.1 输入支持

- `docx`：按 Word 标题样式拆分
- `md/markdown`：按 `#` 标题拆分
- `txt`：按“行内容”模式拆分（无标题时会生成单分段）

### 2.2 输出支持

- 输入 `docx`：
  - 导出 `docx`：优先保留原格式、段落、表格、图片结构
  - 导出 `md/txt`：文本化输出；标题会尽量映射，表格会转 markdown/制表文本
- 输入 `md/txt`：
  - 导出 `md/txt`：文本分段输出
  - 导出 `docx`：文本重建为 Word（支持基础标题与 markdown 表格重建）

### 2.3 关键限制（请注意）

- `md/txt -> docx` 属于“重建”，不可能保留原 Word 样式细节
- `docx -> md/txt` 属于“降级导出”，复杂排版会损失
- OCR 依赖本机 Tesseract，不安装时不会阻塞主流程，但 OCR 结果为空

## 3. 核心代码结构（高频维护入口）

- `src/fileops/gui.py`：主界面、参数收集、线程调度、进度与日志
- `src/fileops/document_split.py`：文档拆分与格式转换核心
- `src/fileops/operations.py`：复制/移动/重命名/删除/按大小拆分
- `src/fileops/reporting.py`：报告输出
- `src/fileops/models.py`：结果模型、状态枚举
- `tests/test_operations.py`：核心回归测试（含文档拆分新能力）

## 4. 执行链路（便于排障）

1. GUI 收集参数（含工作区、输入/输出格式、拆分模式）
2. `OperationWorker` 逐条处理源文件
3. 根据操作路由到 `operations.py` 或 `document_split.py`
4. 每条结果统一写日志 + 进度更新
5. 汇总生成 `RunReport`，可落盘 JSON

排障建议优先看：

- GUI 日志区（第一现场）
- 报告文件（批量任务）
- `document_split.py` 的输入格式校验与异常信息

## 5. 开发与验证命令

### 5.1 本地开发

```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -U pip
pip install . -r requirements-dev.txt
python scripts/entrypoint.py
```

### 5.2 回归测试

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

### 5.3 打包

```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```

## 6. 后续迭代建议（建议优先级）

### P1（建议先做）

- 增加“格式不兼容”更友好的前置提示（例如自动给出推荐输入格式）
- 给文档拆分新增“输出目录预览 + 预计生成文件数”
- 为 `docx -> md/txt` 增加更多结构映射（列表、引用、代码块）

### P2（稳定性增强）

- 补充异常样本测试（损坏文档、超长路径、特殊字符文件名）
- 为大文档处理增加更细粒度进度（按段/按块）
- 增加“失败重试”与“跳过后继续”策略配置

### P3（体验优化）

- 增加最近使用配置（workspace、模式、格式）
- 增加任务历史记录视图
- 增加多任务队列

## 7. 发布前自检清单

- GUI 关键流程手测（每种操作至少 1 次）
- `pytest` 全绿
- 打包后 `dist/fileops.exe` 可启动
- 安装包可安装并正常运行
- README 与 docs 同步更新

---

如果你后续想扩“文档合并”“按规则批量重命名模板库”“拖拽导入”等功能，建议直接在本文件追加“变更记录”段，长期作为迭代主索引。
