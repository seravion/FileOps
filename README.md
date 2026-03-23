# FileOps

FileOps 是一个桌面文件处理工具，支持图形界面点击操作。

## 功能
- 复制、移动、重命名、删除
- 按大小拆分文件（MB）
- 文档拆分（按一级标题、二级标题、一级+二级标题）
- 可选提取图片文字（OCR）
- 预演模式（Dry Run）
- JSON 报告导出

## 图形界面说明
- 操作类型选择：在“操作类型”下拉中选择对应功能
- 文档拆分：
  - 选择“文档拆分”
  - 添加 `.docx` / `.md` / `.txt` 文档
  - 设置输出目录
  - 设置“标题拆分规则”（一级、二级、一级+二级）
  - 勾选“提取图片文字（OCR）”可尝试识别文档中的图片文字
- 按大小拆分：
  - 选择“按大小拆分”
  - 设置“分片大小(MB)”

## 执行进度怎么看
- 点击“开始执行”后，会立即显示：
  - 顶部状态：`执行中...`
  - 进度条：`进度：x/y（%）`
  - 执行日志：逐项实时输出“开始处理/成功/失败”
- 只在“永久删除”时弹确认框，其它操作不再阻塞确认。
- 如果日志里出现 OCR 相关报错，通常是本机未安装 Tesseract OCR，可先关闭“提取图片文字（OCR）”再执行。

## 运行与打包
### 本地运行
```powershell
python -m venv .venv
.\.venv\Scripts\activate
pip install -U pip
pip install . -r requirements-dev.txt
python scripts/entrypoint.py
```

### 运行测试
```powershell
pytest
```

### 构建 GUI EXE
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_exe.ps1
```
输出：`dist/fileops.exe`

### 构建安装包
```powershell
powershell -ExecutionPolicy Bypass -File scripts/build_installer.ps1
```
输出：`dist/FileOps-Setup.exe`

## OCR 说明
- 图片文字识别依赖 `pytesseract` 与本机 Tesseract OCR。
- 如未安装 Tesseract，文档拆分仍可执行，但图片文字可能无法识别。
