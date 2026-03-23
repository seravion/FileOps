# FileOps - 技术框架与架构设计

> 版本：v0.1（草案）  
> 日期：2026-03-23  
> 目标：给出可落地的模块拆分、核心数据模型与执行流程，指导后续实现。

## 1. 形态与技术选型建议
### 1.1 交付形态
- **CLI**：面向一线使用者（本地/CI/定时任务）。
- **Library/SDK**：面向二次开发（脚本/服务调用同一套核心能力）。
- **Plugin**：面向扩展（操作/过滤/报告输出）。

### 1.2 语言与生态（建议）
在未给出硬性约束前，推荐按团队偏好选择其一：
- **Python 3.11+**：文件处理生态强、迭代快、插件机制成熟（entrypoints）。
- **Go**：单文件分发、性能好、跨平台稳定（但插件机制相对重）。
- **Node.js**：前端/全栈团队顺手，生态多，但对底层文件语义需更谨慎。

本文后续以“**Python 风格分层**”表达架构思想（不绑定具体语言）。

## 2. 架构原则
- **Plan-first**：先生成计划，再执行；计划可导出、可审查、可复跑。
- **核心纯逻辑 + 边界适配**：扫描/过滤/计划构建尽量纯函数化；I/O 通过接口注入。
- **安全默认**：默认不覆盖、不硬删、不越界；高风险操作强制二次确认或显式 `--force`。
- **可观测**：结构化日志 + 结果事件流（便于进度、报告、调试）。

## 3. 高层模块划分
建议按以下包/模块组织（名称可调整）：
- `core/`：领域模型（Plan、Operation、Result、Rule）与纯逻辑
- `scan/`：文件枚举与元数据采集（Scanner、StatProvider）
- `filter/`：过滤器与规则组合（glob/regex/size/mtime/...）
- `ops/`：具体操作实现（copy/move/rename/delete/...）
- `engine/`：执行引擎（并发、重试、恢复、事件）
- `report/`：报告聚合与导出（JSON/CSV/HTML）
- `config/`：配置加载与校验（schema、默认值、变量插值）
- `cli/`：命令行解析与 UX（进度条、确认交互、输出摘要）
- `plugins/`：插件加载（本地目录/entrypoints）与扩展接口

## 4. 核心数据模型（建议）
### 4.1 Entry（文件条目）
最小字段：
- `path`：绝对路径或工作区相对路径（推荐内部统一为绝对）
- `type`：file/dir/symlink
- `size`、`mtime`、`mode`（可选）
- `content_hash`（可选，昂贵，按需计算）

### 4.2 Operation（操作）
统一为：
- `kind`：copy/move/rename/delete/...
- `src`：源路径（或 Entry 引用）
- `dst`：目标路径（可为空，如 delete）
- `options`：覆盖策略、保留结构、trash/hard 等
- `risk_level`：low/medium/high（用于 CLI 提示与策略限制）

### 4.3 Plan（计划）
- `workspace`：允许操作范围（root、allowlist/denylist）
- `operations[]`：Operation 列表（顺序执行或可并行）
- `summary`：统计信息（命中数、潜在覆盖数、预计写入量）
- `created_at`、`tool_version`、`config_fingerprint`（用于复现）

### 4.4 Result / Event（结果与事件）
执行过程中产生事件流，既用于进度展示，也用于报告生成：
- `operation_id`
- `status`：success/skipped/failed
- `error_code`、`error_message`、`exception_type`（如适用）
- `duration_ms`
- `bytes_written`（如适用）

## 5. 端到端流程（推荐）
1) **LoadConfig**：读取配置 + 校验（schema）+ 展开变量
2) **Scan**：扫描目录，产出 Entry 流（不要一次性全载入内存）
3) **Filter**：应用规则，得到命中 Entry 流
4) **BuildPlan**：把 Entry 映射为 Operation（含冲突检测、目标路径生成）
5) **PresentPlan**：打印摘要 + 关键风险提示；可导出计划文件
6) **Execute**：执行引擎按计划执行（并发、重试、恢复）
7) **Report**：聚合结果并导出（JSON/CSV/HTML）

## 6. 执行引擎设计要点
### 6.1 并发模型
- 扫描与过滤：天然流式，可单线程 + 背压
- 执行：I/O 密集适合 worker pool；保持“同一目的路径”的写入串行（避免竞争）

### 6.2 幂等与恢复（resume）
- 每个 Operation 都有稳定 `operation_id`（由 src/dst/kind/options 哈希生成）。
- 执行时写入 `state`（例如：`state.jsonl`）：
  - 已成功的 operation_id 可跳过
  - 失败可重试或输出失败清单
- 对跨盘 move：先 copy 再 delete，过程中记录阶段（避免中断造成双份/丢失不明）。

### 6.3 冲突与覆盖策略
在 BuildPlan 阶段尽量发现问题：
- 目标已存在：按策略（fail/skip/overwrite/auto_rename）
- 多源指向同一 dst：必须解决（fail 或 auto_rename）
- src==dst：直接 skip

## 7. 配置与 Schema（建议）
建议配置支持：
- `workspace`：允许/禁止路径、是否跟随 symlink、危险操作开关
- `scan`：roots、max_depth、include_hidden、ignore_patterns
- `filters`：规则列表（AND/OR 组合）
- `operation`：kind 与参数（copy/move/rename/delete 等）
- `execution`：dry_run、concurrency、resume、retry
- `output`：log_format、report_path、report_format

实现建议：
- 配置加载与校验应独立于 CLI，供 SDK 复用。
- 若使用 YAML/TOML，务必提供“严格模式”防止拼写错误悄悄被忽略。

## 8. 插件接口（建议）
最小接口（伪代码概念）：
- `OperationFactory`：把 Entry + config 变成 Operation（或直接提供新 kind）
- `OperationHandler`：执行某 kind 的 Operation
- `FilterProvider`：新增过滤器（如 exif、mime、content）
- `Reporter`：订阅事件流，产出自定义报告/通知

插件加载：
- 本地 `plugins/` 目录扫描（开发期简单）
- 发布后可扩展为“包 entrypoints/registry”

## 9. 错误模型（建议）
错误应可归类、可机器解析：
- `E_PERMISSION` 权限不足
- `E_NOT_FOUND` 源不存在
- `E_CONFLICT` 目标冲突
- `E_IO` 通用 I/O 错误
- `E_POLICY` 被安全策略阻止（越界/危险目录/硬删未允许）
- `E_UNSUPPORTED` 平台不支持（例如回收站实现）

CLI 输出对用户友好，报告中保持结构化字段便于追踪。

## 10. 测试策略（建议）
- 单元测试：规则组合、目标路径生成、冲突检测、计划构建
- 集成测试：在临时目录构造文件树，跑典型命令（dry-run 与实际执行）
- 兼容性测试：Windows 路径（长路径、Unicode、只读/锁定文件）

## 11. 仓库结构建议（落地版）
若后续进入实现阶段，可按以下最小骨架推进：
- `src/fileops/`：库代码
- `src/fileops_cli/` 或 `cli/`：CLI 入口
- `tests/`：pytest/go test/jest（按语言）
- `examples/`：示例配置与示例目录结构
- `docs/`：持续维护 PRD 与架构
