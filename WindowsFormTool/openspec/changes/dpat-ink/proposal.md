## Why

现有工具仅支持 TSK 合并与通用 INK 规则，缺少基于 DPAT CSV 测试数据的图谱异常筛选与自动 INK 能力。新增 DPAT INK 可把测试项异常直接映射回 TSK，提升批量处理效率与一致性。

## What Changes

- 在主界面下拉功能中新增 “DPAT INK” 模式，触发输入弹窗（testName 选择、上下限与 sigma、Ink Bin）
- 参数弹窗支持多测试项选择与配置，已选测试项列表可回显与调整各自配置
- 支持一次导入不超过 25 个 TSK 文件，并匹配对应 CSV（CSV 文件名包含 TSK 文件名，格式参考 `docs/Csv格式.csv`）
- 读取 CSV 中 `Test Name` 行与数据区（包含 X/Y、Bin、测试项列），仅对 Bin=1 的 pass die 参与统计
- 依据输入的上下限进行初筛，再用统计公式计算新上下限并标记越界点
- 将越界坐标写回对应 TSK，将 die 属性 INK 为 Fail，Bin 改为输入的 Ink Bin
- 结果文件输出到 `DPAT_INK` 文件夹，文件名追加后缀 `_DPAT`

## Capabilities

### New Capabilities
- `dpat-ink`: 基于 DPAT CSV 的批量 INK 处理（上传 TSK/CSV、testName 选择、统计阈值计算、回写 TSK 并输出）

### Modified Capabilities
- （无）

## Impact

- UI：`Form1` 下拉选项新增 DPAT INK；新增参数输入弹窗（testName、上下限、sigma、Ink Bin）
- UI：参数弹窗新增多测试项配置列表，支持多次选择并查看/调整已选项配置
- 处理层：新增 DPAT INK 处理器（类似 `TskInkProcessor` / `TskMergeProcessor`）
- 处理层：按多测试项逐一统计并合并越界点后写回 TSK
- CSV 解析：读取 `docs/Csv格式.csv` 所示结构（Test Name 行、数据区列）
- 文件输出：新建 `DPAT_INK` 输出目录；TSK 文件名追加 `_DPAT`
- 规范与数据：保持 `TSK_spec_2013.pdf` 格式兼容与数据完整性
