## ADDED Requirements

### Requirement: DPAT INK 模式入口与参数输入
系统 MUST 在主界面功能下拉中提供 “DPAT INK” 模式，并在选择后弹出参数输入窗口，包含 testName 选择、上限、下限、sigma 与 Ink Bin 的输入项（以及统计公式选择项，默认公式1）。

#### Scenario: 进入 DPAT INK
- **WHEN** 用户在下拉菜单中选择 “DPAT INK”
- **THEN** 系统弹出参数输入窗口并要求填写 testName、上下限、sigma 与 Ink Bin

### Requirement: 多测试项选择与配置
系统 MUST 支持一次选择多个测试项，并为每个测试项保存独立配置（上限、下限、sigma、公式、Ink Bin）。

#### Scenario: 添加多测试项配置
- **WHEN** 用户为一个测试项完成配置并确认添加
- **THEN** 系统将该测试项及其配置加入已选列表

#### Scenario: 回显已选项配置
- **WHEN** 用户继续选择第二个测试项
- **THEN** 系统在界面上展示已选列表及其配置，便于查看与修改

### Requirement: 批量文件选择与匹配
系统 MUST 支持一次选择不超过 25 个 TSK 文件，并按 “CSV 文件名包含 TSK 文件名（忽略扩展名与大小写）” 的规则匹配对应 CSV；若未匹配或出现多匹配 MUST 阻止执行并提示。

#### Scenario: 文件数量超限
- **WHEN** 用户选择的 TSK 文件数量超过 25 个
- **THEN** 系统提示超限并拒绝继续

#### Scenario: CSV 匹配失败
- **WHEN** 任一 TSK 文件无法匹配到唯一 CSV
- **THEN** 系统提示匹配错误并阻止执行

### Requirement: CSV 解析与测试项列表
系统 MUST 按 `docs/Csv格式.csv` 结构解析 CSV：定位 `Test Name` 行以获取测试项列表，并从数据区头部（包含 `X`,`Y`,`Bin` 列）开始解析每行测试数据。

#### Scenario: 读取测试项
- **WHEN** 用户加载 CSV
- **THEN** 系统从 `Test Name` 行解析并显示可选测试项列表

### Requirement: 数据筛选与上下限计算
系统 MUST 仅使用 Bin=1 的数据点参与统计；在用户提供初始上下限时，先对测试值进行区间筛选，再按选定公式计算新上下限。

#### Scenario: 公式1（均值/标准差）
- **WHEN** 用户选择公式1
- **THEN** 系统计算均值与标准差，并以 `新上限=均值+sigma*标准差`、`新下限=均值-sigma*标准差` 得到新阈值

#### Scenario: 公式2（中位数/IQR）
- **WHEN** 用户选择公式2
- **THEN** 系统计算中位数与四分位数，得到 `专用sigma=(P75-P25)/1.35`，并以 `新上限=中位数+sigma*专用sigma`、`新下限=中位数-sigma*专用sigma` 得到新阈值

### Requirement: 多测试项处理与写回顺序
系统 MUST 按已配置测试项列表顺序逐项处理与写回；若同一坐标被多个测试项命中，以后处理的配置结果为准。

#### Scenario: 多测试项写回
- **WHEN** 多个测试项产生越界坐标
- **THEN** 系统按配置顺序依次写回，最终 INK 结果以最后一次命中为准

### Requirement: 越界点 INK 回写
系统 MUST 以新上下限判定越界数据（小于下限或大于上限），记录其 X/Y 坐标并回写对应 TSK：将 die 属性标记为 Fail，Bin 设为输入的 Ink Bin。

#### Scenario: 标记越界点
- **WHEN** 某数据点测试值超出新上下限
- **THEN** 系统将对应 TSK die 设为 Fail 且 Bin=Ink Bin

### Requirement: 输出路径与命名
系统 MUST 将处理后的 TSK 输出到 `DPAT_INK` 文件夹中，文件名在原名后追加 `_DPAT` 后缀，并保持 TSK 文件结构与 `TSK_spec_2013` 兼容。

#### Scenario: 生成输出文件
- **WHEN** 处理完成
- **THEN** 系统在输出目录生成带 `_DPAT` 后缀的 TSK 文件
