# Mercado Livre 月度账单提取规则

## 任务目标

读取 Mercado Livre 巴西站导出的销售账单 Excel，先检查导出数据是否完整，再输出月度财务汇总和按 SKU 统计的出库数量。

不得覆盖原始 Excel。输出文件中不得包含买家姓名、CPF、地址、电话或其他买家个人信息。

## 输入要求

- 输入格式为 `.xlsx`。
- 工作表名称应为 `Vendas BR`。
- 先定位包含 `N.º de venda` 的表头行，不要假定数据永远从固定行开始。
- 当前标准导出中，表头通常位于第 6 行，数据从下一行开始。
- 必须识别下列字段。因为工作表中可能存在同名字段，应同时校验字段名称及其所在的销售区列位置。

| 列 | 字段 | 用途 |
| --- | --- | --- |
| D | `Estado` | 订单状态 |
| H | `Unidades` | 商品数量 |
| I | `Receita por produtos (BRL)` | 订单金额、GMV |
| L | `Tarifa de venda e impostos (BRL)` | 佣金 |
| M | `Receita por envio (BRL)` | 买家支付并计入卖家账目的配送收入 |
| N | `Tarifas de envio (BRL)` | 卖家承担的运费 |
| R | `Cancelamentos e reembolsos (BRL)` | 退款 |
| S | `Total (BRL)` | 回款 |
| W | `SKU` | 商品 SKU |

缺少任何必需字段、表头名称与预期列不符，或者金额字段存在无法解析的非数字内容时，立即停止，不得猜测字段映射。

## 空值与数字规则

- `None`、空字符串以及只包含空格的字符串均视为空值。
- 金额计算时，空值按 `0` 处理。
- SKU、销售编号等标识符必须按文本保存，不得转换为科学计数法或丢失前导字符。
- 金额保留 Mercado Livre 源数据中的正负号。佣金、运费和退款不得擅自转换为正数。
- 每一行的平台补贴先四舍五入到 2 位小数，再进行汇总；最终金额再四舍五入到 2 位小数。

## 第一步：识别组合订单

必须在异常检查和任何汇总之前识别组合订单。

1. 从第一条数据开始按行向下扫描 D 列。
2. 如果 D 列完全匹配正则表达式 `^Pacote de (\d+) produtos$`，该行是组合订单主行。
3. 从状态文本中读取商品行数 `N`，把主行后面连续的 `N` 行标记为该组合订单的商品子行。
4. 组合订单主行通常具有以下特征：I 至 S 列保存整单财务数据，H 列和 W 列为空。
5. 组合订单商品子行通常具有以下特征：H 列和 W 列有值，I 至 S 列为空。
6. 不得依赖灰色背景识别组合订单；颜色只能用于人工查看，不能作为计算依据。
7. 如果主行声明的子行数量超出数据范围，或者预期子行缺少 `Unidades` 或 `SKU`，立即停止并提示组合订单结构错误。

## 第二步：检查 Mercado Livre 导出异常

完成组合订单标记后，逐行检查：

- H 列 `Unidades` 不为空；
- I 列 `Receita por produtos (BRL)` 为空；
- 当前行不是已经确认的组合订单商品子行。

同时满足以上三个条件即为 Mercado Livre 导出异常。

发现一行或多行异常时：

1. 立即停止全部计算。
2. 不输出订单数、金额、SKU 数量或任何部分结果。
3. 返回异常的 Excel 行号、销售编号和 SKU。
4. 明确提示：`Mercado Livre 导出的账单存在数据缺失：商品数量有值，但 Receita por produtos (BRL) 为空。请重新导出并上传完整账单。`
5. 正式任务中不得自动删除、跳过或按 0 处理异常行。

## 已确认公式总表

以下公式为当前已确认的固定统计口径。标准列位置为：`Estado=D`、`Unidades=H`、`Receita por produtos=I`、`Tarifa de venda e impostos=L`、`Receita por envio=M`、`Tarifas de envio=N`、`Cancelamentos e reembolsos=R`、`Total=S`、`SKU=W`。示例以第 7 行为第一条数据；实际运行必须动态识别最后一行。

| 指标 | 业务公式 | Excel 示例 |
| --- | --- | --- |
| 订单数 | `COUNT(Receita por produtos)`，即统计 I 列非空数字财务行 | `=COUNT(I7:I最后一行)` |
| 订单金额 | `SUM(Receita por produtos)` | `=SUM(I7:I最后一行)` |
| 平台补贴 | 每行 `ROUND(S-I-L-M-N-R,2)` 后再 `SUM` | `=ROUND(IFERROR(--S7,0)-IFERROR(--I7,0)-IFERROR(--L7,0)-IFERROR(--M7,0)-IFERROR(--N7,0)-IFERROR(--R7,0),2)` |
| 佣金 | `SUM(Tarifa de venda e impostos)` | `=SUM(L7:L最后一行)` |
| 运费 | `SUM(Tarifas de envio)` | `=SUM(N7:N最后一行)` |
| 退款 | `SUM(Cancelamentos e reembolsos)`，Estado 不筛选 | `=SUM(R7:R最后一行)` |
| 实际营收 | `订单金额 + 退款`；退款保留负号 | `=订单金额单元格+退款单元格` |
| SKU 出库数量 | `Estado <> "Cancelada pelo comprador"` 后按 SKU 分组 `SUM(Unidades)` | `=SUMIFS($H$7:$H$最后一行,$W$7:$W$最后一行,SKU单元格,$D$7:$D$最后一行,"<>Cancelada pelo comprador")` |

### COUNT 与 SUM 的强制原则

- `COUNT` 只用于明确统计订单行数、记录数或出现次数。订单数使用 I 列财务行计数，是因为组合订单只有主财务行的 I 列有金额，商品子行 I 列为空，因此一笔组合订单仍只计 1 单。
- 商品数量、出库件数必须以数量字段为权重进行 `SUM`。不得用 `COUNT(SKU)`、SKU 出现次数、分组行数或 `value_counts()` 替代 `SUM(Unidades)`。
- 如果一条记录 `Unidades = 2`，必须计 2 件；`Unidades = 3` 必须计 3 件。
- 组合订单主行没有 SKU/Unidades，不计 SKU 出库；组合订单商品子行按各自 SKU 和实际 `Unidades` 分别计入。

## 第三步：计算汇总指标

只有异常检查通过后才能计算。

### 1. 订单数

- 统计 I 列 `Receita por produtos (BRL)` 非空的数字财务行数量。
- Excel 示例：`=COUNT(I7:I最后一行)`。
- 普通订单的财务行计为 1 单。
- 组合订单只在主行计为 1 单；商品子行的 I 列为空，因此不重复计数。
- 买家取消、平台取消、退款或退货订单仍包含在订单数中，除非用户另行改变口径。

### 2. 订单金额

```text
订单金额 = SUM(Receita por produtos (BRL))
```

Excel 示例：`=SUM(I7:I最后一行)`。

### 3. 平台补贴

不得直接汇总 Q 列 `Descontos e bônus`。必须对每一行按下式反算：

```excel
=ROUND(
  IFERROR(--Total,0)
  -IFERROR(--Receita_por_produtos,0)
  -IFERROR(--Tarifa_de_venda_e_impostos,0)
  -IFERROR(--Receita_por_envio,0)
  -IFERROR(--Tarifas_de_envio,0)
  -IFERROR(--Cancelamentos_e_reembolsos,0),
  2
)
```

按当前标准列位置，对数据行 7 的 Excel 公式为：

```excel
=ROUND(IFERROR(--S7,0)-IFERROR(--I7,0)-IFERROR(--L7,0)-IFERROR(--M7,0)-IFERROR(--N7,0)-IFERROR(--R7,0),2)
```

最终平台补贴为所有逐行反算结果之和，并四舍五入到 2 位小数。

### 4. 佣金

```text
佣金 = SUM(Tarifa de venda e impostos (BRL))
```

Excel 示例：`=SUM(L7:L最后一行)`。保留原始负号。

### 5. 运费

```text
运费 = SUM(Tarifas de envio (BRL))
```

Excel 示例：`=SUM(N7:N最后一行)`。M 列 `Receita por envio` 是配送收入，不是这里的“运费”。保留原始负号。

### 6. SKU 出库数量

仅此指标按 SKU 拆分，财务指标不按 SKU 分摊。

逐行纳入出库数量必须同时满足：

- W 列 `SKU` 不为空；
- H 列 `Unidades` 是有效数字；
- D 列 `Estado` 不完全等于 `Cancelada pelo comprador`。

按 SKU 分组后汇总 H 列 `Unidades`：

```text
SKU 出库数量 = SUM(Unidades)
WHERE Estado != "Cancelada pelo comprador"
GROUP BY SKU
```

**数量统计必须使用 `SUM(Unidades)`，禁止使用 `COUNT(SKU)`、SKU 出现次数或商品行数代替出库件数。** 一条订单记录可能包含 `Unidades = 2`、`3` 或更高数量；此时必须按实际 `Unidades` 全额计入。只有在明确统计订单行数、记录数或 SKU 出现次数时才允许使用 `COUNT`。

严格执行内部统一口径：只排除 `Cancelada pelo comprador`，不得额外排除 `Pacote cancelado pelo Mercado Livre`、退款、退货、纠纷或其他状态。

组合订单主行没有 SKU 和数量，因此不计入出库数量；其商品子行必须分别计入实际 SKU，并按各子行 `Unidades` 求和。平台套装的组成商品也按各自 SKU 和数量计入。

### 7. 退款

```text
退款 = SUM(Cancelamentos e reembolsos (BRL))
```

Excel 示例：`=SUM(R7:R最后一行)`。退款统计时 `Estado` 全选，不按状态筛选；保留原始正负号。

### 8. 实际营收（经营表派生指标）

```text
实际营收 = 订单金额 + 退款
```

退款本身保留负号，因此不得再额外做减法。例如退款为 `-100` 时，实际营收应为 `订单金额 + (-100)`。

## 输出格式

正常通过校验后，输出一个新的 Excel，至少包含以下工作表：

### `Summary`

按固定顺序输出：

1. Orders
2. Order amount (GMV)
3. Platform subsidy
4. Commission
5. Shipping fee
6. SKU outbound units
7. Refunds

每个指标必须带计算口径说明。金额统一使用 BRL、2 位小数；数量使用整数。平台补贴必须明确标注为反算结果，并注明未使用 `Descontos e bônus`。

### `SKU Outbound`

包含：

| SKU | Outbound units |
| --- | ---: |

按出库数量从高到低排序，数量相同时按 SKU 升序排列，并在底部显示总计。

### `Calculation Detail`

仅保留复核所需字段，不包含买家个人信息。至少包括源表行号、销售编号、日期、Estado、组合订单行类型、SKU、Unidades、I/L/M/N/R/S 列、逐行反算平台补贴、订单计数标记、出库计数标记及排除原因。

## 交付前检查

- 确认正式运行没有被跳过的异常行。
- 确认组合订单主行计 1 单，商品子行计 0 单。
- 确认组合订单商品子行仍计入 SKU 出库数量。
- 确认 SKU 出库数量按 `SUM(Unidades)` 聚合，而不是按 `COUNT(SKU)` 或商品行数统计。
- 对所有 `Unidades > 1` 的商品行进行抽查，确认实际件数被完整累计。
- 确认只有 `Cancelada pelo comprador` 被排除在出库数量之外。
- 确认平台补贴未引用 Q 列 `Descontos e bônus`。
- 确认佣金、运费和退款保留原始符号。
- 确认 SKU 出库明细合计等于 Summary 中的 SKU outbound units。
- 扫描 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?` 和 `#N/A`，出现任一公式错误不得交付。

## 当前示例的回归测试

本节只用于验证当前示例文件，不得把这些数字硬编码到正式处理程序中。

当前示例已由用户明确授权跳过源表第 780 行。采用该一次性测试例外时，结果应为：

| 指标 | 预期结果 |
| --- | ---: |
| Orders | 749 |
| Order amount (GMV) | R$51,701.93 |
| Platform subsidy | R$1,352.71 |
| Commission | -R$6,221.73 |
| Shipping fee | -R$10,298.01 |
| SKU outbound units | 797 |
| Refunds | -R$1,371.70 |

SKU 出库表应包含 86 个 SKU，数量合计为 797。

正式处理任何新文件时不得沿用“排除第 780 行”的例外；必须重新执行完整异常检查。
