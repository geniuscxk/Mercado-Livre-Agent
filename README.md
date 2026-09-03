# Mercado Livre 月度账单数据处理 Agent

本项目用于处理 Mercado Livre 巴西站后台导出的 `Vendas BR` 销售账单 Excel，并生成月度财务汇总、SKU 出库数量和可复核的计算明细。

本 README 记录当前已经确认的统计口径、Excel 公式、组合订单处理方式、SKU 出库逻辑和交付前校验要求。

> 业务口径以 `AGENTS.md` 与 `AGENTFORCODEX.MD` 为执行规范。本 README 用于人工阅读、复核和快速理解公式。

## Codex 使用方式（必读）

如果使用 OpenAI Codex 处理本项目，克隆仓库后必须在仓库目录中工作，并在执行任何账单处理任务之前完整读取项目规则。

推荐安装方式：

```bash
git clone https://github.com/geniuscxk/Mercado-Livre-Agent.git
cd Mercado-Livre-Agent
```

Codex 的读取顺序必须是：

```text
1. AGENTS.md
2. AGENTFORCODEX.MD
3. README.md（用于人工说明和公式复核）
```

其中：

- `AGENTS.md` 是项目级 Agent 入口规则；
- **Codex 在执行任务前必须继续完整读取并遵守 `AGENTFORCODEX.MD`**；
- `AGENTFORCODEX.MD` 包含面向 Codex 的严格执行顺序、数据校验、公式实现、异常停止条件和交付检查；
- 不得只读取 `AGENTS.md` 后跳过 `AGENTFORCODEX.MD`；
- 不得凭 README 自行简化、改写或推断业务公式；
- 如果 `AGENTS.md`、`AGENTFORCODEX.MD` 与用户当前明确指令之间存在冲突，必须停止执行并让用户确认，不得自行决定。

给 Codex 的推荐指令：

```text
请克隆并进入这个仓库：
https://github.com/geniuscxk/Mercado-Livre-Agent.git

执行任何任务前，先完整读取 AGENTS.md，随后必须完整读取 AGENTFORCODEX.MD，并严格按其中规则处理 Mercado Livre 账单。不要跳过 AGENTFORCODEX.MD，也不要自行修改统计口径。
```

仅执行 `git clone` 的作用是下载仓库；**要让 Codex 正确使用本 Agent，还必须让 Codex 在该仓库目录中读取并遵守 `AGENTS.md` 和 `AGENTFORCODEX.MD`。**

## 1. 输入文件与标准字段

输入文件为 `.xlsx`，目标工作表通常为：

```text
Vendas BR
```

不要永久假定表头固定在第 6 行。应先搜索包含 `N.º de venda` 的行作为表头，再通过字段名定位列。

当前标准导出常用字段如下：

| 标准列 | 字段名 | 用途 |
| --- | --- | --- |
| D | `Estado` | 订单状态 |
| H | `Unidades` | 商品实际数量/件数 |
| I | `Receita por produtos (BRL)` | 订单金额、GMV |
| L | `Tarifa de venda e impostos (BRL)` | 佣金及销售相关税费 |
| M | `Receita por envio (BRL)` | 买家支付并进入卖家账目的配送收入 |
| N | `Tarifas de envio (BRL)` | 卖家承担的运费 |
| R | `Cancelamentos e reembolsos (BRL)` | 退款/取消相关金额 |
| S | `Total (BRL)` | Mercado Livre 该行最终回款 |
| W | `SKU` | 商品 SKU |

## 2. 最重要的统计原则：什么时候 COUNT，什么时候 SUM

### 订单数可以 COUNT

订单数统计的是“订单财务行数量”，不是商品件数。

标准逻辑：

```text
订单数 = Receita por produtos (BRL) 非空的数据行数
```

在标准列位置下，可使用：

```excel
=COUNT(I7:I最后一行)
```

组合订单的商品子行通常没有 GMV，因此不会被重复计数。

### 商品数量必须 SUM

凡是统计：

- SKU 出库数量
- 商品件数
- 实际销售件数
- 某 SKU 的数量

都必须优先使用：

```text
SUM(Unidades)
```

禁止使用：

```text
COUNT(SKU)
COUNTIF(SKU)
分组后统计行数
value_counts()
SKU 出现次数
```

来代替真实件数。

原因：一条订单记录可能 `Unidades = 2`、`3` 或更高。此时该记录必须按实际件数计入。

例如：

```text
SKU = MIC-U122CDX-BLK
Unidades = 2
```

这条记录必须计为 2 件，而不是 1 件。

## 3. 订单数

### 业务公式

```text
订单数 = COUNT(Receita por produtos (BRL))
```

### Excel 示例

假设数据从第 7 行开始：

```excel
=COUNT(I7:I328)
```

正式处理时，`328` 必须替换为实际最后一条有效数据行。

### 组合订单

组合订单通常由：

- 1 行组合订单主行
- N 行商品子行

构成。

主行通常有 GMV，但没有 SKU 和 Unidades；商品子行通常有 SKU 和 Unidades，但 GMV 为空。

因此：

```text
组合订单主行：订单数 +1
组合订单商品子行：订单数 +0
```

一个订单购买 2 个不同 SKU，仍然只能算 1 个订单。

## 4. 订单金额

### 业务公式

```text
订单金额 = SUM(Receita por produtos (BRL))
```

### Excel 示例

```excel
=SUM(I7:I328)
```

保留 BRL 原始金额，最终显示 2 位小数。

## 5. 平台补贴

平台补贴不能简单依赖 `Descontos e bônus` 列，必须根据账单财务恒等关系逐行反算。

### 单行公式

```text
单行平台补贴 =
ROUND(
  Total
  - Receita por produtos
  - Tarifa de venda e impostos
  - Receita por envio
  - Tarifas de envio
  - Cancelamentos e reembolsos,
  2
)
```

### 标准列 Excel 公式

例如第 7 行：

```excel
=ROUND(
IFERROR(--S7,0)
-IFERROR(--I7,0)
-IFERROR(--L7,0)
-IFERROR(--M7,0)
-IFERROR(--N7,0)
-IFERROR(--R7,0),
2)
```

也可以写成单行：

```excel
=ROUND(IFERROR(--S7,0)-IFERROR(--I7,0)-IFERROR(--L7,0)-IFERROR(--M7,0)-IFERROR(--N7,0)-IFERROR(--R7,0),2)
```

### 总平台补贴

必须先对每一行进行 `ROUND(...,2)`，再汇总：

```text
总平台补贴 = SUM(所有单行平台补贴)
```

例如辅助列为 X：

```excel
=SUM(X7:X328)
```

不要先把 I/L/M/N/R/S 六列分别求总和后再统一相减，因为逐行四舍五入口径可能产生分差。

## 6. 佣金

### 业务公式

```text
佣金 = SUM(Tarifa de venda e impostos (BRL))
```

### Excel 示例

```excel
=SUM(L7:L328)
```

佣金通常为负数，必须保留源账单原始正负号，禁止取绝对值。

## 7. 运费

### 业务公式

```text
运费 = SUM(Tarifas de envio (BRL))
```

### Excel 示例

```excel
=SUM(N7:N328)
```

注意：

```text
N = Tarifas de envio
```

才是当前经营表中的“运费”。

不要误用：

```text
M = Receita por envio
```

因为 M 列是买家支付并进入卖家账目的配送收入，不是卖家承担的运费。

## 8. 退款

退款统计时 `Estado` 状态全选，不需要先筛状态。

### 业务公式

```text
退款 = SUM(Cancelamentos e reembolsos (BRL))
```

### Excel 示例

```excel
=SUM(R7:R328)
```

保留 Mercado Livre 原始正负号，不要转为正数。

## 9. 实际营收

退款已经以负数记录，所以实际营收计算为：

```text
实际营收 = 订单金额 + 退款
```

Excel 示例：

```excel
=订单金额单元格+退款单元格
```

不要再次把退款取绝对值后相减，否则会重复改变符号。

## 10. SKU 出库数量

这是最容易因为 `COUNT(SKU)` 而算错的指标。

### 唯一状态排除条件

SKU 出库数量只排除：

```text
Estado = "Cancelada pelo comprador"
```

也就是纳入条件为：

```text
Estado != "Cancelada pelo comprador"
```

除非用户以后明确修改业务口径，否则不要额外排除：

- 退款
- 退货
- 纠纷
- 平台取消
- `Pacote cancelado pelo Mercado Livre`
- 其他 Estado 状态

### 正确业务公式

```text
每个 SKU 出库数量 =
SUM(Unidades)
GROUP BY SKU
WHERE Estado != "Cancelada pelo comprador"
```

### Excel SUMIFS 示例

假设 A2 是需要统计的 SKU：

```excel
=SUMIFS(
$H$7:$H$328,
$W$7:$W$328,A2,
$D$7:$D$328,"<>Cancelada pelo comprador"
)
```

### 禁止的写法

```excel
=COUNTIF(W:W,A2)
```

以及任何等价的 `COUNT(SKU)` 逻辑都禁止用于“出库件数”。

## 11. 组合订单对 SKU 出库的处理

组合订单主行通常：

```text
SKU = 空
Unidades = 空
```

因此不计 SKU 出库。

组合订单商品子行通常包含：

```text
SKU
Unidades
```

因此必须分别按实际 SKU 和实际 `Unidades` 计入出库数量。

例如：

```text
组合订单主行：1 单
  子行 1：SKU-A，Unidades = 1
  子行 2：SKU-B，Unidades = 2
```

最终：

```text
订单数 = 1
SKU-A 出库 = 1
SKU-B 出库 = 2
总出库件数 = 3
```

## 12. 公式汇总速查表

| 指标 | 业务公式 | 标准列 |
| --- | --- | --- |
| 订单数 | `COUNT(Receita por produtos)` | I |
| 订单金额 | `SUM(Receita por produtos)` | I |
| 平台补贴 | `SUM(ROUND(S-I-L-M-N-R,2))` | S/I/L/M/N/R |
| 佣金 | `SUM(Tarifa de venda e impostos)` | L |
| 运费 | `SUM(Tarifas de envio)` | N |
| 退款 | `SUM(Cancelamentos e reembolsos)` | R |
| 实际营收 | `订单金额 + 退款` | 派生指标 |
| SKU 出库数量 | `SUM(Unidades) GROUP BY SKU WHERE Estado != Cancelada pelo comprador` | H/W/D |

## 13. 已验证回归样例：第二份账单

以下数字只作为当前测试文件的回归验证值，严禁硬编码进正式程序：

| 指标 | 已确认结果 |
| --- | ---: |
| Orders | 322 |
| Order amount (GMV) | R$50,394.74 |
| Platform subsidy | R$2,216.84 |
| Refunds | -R$2,110.74 |
| Commission | -R$7,553.38 |
| Shipping fee | -R$4,514.23 |
| Actual revenue | R$48,284.00 |
| `MIC-U122CDX-BLK` outbound units | 70 |

U122 的验证重点：源账单中存在一条 `Unidades = 2` 的记录，所以不能用 SKU 行数代替件数统计。

## 14. 导出异常检查

正式计算之前必须识别组合订单，然后检查是否存在：

```text
Unidades 有值
Receita por produtos 为空
且当前行不是已识别的组合订单商品子行
```

如果同时满足以上条件，视为 Mercado Livre 导出异常。

出现异常时必须停止计算，不得：

- 自动跳过
- 删除异常行
- 把金额当 0
- 继续输出部分结果
- 沿用其他文件的特殊排除行号

## 15. 交付前强制校验

每次处理新账单前后至少检查：

1. 表头和字段映射正确。
2. 组合订单主行和商品子行识别正确。
3. 订单数按 GMV 财务行计数，没有按 SKU 行数计数。
4. 订单金额使用 `SUM(I)`。
5. 平台补贴逐行使用 `ROUND(S-I-L-M-N-R,2)` 后再 SUM。
6. 佣金使用 `SUM(L)`。
7. 运费使用 `SUM(N)`，没有误用 M。
8. 退款使用 `SUM(R)`，Estado 全选。
9. SKU 出库只排除 `Cancelada pelo comprador`。
10. SKU 出库使用 `SUM(Unidades)`，禁止 `COUNT(SKU)`。
11. 对所有 `Unidades > 1` 的行进行专项抽查，确认实际数量没有被算成 1。
12. 组合订单商品子行仍正确进入 SKU 出库统计。
13. `SKU Outbound` 合计与 Summary 中总出库数量一致。
14. 所有金额保留原始正负号。
15. 输出文件不存在 `#REF!`、`#DIV/0!`、`#VALUE!`、`#NAME?`、`#N/A` 等错误。

## 16. 文档关系

- `README.md`：给人工阅读，解释公式和业务逻辑。
- `AGENTS.md`：通用 Agent 业务规则。
- `AGENTFORCODEX.MD`：面向 Codex 的严格执行规范和实现约束。

当三份文档出现冲突时，应先停止自动处理并核对最新业务口径，不得自行猜测。
