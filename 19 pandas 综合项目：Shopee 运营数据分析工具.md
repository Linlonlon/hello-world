# 14｜pandas 综合项目：Shopee 运营数据分析工具

## 一、项目目标

本项目用于练习并串联 pandas 数据分析的完整流程。

目标不是单独练习某一个 API，而是完成一套相对完整的数据处理 pipeline：

> 数据读取  
> → 数据结构检查  
> → 数据清洗与标准化  
> → 数据质量检查  
> → 多表关联  
> → 成本利润计算  
> → 分组聚合分析  
> → Excel 多 Sheet 报告输出

当前完成版本：

**v0.1**

v0.1 的核心目标：

> 能够从订单表、产品表、成本表出发，自动完成数据检查、数据关联、成本利润计算、SKU 分析、销售趋势分析、异常汇总，并最终输出 Excel 报告。

---



## 二、项目输入数据

项目读取三个 Excel 数据源：

```python
orders_df = pd.read_excel("input/orders.xlsx")
products_df = pd.read_excel("input/products.xlsx")
cost_df = pd.read_excel("input/cost.xlsx")
```

三张表承担不同职责：

### 1. 订单表 orders

主要字段：

```apl
订单号
下单时间
SKU
数量
商品销售额
订单状态
```

属于业务流水数据。

订单表中的 SKU **允许重复**，因为：

- 不同订单可以购买相同 SKU
- 同一 SKU 会产生多条销售记录

------

### 2. 产品表 products

主要字段：

```apl
SKU
商品名称
类目
站点
```

属于产品主数据。

要求：

```ABAP
一个 SKU → 一套产品资料
```

因此 SKU 必须：

- 非空
- 唯一

------

### 3. 成本表 cost

主要字段：

```apl
SKU
采购成本
包装成本
头程成本
```

属于成本主数据。

要求：

```ABAP
一个 SKU → 一套成本资料
```

因此 SKU 同样必须：

- 非空
- 唯一

----



## 三、整体程序结构

**最终程序结构：**

```apl
1. 输入数据

2. 必要字段结构检查

3. 数据清洗与标准化

4. 阻断型异常检查

5. 记录型异常检查

6. 多表 merge

7. 成本利润计算

8. SKU / 日期 / 异常聚合分析

9. Excel 多 Sheet 输出
```

其中需要**特别注意：**

> 必要字段检查必须发生在数据清洗之前。

**原因是清洗代码本身就会访问这些字段。**

如果字段不存在：

```ABAP
df["采购成本"]
```

在清洗阶段就会先触发 `KeyError`，后面的字段检查根本没有机会执行。

因此正确顺序应该是：

```ABAP
读取
↓
结构检查
↓
清洗
```

而不是：

```apl
读取
↓
清洗
↓
结构检查
```

---



## 四、必要字段检查

### 1. 字段配置

```python
required_order_columns = [
    "订单号",
    "下单时间",
    "SKU",
    "数量",
    "商品销售额",
    "订单状态"
]

required_product_columns = [
    "SKU",
    "商品名称",
    "类目",
    "站点"
]

required_cost_columns = [
    "SKU",
    "采购成本",
    "包装成本",
    "头程成本"
]
```

### 2. 检查函数

```python
def check_required_columns(df, required_columns, table_name):
    missing_columns = []

    for col in required_columns:
        if col not in df.columns:
            missing_columns.append(col)

    if not missing_columns:
        print("{0} 字段检查通过".format(table_name))
    else:
        raise ValueError(
            "{0} 字段检查不通过，缺少字段:{1}".format(
                table_name,
                missing_columns
            )
        )
```

**设计思想：**

```apl
字段缺失
≠ 某一条记录出错

字段缺失
= 整张表的结构不符合程序要求
```

因此属于：

**阻断型异常**

必须停止程序。

---



## 五、数据清洗与标准化

### 1. 文本字段清洗

```python
def clean_text_columns(df, columns):
    for col in columns:
        df[col] = (
            df[col]
            .astype("string")
            .str.strip()
            .replace("", pd.NA)
        )
```

### astype("string")

使用 pandas 的字符串类型，而不是 Python：

```python
astype(str)
```

**原因：**

`astype(str)` 可能将缺失值转换成普通字符串，例如：

```apl
NaN → "nan"
```

这样后续：

```python
.isna()
```

可能无法继续识别其为真正的缺失值。

而：

```python
astype("string")
```

可以更好地保留 pandas 缺失值语义。

------

### 2. str.strip()

```python
.str.strip()
```

用于去掉字符串：

**左右两端空格**

例如：

```apl
" A001 " → "A001"
```

SKU 中的首尾空格尤其危险。

因为：

```apl
"A001"
```

与：

```apl
"A001 "
```

在 merge 时并不是同一个值。

这可能导致本来应该匹配成功的数据变成：

```apl
left_only
```

------

### 3. replace("", pd.NA)

```python
.replace("", pd.NA)
```

用于把：

**空字符串**

统一转换为缺失值。

例如：

```ABAP
"   "
↓ strip
""
↓ replace
<NA>
```

需要注意：

```python
.replace("", pd.NA)
```

不会修改：

```
"A  B"
```

因为 `"A  B"` 并不等于空字符串 `""`。

因此最终效果：

```apl
" A001 " → "A001"

"   " → ""

"" → <NA>

"A  B" → "A  B"
```

这一步很重要，因为：

```python
.isna()
```

本身通常不会把：

```
""
```

判断为缺失值。

**因此最好在清洗阶段统一把各种形式的“空内容”标准化成真正缺失值。**

---



## 六、数值字段清洗

```python
def clean_numeric_columns(df, columns):
    for col in columns:
        df[col] = pd.to_numeric(
            df[col],
            errors="coerce"
        )
```

使用：

```python
pd.to_numeric()
```

而不是：

```python
.astype(float)
```

主要原因不是数据是否需要小数，而是：

**容错转换**

例如：

```apl
数量

2
3
abc
5
```

如果直接：

```python
.astype(float)
```

遇到：

```
abc
```

**可能直接报错，中断程序。**

而：

```python
pd.to_numeric(
    ...,
    errors="coerce"
)
```

会转换成：

```apl
2
3
NaN
5
```

即：

> 无法转换的异常数据 → NaN

这样程序可以继续运行，再由后面的异常检查负责识别问题。

---



## 七、日期数据清洗

```python
orders_df["下单时间"] = pd.to_datetime(
    orders_df["下单时间"],
    errors="coerce"
)
```

非法日期会转换为：

```apl
NaT
```

例如：

```apl
正常时间 → datetime
非法时间 → NaT
```

------



## 八、数据清洗与数据检查的职责区别

**这是本项目的重要设计思想。**

### 数据清洗

负责：

> 把数据转换成统一、可处理的形式。

例如：

```apl
" A001 " → "A001"

"   " → pd.NA

"20" → 20

"abc" → NaN

非法日期 → NaT
```

------

### 数据质量检查

负责：

> 判断清洗后的数据是否符合程序规则和业务规则。

例如：

```apl
数量 <= 0
销售额 < 0
SKU 缺失
状态非法
成本字段缺失
```

因此：

```ABAP
清洗
≠ 判断业务数据是否正确

清洗
= 标准化

检查
= 判断是否合法
```

---



## 九、缺失值体系

本项目涉及：

```ABAP
NaN
pd.NA
NaT
```

可以暂时这样理解：

### NaN

常见于：

**数值型缺失**

例如：

```python
pd.to_numeric(..., errors="coerce")
```

可能产生：

```apl
NaN
```

------

### pd.NA

pandas 的：

**通用型缺失值**

例如：

```python
df["利润率"] = pd.NA
```

表示：

> 当前没有合理的值。

------

### NaT

用于：

**日期时间型缺失**

例如：

```python
pd.to_datetime(..., errors="coerce")
```

产生：

```apl
NaT
```

------

### isna()

```python
.isna()
```

可以统一理解为：

> 判断是否为缺失值

而不是：

> 只判断 NaN。

它可以用于识别：

```
NaN
pd.NA
NaT
```

---



## 十、阻断型异常

阻断型异常指：

> 如果出现，后面的分析结果已经不可靠，应该立即停止程序。

当前包括：

```ABAP
必要字段不存在

产品表 SKU 缺失

成本表 SKU 缺失

产品表 SKU 重复

成本表 SKU 重复
```

------

## 十一、主键缺失检查

```python
def check_key_missing(df, key_column, table_name):
    if df[key_column].isna().any():
        raise ValueError(
            "{0} 中字段 {1} 存在缺失值".format(
                table_name,
                key_column
            )
        )
    else:
        print(
            "{0} 中字段 {1} 无缺失".format(
                table_name,
                key_column
            )
        )
```

产品表、成本表的 SKU 是：

**连接键 / 主数据键**

如果 SKU 本身缺失：

```apl
SKU    商品名称

NaN    产品A
```

虽然产品记录本身存在，但由于缺少连接键：

> 订单永远无法通过 SKU 找到这条产品资料。

因此主数据 SKU 缺失应作为阻断型异常。

---



## 十二、唯一键检查

```python
def check_unique_key(df, key_column, table_name):
    if df[key_column].is_unique:
        print(
            "{0} 中字段 {1} 没有重复".format(
                table_name,
                key_column
            )
        )
    else:
        duplicate_key_df = df[
            df[key_column].duplicated(
                keep=False
            )
        ]

        print(duplicate_key_df)

        raise ValueError(
            "{0} 中字段 {1} 存在重复，请检查表格".format(
                table_name,
                key_column
            )
        )
```

------

### is_unique

```python
df["SKU"].is_unique
```

用于回答：

> SKU 这一整列有没有重复？

返回：

```apl
True / False
```

属于：

**整体判断**

------

### duplicated()

```python
df["SKU"].duplicated(
    keep=False
)
```

用于回答：

> 到底是哪几条记录重复？

返回布尔 Series。

例如：

```apl
A001    False
A002     True
A002     True
A003    False
```

`keep=False` 表示：

> 同一组重复记录全部标记为 True。

------

### 为什么产品表和成本表不能有重复 SKU

因为：

```apl
订单表
SKU = A001

产品表
A001 → 产品A
A001 → 产品B
```

进行 merge 后，原来一条订单可能变成两条。

导致：

```ABAP
订单行数膨胀
销量重复
销售额重复
成本重复
利润重复
```

因此：

```
产品表 SKU
成本表 SKU
```

必须保持唯一。

---



## 十三、记录型异常

记录型异常指：

> 某一条业务记录存在问题，但不需要让整个程序停止。

例如：

```apl
某一个订单数量异常
```

并不影响其他几千条正常订单继续分析。

因此：

> 保存异常 → 继续运行

------



## 十四、异常容器 all_abnormal

```python
all_abnormal = []
```

每发现一种异常：

```python
all_abnormal.append(
    abnormal_df
)
```

最后统一：

```python
all_abnormal_df = pd.concat(
    all_abnormal,
    ignore_index=True
)
```

设计思想：

```apl
先收集
↓
最后一次性 concat
```

而不是每发现一个异常就立即重复 concat。

这样：

- 结构更清楚
- 更容易维护
- 避免反复创建 DataFrame

------



## 十五、copy()

典型写法：

```python
abnormal_quantity_df = orders_df[
    orders_df["数量"] <= 0
].copy()
```

`.copy()` 的作用：

> 创建筛选结果的独立副本。

因为后面还需要：

```ABAP
abnormal_quantity_df["异常类型"] = "数量异常"
```

如果不 `.copy()`，pandas 可能无法确定：

> 是修改筛选结果，还是间接修改原 DataFrame。

可能产生链式赋值相关问题。

因此：

```apl
筛选后还需要修改
```

通常建议：

```ABAP
.copy()
```

---



## 十六、empty

```python
if not abnormal_quantity_df.empty:
```

`.empty` 用于判断：

> DataFrame 是否没有任何数据。

只有真正发现异常时，才把该 DataFrame 放入：

```
all_abnormal
```

------



## 十七、isin() 与 ~

订单合法状态：

```python
valid_status = [
    "已完成",
    "已取消",
    "处理中"
]
```

检查：

```python
orders_df["订单状态"].isin(
    valid_status
)
```

表示：

> 当前状态是否存在于合法状态列表中。

返回布尔 Series。

前面的：

```ABAP
~
```

表示：

**布尔取反**

因此：

```python
~orders_df["订单状态"].isin(
    valid_status
)
```

表示：

> 筛选不在合法状态列表中的订单。

------



## 十八、isna().any(axis=1)

例如：

```python
orders_df[
    [
        "订单号",
        "下单时间",
        "SKU",
        "数量",
        "商品销售额",
        "订单状态"
    ]
].isna().any(axis=1)
```

拆开理解：

### isna()

检查每个元素：

```apl
是否缺失
```

得到布尔矩阵。

------

### any()

```apl
.any()
```

表示：

> 只要存在任意一个 True，就认为满足条件。

------

### axis=1

表示：

**按行判断**

因此：

```apl
.isna().any(axis=1)
```

完整意思：

> 每一行中，只要任意关键字段为空，就把这一行判断为异常。

------



## 十九、merge

本项目的核心数据关系：

```ABAP
订单表
↓
merge 产品表
↓
order_product_df
↓
merge 成本表
↓
order_detail_df
```

------



## 二十、left merge

```python
pd.merge(
    orders_df,
    products_df,
    on="SKU",
    how="left",
    indicator=True
)
```

使用：

```ABAP
how="left"
```

原因：

> 订单表是分析基表，必须尽可能保留全部订单。

如果使用：

```apl
how="inner"
```

那么匹配不到产品资料的订单会直接消失。

这样反而无法发现：

```apl
产品资料缺失
```

---



## 二十一、indicator=True

```apl
indicator=True
```

会增加：

```apl
_merge
```

列。

常见值：

```ABAP
both
left_only
right_only
```

含义：

```apl
both
→ 左右两张表都匹配到

left_only
→ 左表存在，右表没有匹配到

right_only
→ 右表存在，左表没有匹配到
```

在当前 left merge 中，最重要的是：

```
both
left_only
```

------



## 二十二、missing 和 field 的区别

本项目明确区分：

```apl
missing
→ 整套资料不存在 / 匹配不到

field
→ 资料存在，但内部字段缺失
```

------

### 产品资料缺失

```python
order_product_df[
    order_product_df["_merge"]
    == "left_only"
]
```

代表：

> 订单 SKU 在产品表完全找不到对应记录。

------

### 产品字段缺失

```python
(
    order_product_df["_merge"]
    == "both"
)
&
(
    order_product_df[
        ["商品名称", "类目", "站点"]
    ]
    .isna()
    .any(axis=1)
)
```

表示：

> SKU 已经成功匹配产品资料，但是产品资料里的部分字段为空。

------

### 为什么必须加 both

如果产品资料完全不存在，那么 left merge 后：

```ABAP
商品名称 → NaN
类目 → NaN
站点 → NaN
```

如果只判断：

```apl
.isna().any(axis=1)
```

就无法区分：

```ABAP
产品完全不存在

vs

产品存在，但字段没填
```

因此需要结合：

```
_merge
```

判断。

------



## 二十三、同一订单可能命中多个异常

例如订单本身：

```
SKU 缺失
```

首先会触发：

```
订单关键字段缺失
```

随后 merge 无法找到产品：

```
产品资料缺失
```

成本也无法匹配：

```
成本资料缺失
```

因此当前 v0.1：

> 一条记录允许命中多个异常规则。

这是当前设计结果，不属于程序错误。

但会作为 v0.2 的异常体系优化方向。

------



## 二十四、无异常时的 DataFrame 结构

如果：

```python
all_abnormal = []
```

不能简单：

```apl
all_abnormal_df = pd.DataFrame()
```

因为这是：

```
0 行 × 0 列
```

后面：

```
.groupby("异常类型")
```

会因为没有：

```
异常类型
```

列而触发：

```
KeyError
```

因此：

```python
all_abnormal_df = pd.DataFrame(
    columns=["异常类型"]
)
```

虽然仍然是：

```
0 行
```

但已经有：

```
异常类型
```

这个列结构。

因此：

```ABAP
空 DataFrame
≠
没有列结构的 DataFrame
```

------



## 二十五、成本利润计算

### 单件总成本

```python
order_detail_df["单件总成本"] = (
    order_detail_df["采购成本"]
    + order_detail_df["包装成本"]
    + order_detail_df["头程成本"]
)
```

业务含义：

```apl
单件总成本
=
采购成本
+
包装成本
+
头程成本
```

------

### 订单总成本

```python
order_detail_df["订单总成本"] = (
    order_detail_df["单件总成本"]
    * order_detail_df["数量"]
)
```

表示：

```apl
订单总成本
=
一件商品成本
×
购买数量
```

------

### 订单利润

```python
order_detail_df["订单利润"] = (
    order_detail_df["商品销售额"]
    - order_detail_df["订单总成本"]
)
```

---



## 二十六、NaN 在普通算术中的传播

例如：

```apl
采购成本 = NaN
包装成本 = 2
头程成本 = 3
```

那么：

```apl
NaN + 2 + 3
→ NaN
```

继续：

```apl
单件总成本 → NaN

订单总成本
NaN × 数量
→ NaN

订单利润
销售额 - NaN
→ NaN
```

因此：

> 普通逐元素运算中，缺失值通常会继续向后传播。

------



## 二十七、mask

利润率计算：

```python
mask = (
    order_detail_df["商品销售额"].notna()
    & (order_detail_df["商品销售额"] != 0)
)
```

`mask` 本质：

> 一个布尔 Series。

例如：

```apl
销售额

100
0
NaN
50
```

得到：

```
True
False
False
True
```

它负责标记：

> 哪些行可以进行利润率计算。

------



## 二十八、为什么 mask 不能只写 != 0

原来：

```py
mask = (
    order_detail_df["商品销售额"]
    != 0
)
```

不够严谨。

因为业务真正要求的是：

```apl
有有效销售额
并且
销售额不为 0
```

因此应该：

```apl
.notna()
&
!= 0
```

也就是：

```apl
第一关：
不是缺失值

第二关：
不是 0
```

两者同时满足才允许计算。

------



## 二十九、loc

典型结构：

```ABAP
df.loc[
    行条件,
    列
]
```

例如：

```python
order_detail_df.loc[
    mask,
    "利润率"
]
```

表示：

> 满足 mask 条件的行 × 利润率这一列。

------

### 条件赋值经典模式

```python
mask = 条件

df["新列"] = 默认值

df.loc[
    mask,
    "新列"
] = 满足条件时的结果
```

这个模式以后非常常用。

------



## 三十、利润率

先：

```python
order_detail_df["利润率"] = pd.NA
```

给所有行设置：

```
无法合理计算 / 当前无结果
```

然后：

```python
order_detail_df.loc[
    mask,
    "利润率"
] = (
    order_detail_df.loc[
        mask,
        "订单利润"
    ]
    /
    order_detail_df.loc[
        mask,
        "商品销售额"
    ]
)
```

只给满足条件的订单计算利润率。

------

### 为什么不能直接除

如果：

```
订单利润 = -10
销售额 = 0
```

直接：

```
-10 / 0
```

在 pandas / NumPy 浮点运算里可能得到：

```
-inf
```

正数除 0 可能得到：

```
inf
```

这在运营报告里没有合理业务意义。

因此：

> 销售额为 0 → 利润率保留 pd.NA。

---



## 三十一、groupby

SKU 分析：

```ABAP
.groupby("SKU")
```

表示：

> 按 SKU 分组。

例如：

```apl
A001
A001
A002
A003
A003
```

形成：

```apl
A001 一组
A002 一组
A003 一组
```

然后对每一组分别进行聚合计算。

------



## 三十二、agg 命名聚合

例如：

```python
.agg(
    销量=("数量", safe_sum),
    订单数=("订单号", "nunique")
)
```

结构：

```ABAP
销量
↓
最终输出的新列名
数量
↓
原始 DataFrame 中的字段
safe_sum
↓
使用的聚合方法
```

完整含义：

> 对每个分组里的数量进行 safe_sum，结果列命名为“销量”。

------



## 三十三、size / count / nunique

假设：

```
订单号

O001
O001
NaN
O002
```

------

### size()

```ABAP
.size()
```

统计：

> 总行数

结果：

```
4
```

**缺失值也算一行。**

------

### count()

```ABAP
["订单号"].count()
```

统计：

> 非缺失值数量

结果：

```
3
```

------

### nunique()

```ABAP
["订单号"].nunique()
```

统计：

> 非缺失的不同值数量

这里不同订单号：

```
O001
O002
```

因此：

```
2
```

总结：

|    方法     |     含义     |
| :---------: | :----------: |
|  `size()`   |    总行数    |
|  `count()`  | 非缺失值数量 |
| `nunique()` |  不同值数量  |

------



## 三十四、为什么订单数使用 nunique

```ABAP
订单数=("订单号", "nunique")
```

因为同一个订单号可能出现多行记录。

如果：

```
O001
O001
O002
```

`count()`：

```
3
```

但真正订单数：

```
2
```

因此：

```
nunique()
```

更符合业务指标：

**订单数**

---



## 三十五、sum() 对 NaN 的默认处理

这是本项目后半段最重要的坑之一。

普通算术：

```ABAP
100 + NaN
→ NaN
```

但 pandas：

```python
pd.Series(
    [100, NaN, 200]
).sum()
```

通常得到：

```ABAP
300
```

因为：

```
sum()
```

默认：

```ABAP
跳过 NaN
```

相当于：

```
skipna=True
```

------



## 三十六、为什么 sum 跳过 NaN 可能危险

假设某 SKU：

```apl
订单利润

20
NaN
30
```

如果直接：

```ABAP
sum()
```

得到：

```
50
```

但这个 50 只是：

> 已知利润合计

并不能代表：

> 实际总利润。

缺失那条订单可能本来：

```
+100
```

也可能：

```
-80
```

因此真实利润可能更高，也可能更低。

------



## 三十七、safe_sum

为了避免这种误导：

```python
def safe_sum(series):
    if series.isna().any():
        return pd.NA

    return series.sum()
```

逻辑：

```ABAP
这一组只要存在一个缺失值
↓
整个汇总结果返回 pd.NA
```

否则：

```
正常 sum
```

------

### agg 中使用自定义函数

```
总利润=(
    "订单利润",
    safe_sum
)
```

pandas 会把：

> 当前分组对应的 Series

传给：

```
safe_sum(series)
```

例如：

```
A001 的利润

20
30
NaN
```

相当于 pandas 调用：

```
safe_sum(
    pd.Series(
        [20, 30, NaN]
    )
)
```

最终函数应该返回：

> 一个聚合后的单值。

------

### 注意

传入函数时：

```
safe_sum
```

不是：

```
safe_sum()
```

因为：

```ABAP
safe_sum
→ 把函数本身交给 pandas

safe_sum()
→ 立刻执行函数
```

------



## 三十八、SKU 分析

最终：

```python
sku_report = (
    order_detail_df
    .groupby("SKU")
    .agg(
        销量=("数量", safe_sum),
        销售额=("商品销售额", safe_sum),
        订单数=("订单号", "nunique"),
        总成本=("订单总成本", safe_sum),
        总利润=("订单利润", safe_sum)
    )
    .reset_index()
)
```

分析维度：

> 商品 / SKU

回答：

```ABAP
每个 SKU 卖了多少？

产生多少销售额？

有多少订单？

成本多少？

利润多少？
```

------



## 三十九、dt.date

原始：

```apl
2026-08-26 09:15:32
2026-08-26 14:30:10
2026-08-27 08:05:00
```

执行：

```python
order_detail_df[
    "下单时间"
].dt.date
```

得到：

```apl
2026-08-26
2026-08-26
2026-08-27
```

目的是：

> 把时间粒度降低到“日”。

如果直接：

```
.groupby("下单时间")
```

两个不同时间：

```
09:15
14:30
```

会被分成不同组。

因此先：

```
.dt.date
```

才能统计：

**每日销售趋势**

------



## 四十、to_period("M")

扩展知识：

```ABAP
df["下单时间"].dt.to_period("M")
```

把：

```
2026-08-26 13:25:10
```

转换成：

```
2026-08
```

表示：

> 这条记录属于 2026 年 8 月这个时间周期。

------

### 与 dt.month 的区别

```
.dt.month
```

只会得到：

```
8
```

如果数据跨年：

```
2025-08
2026-08
```

都会变成：

```
8
```

从而被错误聚合。

而：

```
.dt.to_period("M")
```

能够保留：

```
2025-08
2026-08
```

因此更适合月度趋势分析。

---



## 四十一、每日销售趋势

```python
daily_sales = (
    order_detail_df
    .groupby("日期")
    .agg(
        销量=("数量", safe_sum),
        销售额=("商品销售额", safe_sum),
        订单数=("订单号", "nunique"),
        总利润=("订单利润", safe_sum)
    )
    .reset_index()
)
```

和 SKU 分析最大的区别不是聚合技术，而是：

**分析维度不同。**

```ABAP
sku_report
→ SKU 维度
→ 每个商品表现

daily_sales
→ 日期维度
→ 每天整体表现
```

------



## 四十二、异常汇总

```python
abnormal_summary = (
    all_abnormal_df
    .groupby("异常类型")
    .size()
    .reset_index(
        name="异常数量"
    )
)
```

流程：

```ABAP
按异常类型分组
↓
size() 统计每组多少行
↓
reset_index()
↓
恢复普通 DataFrame 结构
```

------



## 四十三、reset_index()

例如：

```ABAP
.groupby("异常类型").size()
```

结果更接近：

```apl
异常类型

成本字段缺失    5
产品资料缺失    3
数量异常        2
dtype: int64
```

其中：

```ABAP
异常类型
```

是索引。

通过：

```python
.reset_index(
    name="异常数量"
)
```

转换成：

```apl
异常类型        异常数量

成本字段缺失       5
产品资料缺失       3
数量异常           2
```

即普通 DataFrame。

------



## 四十四、ExcelWriter

```python
with pd.ExcelWriter(
    "output/运营分析报告.xlsx"
) as writer:
```

用于：

> 管理同一个 Excel 工作簿的多次写入。

然后：

```python
df.to_excel(
    writer,
    sheet_name="..."
)
```

可以分别写入不同 Sheet。

---



## 四十五、为什么不直接多次 to_excel

如果每次：

```python
df.to_excel(
    "运营分析报告.xlsx"
)
```

**属于独立写文件操作。**

**后续写入可能覆盖前面的 Excel 内容。**

`ExcelWriter` 则相当于：

```apl
创建一个工作簿
↓
写 Sheet 1
↓
写 Sheet 2
↓
写 Sheet 3
↓
统一保存并关闭
```

------



## 四十六、with 的意义

```ABAP
with pd.ExcelWriter(...) as writer:
```

和：

```ABAP
with open(...) as f:
```

思想类似。

优点：

> 离开 with 代码块以后，自动进行保存、关闭等资源收尾工作。

无需手动处理 writer 的关闭。

------



## 四十七、最终输出 Sheet

v0.1 最终生成：

```ABAP
运营分析报告.xlsx
```

包含：

```ABAP
订单明细

SKU分析

销售趋势

异常订单明细

异常订单分析
```

---



## 四十八、v0.1 最终完成能力

当前 v0.1 已经完成：

### 数据读取

```
orders
products
cost
```

------

### 数据结构校验

```apl
必要字段检查

主数据 SKU 非空检查

主数据 SKU 唯一性检查
```

------

### 数据标准化

```apl
文本清洗

空字符串标准化

数值转换

日期转换
```

------

### 异常识别

```apl
数量异常

销售额异常

状态异常

订单关键字段缺失

产品资料缺失

产品字段缺失

成本资料缺失

成本字段缺失
```

------

### 数据关联

```
订单 × 产品

订单产品 × 成本
```

------

### 成本利润

```
单件总成本

订单总成本

订单利润

利润率
```

------

### 聚合分析

```
SKU分析

每日销售趋势

异常数量分析
```

------

### 报告输出

```ABAP
Excel 多 Sheet
```

------



## 四十九、v0.1 当前数据处理 Pipeline

最终可以总结为：

```apl
Excel 原始数据
        ↓
必要字段结构检查
        ↓
数据清洗与标准化
        ↓
主数据质量检查
        ↓
订单记录异常检查
        ↓
订单 × 产品 merge
        ↓
产品关联异常检查
        ↓
订单产品 × 成本 merge
        ↓
成本关联异常检查
        ↓
成本利润计算
        ↓
SKU 聚合
        ↓
日期聚合
        ↓
异常聚合
        ↓
Excel 多 Sheet 报告
```

---



## 五十、v0.1 已知限制

当前 v0.1 已经可以完整运行，但仍存在一些需要后续优化的问题。

这些问题暂时不继续修改，而是保留给 v0.2。

------

### 1. 订单状态与分析口径没有完全分离

现在存在：

```
valid_status = [
    "已完成",
    "已取消",
    "处理中"
]
```

但是：

```
已取消订单
处理中订单
```

仍然可能进入：

```
销量
销售额
成本
利润
销售趋势
```

分析。

因此未来必须明确：

> 哪些订单状态应该真正参与经营指标统计？

这属于：

**业务分析口径**

而不是单纯 pandas 技术问题。

------

### 2. 根因异常与派生异常重复

例如：

```
订单 SKU 缺失
```

可能同时被记录成：

```ABAP
订单关键字段缺失
产品资料缺失
成本资料缺失
```

实际根因可能只有：

```
订单 SKU 缺失
```

后面的两个只是派生结果。

因此异常明细可能出现重复报警。

------

### 3. 无异常时异常明细结构不完整

当前：

```apl
pd.DataFrame(
    columns=["异常类型"]
)
```

虽然能够保证：

```
groupby("异常类型")
```

正常运行，但最终导出的：

```
异常订单明细
```

只有：

```
异常类型
```

一列。

未来可以让空异常表也保留完整字段结构。

------

### 4. output 文件夹依赖人工提前建立

当前：

```ABAP
"output/运营分析报告.xlsx"
```

要求：

```
output
```

目录已经存在。

如果不存在，程序可能报错。

------

### 5. 部分业务规则硬编码

例如：

```
valid_status
required_order_columns
required_product_columns
required_cost_columns
```

都直接写在主程序中。

未来可以集中配置。

------

### 6. 主程序仍然较长

虽然已经有：

```
clean_text_columns()
clean_numeric_columns()
check_required_columns()
check_unique_key()
check_key_missing()
safe_sum()
```

但：

```
异常检查
merge
计算
报表
导出
```

仍然集中在主程序中。

以后可以进一步函数化。

---



## 五十一、v0.2 优化规划

v0.2 不计划现在立即开发。

当前优先继续学习：

```ABAP
openpyxl
↓
实际运费申诉自动化工具
↓
积累真实项目经验
↓
再回头升级本项目 v0.2
```

v0.2 的目标不再只是：

> 能跑通。

而应该升级成：

> 业务口径更明确、结果更可信、异常更容易处理、程序更加稳定。

------



## 五十二、v0.2 第一优先级：明确业务分析口径

需要确定：

```
已完成
已取消
处理中
退款
退货
其他状态
```

分别如何参与分析。

未来可能形成：

```
全部订单
→ 用于异常检查和原始明细

有效订单
→ 用于销量 / 销售额 / 利润分析
```

例如：

```
analysis_df = order_detail_df[
    order_detail_df[
        "订单状态"
    ].isin(
        analysis_status
    )
].copy()
```

其中：

```
analysis_status
```

由业务规则明确确定。

------

### 需要明确的问题

例如：

```
已取消订单是否算销量？

处理中订单是否算销售额？

退款订单如何计算利润？

退货订单成本如何处理？

零销售额订单属于什么情况？
```

这些不能仅靠代码猜测。

必须先定义：

**运营统计规则**

------

## 五十三、v0.2 第二优先级：异常体系升级

当前：

```
一条记录可以对应多个异常类型
```

未来可以设计：

```
根因异常
派生异常
```

例如：

```
订单 SKU 缺失

根因：
订单关键字段缺失

派生：
产品资料无法匹配
成本资料无法匹配
```

未来可以考虑：

```
异常优先级
```

例如：

```
P0：结构错误

P1：主数据错误

P2：订单关键字段错误

P3：业务规则异常
```

这样异常报告会更容易处理。

------



## 五十四、v0.2 第三优先级：完善异常报告

可以加入：

```
异常订单数

异常率

异常类型占比

成本缺失数量

产品资料缺失数量

重复 SKU 数量
```

甚至：

```
异常处理建议
```

例如：

```
成本资料缺失
→ 补充 cost.xlsx

产品资料缺失
→ 检查 products.xlsx

状态异常
→ 检查平台状态映射
```

------



## 五十五、v0.2 第四优先级：完善数据完整性标识

现在通过：

```
safe_sum()
```

解决了：

> 缺失数据被 pandas sum 自动跳过

的问题。

未来可以进一步让报表明确展示：

```
销量数据是否完整

销售额数据是否完整

成本数据是否完整

利润数据是否完整
```

例如新增：

```
成本缺失订单数

利润缺失订单数

数据完整状态
```

这样：

```
总利润 = <NA>
```

时，可以直接知道：

> 为什么无法计算。

------



## 五十六、v0.2 第五优先级：输出目录自动创建

未来可以：

```
from pathlib import Path

Path(
    "output"
).mkdir(
    parents=True,
    exist_ok=True
)
```

这样程序自己保证：

```
output
```

目录存在。

------



## 五十七、v0.2 第六优先级：输入文件健壮性

可以增加：

```
文件是否存在

Excel 是否为空

Sheet 是否存在

文件是否可以读取

字段类型是否异常
```

例如：

```
Path(
    "input/orders.xlsx"
).exists()
```

避免因为：

```
文件路径错误
文件没放进去
文件名变化
```

导致难以理解的报错。

------



## 五十八、v0.2 第七优先级：固定异常明细结构

当前无异常时：

```
all_abnormal_df = pd.DataFrame(
    columns=["异常类型"]
)
```

未来可以构建固定字段：

```
订单号
下单时间
SKU
数量
商品销售额
订单状态
商品名称
类目
站点
采购成本
包装成本
头程成本
异常类型
```

这样：

> 不管有没有异常，Excel Sheet 结构始终一致。

更方便：

```
自动化读取
人工检查
后续脚本二次处理
```

------



## 五十九、v0.2 第八优先级：配置集中管理

例如统一：

```
VALID_STATUS = [
    "已完成",
    "已取消",
    "处理中"
]
```

以及：

```
字段名称
输入路径
输出路径
合法订单状态
分析订单状态
```

未来甚至可以单独放：

```
config.py
```

但当前阶段不急着拆文件。

------



## 六十、v0.2 第九优先级：函数化重构

未来主流程可以逐步变成：

```
load_data()

check_structure()

clean_data()

check_master_data()

check_order_abnormal()

merge_product_data()

merge_cost_data()

calculate_profit()

build_sku_report()

build_daily_report()

build_abnormal_report()

export_report()
```

这样主程序最终可能变成：

```
读取
↓
检查
↓
清洗
↓
分析
↓
输出
```

更加接近真实项目结构。

但不建议现在立即大规模重构。

原因：

> 当前阶段主要目标仍然是掌握 pandas 和数据处理逻辑，而不是为了重构而重构。

------



## 六十一、v0.2 第十优先级：增加运营指标

未来可以逐步加入：

```
客单价

件单价

单均利润

SKU利润率

亏损订单数

亏损SKU数

亏损率

日均销售额

SKU销售贡献度

Top SKU

异常率

成本缺失率
```

这样项目才能从：

> 数据处理工具

进一步升级成：

> 运营分析工具。

------



## 六十二、v0.2 优先级建议

不建议一次全部实现。

推荐顺序：

### 第一阶段

```
业务分析状态口径

异常体系优化

数据完整性
```

目标：

> 让结果更加可信。

------

### 第二阶段

```
输入输出健壮性

固定报表结构

配置集中
```

目标：

> 让工具更加稳定。

------

### 第三阶段

```
函数化

模块化
```

目标：

> 提高代码可维护性。

------

### 第四阶段

```
新增运营指标
```

目标：

> 提高实际运营价值。

------



## 六十三、本项目最重要的 pandas 知识链

本项目真正需要带走的不是单独 API，而是：

```apl
read_excel
↓
结构检查
↓
string / to_numeric / to_datetime
↓
isna / duplicated / is_unique
↓
布尔筛选
↓
merge
↓
indicator
↓
mask + loc
↓
groupby + agg
↓
size / count / nunique
↓
自定义聚合函数
↓
日期维度处理
↓
ExcelWriter
```

最终串成：

```ABAP
数据清洗
↓
数据质量控制
↓
数据建模
↓
业务计算
↓
分析聚合
↓
自动化报告
```

---



## 六十四、当前阶段总结

这个项目最大的意义不是：

> 写出了多少行 pandas 代码。

而是第一次把之前零散学习的 pandas 知识串成了完整流程：

```apl
原始业务数据
↓
清洗
↓
检查
↓
关联
↓
计算
↓
分析
↓
报告
```

目前 v0.1 已经完成。

接下来不继续无限优化该项目。

后续路线：

```ABAP
20｜pandas 综合项目 v0.1
        ↓
完成项目笔记
        ↓
21｜openpyxl
        ↓
运费申诉自动化工具 v0.1
        ↓
积累真实业务项目
        ↓
再回头升级本项目 v0.2
```

------



## 六十五、当前项目状态

```ABAP
项目：
20｜pandas 综合项目：Shopee运营数据分析工具

版本：
v0.1

状态：
完成 / 封版

下一学习模块：
openpyxl

后续升级：
v0.2 暂不立即开发
```
