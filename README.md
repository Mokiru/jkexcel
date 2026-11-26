# JKExcel

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

基于 COM 接口的 Python Excel 自动化工具，同时支持 **Microsoft Excel** 和 **WPS 表格**，所有功能均在双引擎下验证通过。

## 特性

- 🎯 **完整封装**：Application / Workbook / Worksheet / Range 全对象模型，外加图表、数据透视表、事件监听
- 🔄 **双引擎支持**：Microsoft Excel（Excel.Application）与 WPS 表格（KET.Application）一套 API 通用
- 📊 **Pandas 集成**：DataFrame 与 Excel 区域互转，自动处理 NaN / NaT / 时间戳
- 🎨 **格式操作**：样式快照读取、边框、条件格式（色阶/数据条/图标集）、数据验证
- ⚡ **批量高性能**：10 万单元格批量写入约 0.5 秒，`suspend()` 上下文一键进入高性能模式
- 🛡️ **异常处理**：完善的异常体系，兼容动态绑定与 gen_py 早期绑定
- 📝 **上下文管理**：支持 `with` 语句，自动资源管理

## 性能

以下数据来自 [`test/benchmark.py`](test/benchmark.py) 在本机的实测结果
（Windows / Python 3.10.5 / Excel 16.0 与 WPS 表格，后台模式，多次运行取平均，仅供参考）：

| 操作 | Office | WPS | 说明 |
|---|---|---|---|
| 逐单元格写入 500 个（默认状态） | 3.63 s | 2.45 s | 500 次 COM 调用，反面示例 |
| 逐单元格写入 500 个（suspend） | 1.69 s | 2.23 s | 关闭屏幕刷新/事件/自动计算 |
| **批量写入 10 万单元格** | **529 ms** | **764 ms** | 单次 COM 调用 |
| **批量读取 10 万单元格** | **150 ms** | **243 ms** | 单次 COM 调用 |
| DataFrame 写入 1 万行 × 10 列 | 775 ms | 781 ms | 含 pandas 转换开销 |
| DataFrame 读取 1 万行 × 10 列 | 240 ms | 195 ms | 含 pandas 转换开销 |
| 应用样式+边框 1000 单元格（默认状态） | 419 ms | 86 ms | |
| 应用样式+边框 1000 单元格（suspend） | 61 ms | 72 ms | Office 下提速约 7 倍 |
| find_all 全表查找（10 万单元格） | 247 ms | 214 ms | |
| 双关键字排序 1 万行 | 779 ms | 216 ms | |
| last_row 定位末行 | 9.4 ms | 6.7 ms | End 模式，与数据量无关 |
| 保存 xlsx（1 万行 × 10 列） | 234 ms | 225 ms | |
| 打开 xlsx（1 万行 × 10 列） | 607 ms | 731 ms | |

**关键结论：**

1. **批量操作比逐单元格快 2~3 个数量级**。写入 500 个单元格逐个调用需要 3.6 秒，而一次批量写入 10 万个单元格只要 0.5 秒。请始终优先使用 `write_data()` / `get_values()` / `value = 二维列表`。
2. **`app.suspend()` 适合大量分散操作**（逐格写、循环设样式、多表操作），Office 下样式操作可提速约 7 倍；对单次大批量读写本身收益不大。
3. **定位末行用 `last_row()` 而非遍历**，耗时与数据量无关（约 10 ms）。
4. 与 openpyxl 等文件级库相比，COM 方案的优势在于**完整保真**（公式计算、图表、透视表、宏、打印导出都能操作），代价是需要安装 Excel/WPS 且仅限 Windows。纯数据读写场景两者速度相当；需要操作 Excel 对象模型时 COM 是唯一选择。

运行 benchmark：

```bash
python test/benchmark.py office   # 测试 Microsoft Excel
python test/benchmark.py wps      # 测试 WPS 表格
```

## 安装

### 环境要求

- Python 3.7+
- Windows 操作系统
- Microsoft Excel 或 WPS Office

### 安装依赖

```bash
pip install -r requirements.txt
```

或直接安装：

```bash
pip install pywin32 psutil pandas
```

## 快速开始

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()

    # 批量写入（推荐方式）
    ws.write_data(
        [["张三", 25], ["李四", 30]],
        headers=["姓名", "年龄"],
    )

    wb.save("output.xlsx")
# with 退出时自动 quit，释放 COM 资源
```

打开现有文件：

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.open_workbook("data.xlsx", read_only=True)
    ws = wb.worksheets["Sheet1"]
    print(ws.get_range("A1").value)
    print(ws.used_range.get_values())
```

## 各部分使用指南

### 1. 应用与配置（ExcelApp / ExcelConfig）

```python
from jkexcel import ExcelApp, ExcelConfig, ExcelType, XlCalculation

config = ExcelConfig(
    driver=ExcelType.OFFICE,   # OFFICE 或 WPS
    visible=True,              # 是否显示窗口
    display_alerts=False,      # 关闭警告弹窗（自动化推荐）
    screen_updating=True,      # 屏幕刷新
    enable_events=True,        # 事件响应
    window_state="maximized",  # normal / maximized / minimized
)

with ExcelApp(config) as app:
    print(app.version)         # Excel/WPS 版本
    print(app.is_running)

    wb = app.create_workbook()  # 设置 Calculation 需至少有一个打开的工作簿

    # 应用级开关（随时可读写）
    app.screen_updating = False
    app.calculation = XlCalculation.xlCalculationManual
    app.calculate()            # 手动触发全量计算

    # 运行宏（需文件允许宏）
    # app.run_macro("MyMacro", arg1, arg2)
```

**高性能模式 `suspend()`**：一次性关闭屏幕刷新/事件/警告并切手动计算，退出时自动恢复：

```python
with ExcelApp() as app:
    wb = app.create_workbook()
    with app.suspend():
        for i in range(1, 1001):
            wb.worksheets[1].get_cell(i, 1).value = i
    wb.save("fast.xlsx")
```

同一进程可同时持有 Office 和 WPS 两个实例（按驱动区分单例）：

```python
office = ExcelApp(ExcelConfig(driver=ExcelType.OFFICE)).start()
wps = ExcelApp(ExcelConfig(driver=ExcelType.WPS)).start()
# ... 分别操作 ...
ExcelApp.cleanup_all()   # 统一清理
```

### 2. 工作簿（Workbook）

```python
from jkexcel import ExcelApp, SaveFormat

with ExcelApp() as app:
    wb = app.create_workbook()
    # wb = app.open_workbook("data.xlsx", read_only=False, password=None)

    # 保存 / 另存为（指定格式时自动修正扩展名）
    wb.save("out.xlsx")
    wb.save_as("out.csv", file_format=SaveFormat.xlCSV)          # 实际保存为 out.csv
    wb.save_as("out", file_format=SaveFormat.xlOpenXMLWorkbook)  # 实际保存为 out.xlsx

    # 命名区域
    ws = wb.worksheets[1]
    wb.add_name("SalesData", ws.get_range("A1:C10"))
    print(wb.names)                          # ['SalesData']
    print(wb.get_name("SalesData").address)  # A1:C10
    wb.delete_name("SalesData")

    # 属性
    print(wb.name, wb.full_name, wb.saved, wb.read_only)

    wb.close(save_changes=True)
```

### 3. 工作表（Worksheets / Worksheet）

```python
from jkexcel import ExcelApp, SheetVisibility

with ExcelApp() as app:
    wb = app.create_workbook()

    # 获取：按名称 / 1 基索引 / 迭代
    ws = wb.worksheets[1]
    ws = wb.worksheets["Sheet1"]
    for sheet in wb.worksheets:
        print(sheet.name, sheet.index)

    # 增删 / 复制 / 移动
    ws2 = wb.worksheets.add(after=ws)
    ws2.name = "数据"
    ws3 = wb.worksheets.add(count=3)            # 批量添加，返回列表
    copied = ws.copy(after=ws2)                 # 复制
    ws2.move(before=ws)                         # 移动
    wb.worksheets.delete("Sheet3")

    # 可见性 / 标签颜色（深度隐藏的表无法激活，不影响其他表操作）
    ws2.visible = False
    ws2.set_visibility(SheetVisibility.VERY_HIDDEN)
    from jkexcel.utils.helpers import rgb_to_excel_color
    ws.tab_color = rgb_to_excel_color(255, 0, 0)  # Excel BGR 颜色值，None 清除

    # 行列操作
    ws.get_cells(1, 1, 1, 2).value = ["表头A", "表头B"]  # (起行, 起列, 止行, 止列)
    ws.hide_rows(5, 6)                          # 隐藏第 5~10 行（起始行, 行数）
    ws.hide_columns(2, 2, hidden=False)         # 取消隐藏 B~C 列
    ws.auto_fit_columns()

    # 数据边界（End 模式，大数据量下极快）
    print(ws.last_row(1), ws.last_column(1))    # 末行 / 末列
    print(ws.used_range.address)

    # 冻结窗格
    ws.freeze_panes("B2")                       # 冻结首行+首列
    ws.unfreeze_panes()

    # 清空整表（保留工作表本身）
    ws.clear()
```

### 4. 单元格与区域（Range）

```python
from jkexcel import ExcelApp, XlDirection, XlCellType, XlAutoFilterOperator

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]
    ws.write_data([[1, 2], [3, 4], [None, 6]])

    rg = ws.get_range("A1:B3")
    cell = ws.get_cell(1, 1)                    # 等价 ws.get_range("A1")

    # ---- 读写 ----
    print(rg.get_values())                      # 二维列表（推荐批量读取）
    rg.value = [[10, 20], [30, 40], [50, 60]]   # 批量写入
    cell.formula = "=SUM(A1:A3)"
    print(cell.formula)                         # 读公式
    print(cell.text, cell.row, cell.column, cell.address)

    # ---- 形状 / 索引 / 迭代 ----
    print(rg.count, rg.shape)                   # 6 (3, 2)
    print(rg[1, 2].value)                       # 相对索引（1 基）→ B1
    print(rg["A1"].value)                       # 绝对地址
    for c in rg:                                # 行优先迭代单元格
        print(c.address, c.value)

    # ---- 导航 ----
    print(cell.offset(1, 2).address)            # A1 → C2
    print(cell.end(XlDirection.xlDown).address) # Ctrl+↓ 跳到数据边界
    print(rg.entire_row.address)                # 整行 / entire_column 整列
    print(rg.resize(rows=5).address)            # 调整区域大小

    # ---- 插入 / 删除 / 合并 ----
    ws.get_range("B1").insert()                 # 插入单元格，默认下移
    ws.get_range("A5:B5").delete()              # 删除单元格，默认上移
    ws.get_range("A1:C1").merge()               # 合并（unmerge() 取消）
    print(rg.merge_cells)                       # 读取合并状态

    # ---- 特殊单元格定位 ----
    blanks = rg.special_cells(XlCellType.xlCellTypeBlanks)   # 空值格，无匹配返回 None
    formulas = rg.special_cells(XlCellType.xlCellTypeFormulas)

    # ---- 排序（支持 3 关键字）----
    data = ws.get_range("A1:C100")
    data.sort(ws.get_range("B1:B100"), order=1, header=2,
              key2=ws.get_range("C1:C100"), order2=2)

    # ---- 自动筛选 ----
    ws.get_range("A1:C100").auto_filter(field=2, criteria1=">10")
    ws.get_range("A1:C100").auto_filter(field=2, criteria1=">10",
                                        operator=XlAutoFilterOperator.xlAnd, criteria2="<50")
    ws.get_range("A1:C100").auto_filter()       # 无参切换开关（需作用于筛选区域）

    # ---- 复制 / 粘贴 ----
    ws.get_range("A1:B3").copy(ws.get_range("E1"))

    # ---- 数字格式 / 隐藏 ----
    rg.number_format = "0.00"
    rg.wrap_text = True
    rg.entire_row.hidden = True                 # 隐藏所在整行（整列用 entire_column）

    # ---- 查找 ----
    found = ws.find_all("目标文本")              # 返回 Range 列表
```

### 5. 数据操作（ws.data / DataOperations）

每个 Worksheet 自带 `data` 属性，无需手动构造：

```python
import pandas as pd
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]

    # DataFrame 写入（NaN/NaT 自动转空单元格，Timestamp 自动转 datetime）
    df = pd.DataFrame({"姓名": ["张三", "李四"], "分数": [90.5, None]})
    ws.data.write_from_dataframe(df, start_cell="A1", include_index=False)

    # 读取为 DataFrame（自动识别已用区域）
    df2 = ws.data.read_to_dataframe(has_headers=True)

    # 条件筛选：返回匹配行的 Range 列表，条件键支持表头名或 1 基列号（等值匹配）
    rows = ws.data.filter_data("A1:B3", {"姓名": "张三"})
    rows = ws.data.filter_data("A1:B3", {2: 90.5})

    # 重复值处理
    dups = ws.data.find_duplicates("A1:B3")     # 返回重复行 Range 列表
    ws.data.remove_duplicates("A1:B3", columns=[1, 2], has_headers=True)  # 原生去重
```

### 6. 格式与样式（RangeStyle / ws.format）

```python
from jkexcel import (
    ExcelApp, RangeStyle, BorderLineStyle, BorderWeight,
    XlFormatConditionOperator, XlIconSet, XlUnderlineStyle,
)
from jkexcel.utils.helpers import rgb_to_excel_color

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]
    ws.write_data([[10], [60], [120]])
    rg = ws.get_range("A1:A3")

    # ---- 样式设置与读取 ----
    style = RangeStyle(
        font_name="微软雅黑", font_size=12, font_bold=True,
        font_color_rgb=(255, 0, 0),             # RGB 元组，自动转 Excel BGR
        fill_color=RangeStyle.Colors.YELLOW,    # 或内置颜色常量
        horizontal_alignment=RangeStyle.Alignment.CENTER,
        font_underline=XlUnderlineStyle.xlUnderlineStyleSingle,
        borders={"all": {"line_style": BorderLineStyle.CONTINUOUS,
                         "weight": BorderWeight.THIN, "color": 0}},
    )
    rg.apply_style(style)
    snapshot = rg.get_style()                   # 读取当前样式快照
    print(snapshot.font_bold, snapshot.fill_color)

    # ---- 条件格式（FormatOperations 统一接收地址字符串）----
    fmt = ws.format
    fmt.add_cell_value_rule("A1:A3", XlFormatConditionOperator.xlGreater, "50",
                            style=RangeStyle(font_color=RangeStyle.Colors.RED))
    fmt.add_cell_value_rule("A1:A3", XlFormatConditionOperator.xlBetween, "10", "60",
                            style=RangeStyle(fill_color=RangeStyle.Colors.GREEN))
    fmt.add_data_bar("A1:A3", fill_color=rgb_to_excel_color(99, 190, 123))  # 数据条
    fmt.add_color_scale("A1:A3",                                        # 双色色阶
                        min_color=rgb_to_excel_color(248, 105, 107),
                        max_color=rgb_to_excel_color(99, 190, 123))
    fmt.add_icon_set("A1:A3", XlIconSet.xl3TrafficLights1)              # 图标集
    fmt.apply_conditional_formatting("A1:A3", "=A1>100",
                                     RangeStyle(font_bold=True))        # 公式规则
    rg.clear_conditional_formats()                                      # 清除规则

    # ---- 数据验证 ----
    fmt.apply_data_validation(
        "B1:B10",
        validation_type=1,          # xlValidateWholeNumber
        formula1="0", formula2="100",
        error_title="输入错误", error_message="请输入 0-100 之间的数字",
    )
    ws.get_range("B1:B10").clear_validation()

    # ---- 其他 ----
    fmt.apply_number_format("A1:A3", "0.00")
    fmt.apply_borders("A1:A3", line_style=BorderLineStyle.CONTINUOUS,
                      weight=BorderWeight.THIN, edge="all")

    # 套用表格样式（区域需含表头和数据）
    ws2 = wb.worksheets.add()
    ws2.write_data([["甲", 1], ["乙", 2]], headers=["名称", "数量"])
    ws2.format.apply_table_format("A1:B3", "TableStyleMedium9")
```

### 7. 批注与超链接

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]

    cell = ws.get_range("A1")
    cell.add_comment("这是批注")
    print(cell.comment)                # 读取批注文本，无批注返回 None
    cell.delete_comment()

    cell.add_hyperlink("https://example.com", text_to_display="官网")
    ws.get_range("A2").add_hyperlink("", sub_address="Sheet2!A1",
                                     text_to_display="跳转到 Sheet2")
    cell.clear_hyperlinks()
```

### 8. 图片与图表

```python
from jkexcel import ExcelApp, XlChartType

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]
    ws.write_data([["苹果", 10], ["香蕉", 5], ["白菜", 3]],
                  headers=["名称", "数量"])

    # 插入图片（按单元格定位，-1 保持原始尺寸）
    ws.add_picture("logo.png", cell="E2", width=120)
    print(ws.shapes.Count)

    # 图表
    chart = ws.add_chart(XlChartType.xlColumnClustered,
                         left=200, top=20, width=400, height=260)
    chart.title = "销量统计"
    chart.set_source_data(ws.get_range("A1:B4"))
    chart.has_legend = False
    chart.export_image("chart.png")    # 导出为 PNG
    # chart.delete()
```

### 9. 数据透视表

```python
from jkexcel import ExcelApp, XlConsolidationFunction

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]
    ws.write_data([["苹果", "水果", 10], ["香蕉", "水果", 5], ["白菜", "蔬菜", 3]],
                  headers=["名称", "类别", "数量"])

    pt = ws.create_pivot_table(
        source_range="A1:C4",          # 地址字符串，首行必须为表头
        dest_cell="F1",
        name="销量透视",
        row_fields=["类别"],           # 也可建表时直接指定字段
        data_fields=["数量"],
    )
    # 或建表后逐步配置：
    # pt.add_row_field("类别")
    # pt.add_data_field("数量", caption="总数量",
    #                   function=XlConsolidationFunction.xlSum)
    pt.refresh()                       # 源数据变化后刷新
    print(pt.table_range.address)
    # pt.delete()
```

### 10. 页面设置与导出

```python
from jkexcel import ExcelApp, PageOrientation, PaperSize

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]
    ws.write_data([[i, i * 2] for i in range(1, 101)], headers=["A", "B"])

    # 页面设置（边距单位：厘米）
    ws.print_area = "A1:B100"
    ws.set_print_titles(rows="1:1")            # 每页重复标题行
    ws.set_page_setup(
        orientation=PageOrientation.LANDSCAPE,
        paper_size=PaperSize.A4,
        fit_to_pages_wide=1,                   # 宽 1 页
        fit_to_pages_tall=0,                   # 高不限
        left_margin=1.5, right_margin=1.5,
        top_margin=2, bottom_margin=2,
        center_horizontally=True,
    )
    ws.set_header(center="销售报表")            # left/center/right 三段
    ws.set_footer(center="第 &P 页 / 共 &N 页")  # 支持 &P &N &D 等代码

    # 导出 PDF（表级 / 工作簿级）
    ws.export_as_pdf("report.pdf")
    wb.export_as_pdf("workbook.pdf")

    # 打印
    # ws.print_out(copies=1, preview=True)
```

### 11. 事件监听

```python
import time
from jkexcel import ExcelApp, ExcelEvents

class MyEvents(ExcelEvents):
    def on_sheet_change(self, sh, target):
        print(f"单元格 changed: {target.Address}")

    def on_sheet_selection_change(self, sh, target):
        print(f"选区变化: {target.Address}")

    def on_workbook_before_save(self, wb, save_as_ui, cancel):
        print("即将保存")

with ExcelApp() as app:
    handler = MyEvents()
    app.attach_events(handler)

    wb = app.create_workbook()
    wb.worksheets[1].get_range("A1").value = "触发事件"

    # 回调在消息泵中触发：交互使用时 Excel 自身消息循环即可，
    # 后台自动化需主动泵消息
    for _ in range(10):
        app.pump_events()
        time.sleep(0.1)

    app.detach_events()
```

可重写的事件：`on_new_workbook`、`on_workbook_open`、`on_workbook_before_save`、`on_workbook_before_close`、`on_sheet_change`、`on_sheet_selection_change`、`on_sheet_activate`、`on_sheet_calculate`、`on_workbook_activate`。

### 12. 异常处理

```python
from jkexcel import (
    ExcelApp, ExcelCOMError, ExcelNotRunningError,
    WorkbookNotFoundError, WorksheetNotFoundError, RangeError,
)

try:
    with ExcelApp() as app:
        wb = app.open_workbook("不存在的文件.xlsx")
except WorkbookNotFoundError as e:
    print(f"工作簿未找到: {e}")
except WorksheetNotFoundError as e:
    print(f"工作表未找到: {e}")
except RangeError as e:
    print(f"范围错误: {e}")
except ExcelNotRunningError as e:
    print(f"Excel 未运行: {e}")
except ExcelCOMError as e:
    print(f"COM 错误: {e}")
```

## 性能优化建议

```python
from jkexcel import ExcelApp, XlCalculation

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.worksheets[1]

    # ✅ 推荐：一次批量写入
    ws.write_data(big_data)                     # 10 万格 ≈ 0.5s

    # ❌ 避免：循环逐格写入
    # for r, row in enumerate(big_data, 1):     # 10 万格要几分钟
    #     for c, v in enumerate(row, 1):
    #         ws.get_cell(r, c).value = v

    # ✅ 大量分散操作包在 suspend 里
    with app.suspend():
        for ws_i in wb.worksheets:
            ws_i.get_range("A1:Z100").apply_style(style)

    # ✅ 只改数据不改公式时，手动计算模式避免每次写入触发重算
    app.calculation = XlCalculation.xlCalculationManual
    ws.write_data(big_data)
    app.calculate()                             # 最后统一算一次
    app.calculation = XlCalculation.xlCalculationAutomatic
```

## 注意事项

1. **Windows 平台**：本库基于 COM，仅支持 Windows，且需安装 Microsoft Excel 或 WPS Office
2. **单线程使用**：COM 接口不支持跨线程共享对象，请在同一线程中操作；多线程场景每线程独立创建实例并调用 `pythoncom.CoInitialize()`
3. **重复打开同一文件**：Excel 对已打开的路径再次 `open_workbook()` 会返回同一个工作簿对象，关闭它会同时影响先前持有的引用
4. **资源管理**：优先使用 `with` 语句；异常退出时可用 `ExcelApp.cleanup_all()` 兜底清理
5. **WPS 差异**：核心 API 双引擎行为一致（含事件、透视表、PDF 导出），但部分高级特性（如个别图表类型、图标集样式）渲染效果可能有差异
6. **绑定模式**：自动兼容 win32com 动态绑定与 gen_py 早期绑定，无需手动配置

## 测试

所有功能验证均要求 **Office 与 WPS 双环境通过**：

```bash
# 功能冒烟测试（21 项）
python test/test_new_features.py office
python test/test_new_features.py wps

# README 示例代码验证（12 节，保证文档与实现一致）
python test/test_readme_examples.py office
python test/test_readme_examples.py wps

# 性能基准测试
python test/benchmark.py office
python test/benchmark.py wps
```

## 许可证

MIT License

## 作者

Mokiru

## 项目链接

- GitHub: https://github.com/Mokiru/jkexcel

## 贡献

欢迎提交 Issue 和 Pull Request！