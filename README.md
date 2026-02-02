# JKExcel

[![Python](https://img.shields.io/badge/Python-3.7%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/License-MIT-green.svg)](LICENSE)
[![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey.svg)](https://www.microsoft.com/windows)

基于 COM 接口的 Python Excel 自动化工具，支持 Microsoft Excel 和 WPS 表格。

## 特性

- 🚀 **简单易用**：提供简洁的 Python API，无需深入了解 COM 编程
- 🎯 **完整封装**：全面封装 Excel 对象模型（Application、Workbook、Worksheet、Range）
- 🔄 **双引擎支持**：同时支持 Microsoft Excel 和 WPS 表格
- 📊 **Pandas 集成**：原生支持 pandas DataFrame 读写
- 🎨 **格式操作**：丰富的单元格格式和样式设置功能
- ⚡ **高性能**：基于 COM 接口，操作效率高
- 🛡️ **异常处理**：完善的异常处理机制
- 📝 **上下文管理**：支持 `with` 语句，自动资源管理

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
pip install pywin32 psutil pandas win32security
```

## 快速开始

### 基本使用

```bash
pip install jkexcel
```

```python
from jkexcel import ExcelApp, ExcelConfig

# 创建并启动 Excel 应用
app = ExcelApp()
app.start()

# 创建新工作簿
wb = app.create_workbook()

# 获取活动工作表
ws = wb.get_active_sheet()

# 写入数据
ws.get_range("A1").value = "Hello"
ws.get_range("B1").value = "World"

# 保存工作簿
wb.save("C:/path/to/output.xlsx")

# 关闭应用
app.quit()
```

### 使用上下文管理器

```python
from jkexcel import ExcelApp, ExcelConfig

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()
    ws.get_range("A1").value = "自动关闭"
    wb.save("output.xlsx")
```

### 打开现有工作簿

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    # 打开现有工作簿
    wb = app.open_workbook("C:/path/to/existing.xlsx")
    
    # 获取工作表
    ws = wb.worksheets["Sheet1"]
    
    # 读取数据
    value = ws.get_range("A1").value
    print(value)
```

## 核心功能

### 1. 应用配置

`ExcelConfig` 类提供应用相关配置：

```python
from jkexcel import ExcelApp, ExcelConfig
from jkexcel.models.enums import ExcelType

# 创建配置
config = ExcelConfig(
    driver=ExcelType.OFFICE,  # OFFICE 或 WPS
    visible=True,              # 是否显示窗口
    display_alerts=False,      # 是否显示警告弹窗
    screen_updating=True,      # 是否刷新屏幕
    enable_events=True,        # 是否启用事件
    user_control=True,         # 是否允许用户控制
    window_state="normal",    # 窗口状态: normal/maximized/minimized
    read_only_recommended=False,  # 是否只读推荐
    update_links=True          # 是否更新链接
)

# 使用配置启动
app = ExcelApp(config)
app.start()
```

### 2. 工作簿操作

```python
from jkexcel import ExcelApp, Workbook

with ExcelApp() as app:
    # 创建新工作簿
    wb = app.create_workbook()
    
    # 打开现有工作簿
    wb = app.open_workbook("path/to/file.xlsx")
    
    # 保存工作簿
    wb.save("path/to/save.xlsx")
    
    # 另存为其他格式
    from jkexcel.models.config import SaveFormat
    wb.save_as("output.csv", SaveFormat.xlCSV)
    
    # 关闭工作簿
    wb.close(save_changes=True)
    
    # 获取工作簿信息
    print(wb.name)        # 工作簿名称
    print(wb.full_name)   # 完整路径
    print(wb.path)        # 文件路径
    print(wb.saved)       # 是否已保存
    print(wb.read_only)   # 是否只读
```

### 3. 工作表操作

```python
from jkexcel import ExcelApp
from jkexcel.models.config import SheetVisibility

with ExcelApp() as app:
    wb = app.create_workbook()
    
    # 获取工作表
    ws = wb.get_active_sheet()           # 获取活动工作表
    ws = wb.worksheets["Sheet1"]          # 通过名称获取
    ws = wb.worksheets[1]                 # 通过索引获取
    
    # 工作表属性
    print(ws.name)        # 工作表名称
    print(ws.index)       # 工作表索引
    print(ws.visible)     # 是否可见
    
    # 设置工作表名称
    ws.name = "新名称"
    
    # 设置可见性
    ws.visible = False
    ws.set_visibility(SheetVisibility.VERY_HIDDEN)
    
    # 获取已使用的范围
    used_range = ws.used_range
    print(used_range.address)
    
    # 添加新工作表
    new_ws = wb.worksheets.add("新工作表")
    
    # 删除工作表
    wb.worksheets.delete("Sheet1")
```

### 4. 单元格和范围操作

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()
    
    # 获取范围
    cell = ws.get_range("A1")              # 单个单元格
    range_obj = ws.get_range("A1:C10")    # 多个单元格
    
    # 读写值
    cell.value = "Hello"
    print(cell.value)
    
    # 读写公式
    cell.formula = "=SUM(A1:A10)"
    print(cell.formula)
    
    # 批量写入
    data = [["姓名", "年龄"], ["张三", 25], ["李四", 30]]
    range_obj.value = data
    
    # 获取文本
    print(cell.text)
    
    # 获取行列信息
    print(cell.row)        # 行号
    print(cell.column)     # 列号
    print(range_obj.rows.count)    # 行数
    print(range_obj.columns.count) # 列数
    
    # 清除内容
    cell.clear_contents()
    range_obj.clear()
```

### 5. Pandas 集成

```python
import pandas as pd
from jkexcel import ExcelApp, DataOperations

with ExcelApp() as app:
    wb = app.open_workbook("data.xlsx")
    ws = wb.worksheets["Sheet1"]
    
    # 创建数据操作对象
    data_ops = DataOperations(ws)
    
    # 读取为 DataFrame
    df = data_ops.read_to_dataframe(start_cell="A1", has_headers=True)
    print(df)
    
    # 从 DataFrame 写入
    new_df = pd.DataFrame({
        "姓名": ["王五", "赵六"],
        "年龄": [28, 35]
    })
    data_ops.write_from_dataframe(
        new_df,
        start_cell="D1",
        include_index=False,
        include_headers=True
    )
```

### 6. 格式操作

```python
from jkexcel import ExcelApp, FormatOperations, RangeStyle

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()
    
    # 创建格式操作对象
    format_ops = FormatOperations(ws)
    
    # 应用表格格式
    format_ops.apply_table_format("A1:D10", "TableStyleMedium9")
    
    # 应用数字格式
    format_ops.apply_number_format("B2:B10", "0.00")
    format_ops.apply_number_format("C2:C10", "yyyy-mm-dd")
    
    # 应用条件格式
    style = RangeStyle(
        font_bold=True,
        font_color=RangeStyle.Colors.RED,
        fill_color=RangeStyle.Colors.YELLOW
    )
    format_ops.apply_conditional_formatting(
        "B2:B10",
        "=B2>100",
        style
    )
    
    # 应用数据验证
    format_ops.apply_data_validation(
        "B2:B10",
        validation_type=1,  # xlValidateWholeNumber
        formula1="0",
        formula2="100",
        error_title="输入错误",
        error_message="请输入0-100之间的数字"
    )
```

### 7. 单元格样式

```python
from jkexcel import ExcelApp, RangeStyle

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()
    cell = ws.get_range("A1")
    
    # 创建样式
    style = RangeStyle(
        font_name="微软雅黑",
        font_size=12,
        font_bold=True,
        font_italic=False,
        font_color=RangeStyle.Colors.BLUE,
        fill_color=RangeStyle.Colors.LIGHT_GRAY,
        horizontal_alignment=RangeStyle.Alignment.CENTER,
        vertical_alignment=RangeStyle.Alignment.MIDDLE,
        wrap_text=True
    )
    
    # 应用样式
    cell.apply_style(style)
    
    # 设置边框
    style.borders = {
        "left": {"style": 1, "color": 0x000000},
        "right": {"style": 1, "color": 0x000000},
        "top": {"style": 1, "color": 0x000000},
        "bottom": {"style": 1, "color": 0x000000}
    }
    cell.apply_style(style)
```

## API 参考

### 核心类

#### ExcelApp

Excel 应用程序主类，用于管理 Excel 实例。

**主要方法：**
- `start()` - 启动 Excel 应用
- `quit(force=False)` - 退出 Excel 应用
- `create_workbook()` - 创建新工作簿
- `open_workbook(filepath)` - 打开现有工作簿
- `close_all_workbooks(save_changes=False)` - 关闭所有工作簿
- `calculate()` - 强制计算所有公式
- `run_macro(macro_name, *args)` - 运行宏

**主要类方法：**
- `ExcelApp.get_all_instances()` - 获取所有 ExcelApp 实例
- `ExcelApp.close_all_instances(force=False)` - 关闭所有 ExcelApp 实例

**主要属性：**
- `is_running` - Excel 是否在运行
- `version` - Excel 版本
- `workbooks` - 工作簿集合
- `active_workbook` - 活动工作簿
- `active_sheet` - 活动工作表

#### Workbook

工作簿类，表示一个 Excel 文件。

**主要方法：**
- `save(filepath=None)` - 保存工作簿
- `save_as(filepath, file_format)` - 另存为
- `close(save_changes=False)` - 关闭工作簿
- `get_active_sheet()` - 获取活动工作表

**主要属性：**
- `name` - 工作簿名称
- `full_name` - 完整路径
- `path` - 文件路径
- `saved` - 是否已保存
- `read_only` - 是否只读
- `worksheets` - 工作表集合

#### Worksheet

工作表类，表示一个工作表。

**主要方法：**
- `get_range(address)` - 获取范围
- `read_data(start_cell, has_headers)` - 读取数据
- `write_data(data, start_cell, headers)` - 写入数据
- `set_visibility(visibility)` - 设置可见性

**主要属性：**
- `name` - 工作表名称
- `index` - 工作表索引
- `visible` - 是否可见
- `used_range` - 已使用的范围

#### Range

范围类，表示一个或多个单元格。

**主要方法：**
- `apply_style(style)` - 应用样式
- `clear()` - 清除所有
- `clear_contents()` - 清除内容

**主要属性：**
- `address` - 范围地址
- `value` - 值
- `formula` - 公式
- `text` - 文本
- `row` - 起始行
- `column` - 起始列
- `rows` - 行集合
- `columns` - 列集合

### 操作类

#### DataOperations

数据操作类，提供数据读写和处理功能。

**主要方法：**
- `read_to_dataframe(start_cell, has_headers)` - 读取为 DataFrame
- `write_from_dataframe(df, start_cell, include_index, include_headers)` - 从 DataFrame 写入
- `filter_data(criteria_range, criteria)` - 筛选数据
- `find_duplicates(search_range, columns)` - 查找重复值
- `remove_duplicates(search_range, columns)` - 删除重复值

#### FormatOperations

格式操作类，提供格式设置功能。

**主要方法：**
- `apply_table_format(range_address, style_name)` - 应用表格格式
- `apply_number_format(range_address, number_format)` - 应用数字格式
- `apply_conditional_formatting(range_address, formula, style)` - 应用条件格式
- `apply_data_validation(range_address, validation_type, ...)` - 应用数据验证

### 配置和枚举

#### ExcelConfig

Excel 应用配置类。

**配置项：**
- `driver` - 驱动类型（OFFICE/WPS）
- `visible` - 是否可见
- `display_alerts` - 是否显示警告
- `screen_updating` - 是否刷新屏幕
- `enable_events` - 是否启用事件
- `user_control` - 是否允许用户控制
- `window_state` - 窗口状态
- `read_only_recommended` - 是否只读推荐
- `update_links` - 是否更新链接

#### RangeStyle

范围样式配置类。

**样式属性：**
- `font_name` - 字体名称
- `font_size` - 字体大小
- `font_bold` - 是否加粗
- `font_italic` - 是否斜体
- `font_color` - 字体颜色
- `fill_color` - 填充颜色
- `horizontal_alignment` - 水平对齐
- `vertical_alignment` - 垂直对齐
- `number_format` - 数字格式
- `wrap_text` - 是否自动换行
- `borders` - 边框设置

**内置颜色：**
- `RangeStyle.Colors.BLACK`
- `RangeStyle.Colors.WHITE`
- `RangeStyle.Colors.RED`
- `RangeStyle.Colors.GREEN`
- `RangeStyle.Colors.BLUE`
- `RangeStyle.Colors.YELLOW`
- `RangeStyle.Colors.ORANGE`
- `RangeStyle.Colors.GRAY`
- `RangeStyle.Colors.LIGHT_GRAY`

**对齐方式：**
- `RangeStyle.Alignment.LEFT`
- `RangeStyle.Alignment.CENTER`
- `RangeStyle.Alignment.RIGHT`
- `RangeStyle.Alignment.TOP`
- `RangeStyle.Alignment.MIDDLE`
- `RangeStyle.Alignment.BOTTOM`
- `RangeStyle.Alignment.JUSTIFY`

#### SaveFormat

保存格式枚举。

**常用格式：**
- `SaveFormat.xlOpenXMLWorkbook` - .xlsx
- `SaveFormat.xlOpenXMLWorkbookMacroEnabled` - .xlsm
- `SaveFormat.xlCSV` - .csv
- `SaveFormat.xlWorkbookNormal` - .xls

#### SheetVisibility

工作表可见性枚举。

**值：**
- `SheetVisibility.VISIBLE` - 可见
- `SheetVisibility.HIDDEN` - 隐藏
- `SheetVisibility.VERY_HIDDEN` - 深度隐藏

## 异常处理

```python
from jkexcel import ExcelApp
from jkexcel.models.exceptions import (
    ExcelCOMError,
    ExcelNotRunningError,
    WorkbookNotFoundError,
    WorksheetNotFoundError,
    RangeError
)

try:
    app = ExcelApp()
    app.start()
    # ... 操作代码 ...
except ExcelNotRunningError as e:
    print(f"Excel 未运行: {e}")
except WorkbookNotFoundError as e:
    print(f"工作簿未找到: {e}")
except WorksheetNotFoundError as e:
    print(f"工作表未找到: {e}")
except RangeError as e:
    print(f"范围错误: {e}")
except ExcelCOMError as e:
    print(f"COM 错误: {e}")
finally:
    if app.is_running:
        app.quit()
```

## 高级用法

### 多实例管理

```python
from jkexcel import ExcelApp, ExcelConfig
from jkexcel.models.enums import ExcelType

# 创建 Office Excel 实例
office_config = ExcelConfig(driver=ExcelType.OFFICE)
office_app = ExcelApp(office_config)
office_app.start()

# 创建 WPS 实例
wps_config = ExcelConfig(driver=ExcelType.WPS)
wps_app = ExcelApp(wps_config)
wps_app.start()

# ... 分别操作 ...

# 获取所有实例
instances = ExcelApp.get_all_instances()
print(f"当前有 {len(instances)} 个 Excel 实例在运行")

# 关闭所有实例
closed_count = ExcelApp.close_all_instances()
print(f"已关闭 {closed_count} 个 Excel 实例")
```

### 批量操作

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    # 批量创建工作簿
    for i in range(5):
        wb = app.create_workbook()
        ws = wb.get_active_sheet()
        ws.get_range("A1").value = f"工作簿 {i+1}"
        wb.save(f"workbook_{i+1}.xlsx")
        wb.close()
```

### 公式操作

```python
from jkexcel import ExcelApp

with ExcelApp() as app:
    wb = app.create_workbook()
    ws = wb.get_active_sheet()
    
    # 写入数据
    ws.get_range("A1").value = 10
    ws.get_range("A2").value = 20
    ws.get_range("A3").value = 30
    
    # 使用公式
    ws.get_range("B1").formula = "=A1*2"
    ws.get_range("B2").formula = "=A2*2"
    ws.get_range("B3").formula = "=A3*2"
    
    # 求和
    ws.get_range("A4").formula = "=SUM(A1:A3)"
    
    # 强制计算
    app.calculate()
    
    # 获取结果
    print(ws.get_range("A4").value)
```

## 注意事项

1. **Windows 平台**：本库仅支持 Windows 操作系统
2. **Excel/WPS 安装**：需要安装 Microsoft Excel 或 WPS Office
3. **资源管理**：建议使用上下文管理器（`with` 语句）确保资源正确释放
4. **性能优化**：批量操作时建议关闭屏幕刷新（`screen_updating=False`）
5. **异常处理**：建议使用 try-except 块捕获可能的异常
6. **多线程**：COM 接口不支持多线程，请在单线程环境中使用

## 许可证

MIT License

## 作者

Mokiru

## 项目链接

- GitHub: https://github.com/Mokiru/jkexcel

## 贡献

欢迎提交 Issue 和 Pull Request！

