from enum import Enum


class ExcelType(Enum):
    OFFICE = ("Excel.Application", "EXCEL.EXE", "Office")
    WPS = ("KET.Application", "wps.exe", "Wps")


class Platform(Enum):
    """
    指定生成文本文件的平台
    """
    xlMacintosh = 1  # macintosh
    xlMSDOS = 3  # MS-DOS
    xlWindows = 2  # Winodws


class SeparatorFormat(Enum):
    """
    sep_format 使用 确定文件的分隔符
    """
    Tab = 1  # \t
    Commas = 2  # 逗号
    Spaces = 3  # 空格
    Semicolons = 4  # 分号
    Noting = 5  # 无
    Custom = 6  # 自定义字符


class CorruptLoad(Enum):
    """
    指定文件打开时处理
    """
    xlNormalLoad = 0  # 正常打开工作簿
    xlRepairFile = 1  # 尝试修复文件
    xlExtractData = 2  # 尝试恢复工作簿中的数据


class BorderWeight(Enum):
    """边框粗细"""
    HAIRLINE = 1
    THIN = 2
    MEDIUM = -4138
    THICK = 4


class BorderLineStyle(Enum):
    """边框线型"""
    NONE = -4142
    CONTINUOUS = 1
    DASH = -4115
    DOT = -4118
    DASH_DOT = 4
    DASH_DOT_DOT = 5
    DOUBLE = -4119


class SortOrder(Enum):
    """排序顺序"""
    ASCENDING = 1
    DESCENDING = 2


class FilterOperator(Enum):
    """筛选运算符"""
    AND = 1
    OR = 2


class PageOrientation(Enum):
    """页面方向"""
    PORTRAIT = 1
    LANDSCAPE = 2


class PaperSize(Enum):
    """纸张大小"""
    LETTER = 1
    A4 = 9
    A5 = 11


class XlFindLookIn(Enum):
    """搜索的数据类型"""
    xlComments = -4144  # 注释
    xlCommentsThreaded = -4184  # 线程注释
    xlFormulas = -4123  # 公式
    xlValues = -4163  # 值


class XlLookAt(Enum):
    """
    匹配全部搜索文本还是匹配任一部分搜索文本
    """
    xlPart = 2  # 匹配任一部分搜索文本
    xlWhole = 1  # 匹配全部搜索文本


class XlSearchOrder(Enum):
    """
    指定搜索区域的次序
    """
    xlByColumns = 2  # 搜索列，然后移到下一列
    xlByRows = 1  # 搜索行，然后移到下一行


class XlSearchDirection(Enum):
    """
    指定搜索区域时的搜索方向
    """
    xlNext = 1  # 在区域中搜索下一匹配值
    xlPrevious = 2  # 在区域中搜索上一匹配值


class XlPasteType(Enum):
    """
    指定要粘贴的区域部分
    """
    xlPasteAll = -4104  # 粘贴所有内容
    xlPasteAllExceptBorders = 7  # 粘贴所有内容，除了边框
    xlPasteAllMergingConditionalFormats = 14  # 粘贴所有内容并合并所有条件格式
    xlPasteAllUsingSourceTheme = 13  # 粘贴所有内容并使用源主题
    xlPasteColumnWidths = 8  # 粘贴复制的列宽
    xlPasteComments = -4144  # 粘贴批注
    xlPasteFormats = -4122  # 粘贴复制的源格式
    xlPasteFormulas = -4123  # 粘贴公式
    xlPasteFormulasAndNumberFormats = 11  # 粘贴公式和数字格式
    xlPasteValidation = 6  # 粘贴有效性
    xlPasteValues = -4163  # 粘贴值
    xlPasteValuesAndNumberFormats = 12  # 粘贴值和数字格式


class XlPasteSpecialOperation(Enum):
    """
    指定如何使用工作表上的目标单元格计算数值数据
    """
    xlPasteSpecialOperationAdd = 2  # 复制的数据将添加到目标单元格中的值
    xlPasteSpecialOperationDivide = 5  # 复制的数据将除以目标单元格中的值
    xlPasteSpecialOperationMultiply = 4  # 复制的数据将乘以目标单元格中的值
    xlPasteSpecialOperationNone = -4142  # 粘贴操作中不执行任何计算
    xlPasteSpecialOperationSubtract = 3  # 复制的数据将从目标单元格中的值中减去


class XlCalculation(Enum):
    """计算模式"""
    xlCalculationAutomatic = -4105  # 自动计算
    xlCalculationManual = -4135  # 手动计算
    xlCalculationSemiautomatic = 2  # 半自动（除表格外自动）


class XlDirection(Enum):
    """移动方向（End 方法）"""
    xlUp = -4162
    xlDown = -4121
    xlToLeft = -4159
    xlToRight = -4161


class XlInsertShiftDirection(Enum):
    """插入单元格时的移动方向"""
    xlShiftDown = -4121  # 下移
    xlShiftToRight = -4161  # 右移


class XlDeleteShiftDirection(Enum):
    """删除单元格时的移动方向"""
    xlShiftUp = -4162  # 上移
    xlShiftToLeft = -4159  # 左移


class XlCellType(Enum):
    """SpecialCells 单元格类型"""
    xlCellTypeBlanks = 4  # 空单元格
    xlCellTypeConstants = 2  # 常量
    xlCellTypeFormulas = -4123  # 公式
    xlCellTypeVisible = 12  # 可见单元格
    xlCellTypeLastCell = 11  # 最后一个单元格
    xlCellTypeComments = -4144  # 含批注
    xlCellTypeAllValidation = -4174  # 含数据验证
    xlCellTypeAllFormatConditions = -4172  # 含条件格式


class XlAutoFilterOperator(Enum):
    """自动筛选运算符"""
    xlAnd = 1  # 两个条件逻辑与
    xlOr = 2  # 两个条件逻辑或
    xlTop10Items = 3  # 前 N 项
    xlBottom10Items = 4  # 后 N 项
    xlTop10Percent = 5  # 前 N 百分比
    xlBottom10Percent = 6  # 后 N 百分比
    xlFilterValues = 7  # 值列表筛选


class XlFormatConditionOperator(Enum):
    """条件格式单元格值规则运算符"""
    xlBetween = 1
    xlNotBetween = 2
    xlEqual = 3
    xlNotEqual = 4
    xlGreater = 5
    xlLess = 6
    xlGreaterEqual = 7
    xlLessEqual = 8


class XlIconSet(Enum):
    """条件格式图标集（常用子集）"""
    xl3Arrows = 1
    xl3ArrowsGray = 2
    xl3Flags = 3
    xl3TrafficLights1 = 4
    xl3TrafficLights2 = 5
    xl3Signs = 6
    xl3Symbols = 7
    xl3Symbols2 = 8
    xl4Arrows = 9
    xl4ArrowsGray = 10
    xl4CRV = 11
    xl4TrafficLights = 12
    xl5Arrows = 13
    xl5ArrowsGray = 14
    xl5CRV = 15
    xl5Quarters = 16


class XlBordersIndex(Enum):
    """边框位置"""
    xlDiagonalDown = 5
    xlDiagonalUp = 6
    xlEdgeLeft = 7
    xlEdgeTop = 8
    xlEdgeBottom = 9
    xlEdgeRight = 10
    xlInsideVertical = 11
    xlInsideHorizontal = 12


class XlUnderlineStyle(Enum):
    """字体下划线样式"""
    xlUnderlineStyleNone = -4142
    xlUnderlineStyleSingle = 2
    xlUnderlineStyleDouble = -4119
    xlUnderlineStyleSingleAccounting = 4
    xlUnderlineStyleDoubleAccounting = 5


class XlChartType(Enum):
    """图表类型（常用子集）"""
    xlColumnClustered = 51  # 簇状柱形图
    xlColumnStacked = 52  # 堆积柱形图
    xlColumnStacked100 = 53  # 百分比堆积柱形图
    xlBarClustered = 57  # 簇状条形图
    xlBarStacked = 58  # 堆积条形图
    xlLine = 4  # 折线图
    xlLineMarkers = 65  # 带数据标记的折线图
    xlLineStacked = 63  # 堆积折线图
    xlPie = 5  # 饼图
    xlPieExploded = 69  # 分离型饼图
    xlDoughnut = -4120  # 圆环图
    xlArea = 1  # 面积图
    xlAreaStacked = 76  # 堆积面积图
    xlXYScatter = -4169  # 散点图
    xlXYScatterLines = 74  # 带平滑线的散点图
    xlRadar = -4151  # 雷达图
    xlBubble = 15  # 气泡图


class XlPivotFieldOrientation(Enum):
    """数据透视表字段方向"""
    xlHidden = 0
    xlRowField = 1  # 行字段
    xlColumnField = 2  # 列字段
    xlPageField = 3  # 筛选字段
    xlDataField = 4  # 值字段


class XlConsolidationFunction(Enum):
    """数据透视表值字段汇总方式"""
    xlSum = -4157
    xlCount = -4112
    xlAverage = -4106
    xlMax = -4136
    xlMin = -4139
    xlProduct = -4149
