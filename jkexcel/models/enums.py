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
