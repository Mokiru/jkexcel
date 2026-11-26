from typing import List, Tuple, Optional, Dict, Any, Union

from jkexcel.core.worksheet import Worksheet
from jkexcel.models.config import RangeStyle
from jkexcel.models.enums import XlFormatConditionOperator, XlIconSet, BorderLineStyle, BorderWeight
from jkexcel.utils.validators import validate_range_address


class FormatOperations:
    """格式操作类"""

    def __init__(self, worksheet: Worksheet):
        """
        初始化

        Args:
            worksheet: Worksheet 对象
        """
        self.ws = worksheet

    def apply_table_format(self,
                           range_address: str,
                           style_name: str = "TableStyleMedium9"):
        """
        应用表格格式

        Args:
            range_address: 范围地址
            style_name: 表格样式名称
        """
        validate_range_address(range_address)

        table_range = self.ws.get_range(range_address)

        # 添加表格
        table = self.ws.com_object.ListObjects.Add(
            SourceType=1,  # xlSrcRange
            Source=table_range.com_object,
            XlListObjectHasHeaders=1
        )

        # 应用样式
        table.TableStyle = style_name

    def apply_number_format(self,
                            range_address: str,
                            number_format: str):
        """
        应用数字格式

        Args:
            range_address: 范围地址
            number_format: 数字格式
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)
        cell_range.com_object.NumberFormat = number_format

    def apply_conditional_formatting(self,
                                     range_address: str,
                                     formula: str,
                                     style: RangeStyle):
        """
        应用条件格式

        Args:
            range_address: 范围地址
            formula: 条件公式
            style: 样式
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)

        # 添加条件格式
        cf = cell_range.com_object.FormatConditions.Add(
            Type=2,  # xlExpression
            Formula1=formula
        )

        # 应用样式
        font = cf.Font
        interior = cf.Interior

        if style.font_bold is not None:
            font.Bold = style.font_bold
        if style.font_color is not None:
            font.Color = style.font_color
        if style.fill_color is not None:
            interior.Color = style.fill_color

    def apply_data_validation(self,
                              range_address: str,
                              validation_type: int,
                              formula1: str = "",
                              formula2: str = "",
                              ignore_blank: bool = True,
                              show_error: bool = True,
                              error_title: str = "",
                              error_message: str = ""):
        """
        应用数据验证

        Args:
            range_address: 范围地址
            validation_type: 验证类型
            formula1: 公式1
            formula2: 公式2
            ignore_blank: 是否忽略空值
            show_error: 是否显示错误
            error_title: 错误标题
            error_message: 错误消息
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)

        dv = cell_range.com_object.Validation
        dv.Add(
            Type=validation_type,
            Formula1=formula1,
            Formula2=formula2
        )

        dv.IgnoreBlank = ignore_blank
        dv.ShowError = show_error

        if error_title:
            dv.ErrorTitle = error_title
        if error_message:
            dv.ErrorMessage = error_message

    def add_cell_value_rule(self,
                            range_address: str,
                            operator: XlFormatConditionOperator,
                            formula1: str,
                            formula2: str = None,
                            style: RangeStyle = None):
        """
        添加单元格值条件格式（如：大于 100 标红）

        Args:
            range_address: 范围地址
            operator: 比较运算符
            formula1: 比较值或公式（如 "100"）
            formula2: 第二比较值（xlBetween/xlNotBetween 时必填）
            style: 命中时应用的样式
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)
        # 位置参数：Type, Operator, Formula1, Formula2（1=xlCellValue）
        cf = cell_range.com_object.FormatConditions.Add(
            1, operator.value, formula1,
            formula2 if formula2 is not None else ""
        )
        self._apply_cf_style(cf, style)
        return cf

    def add_data_bar(self,
                     range_address: str,
                     fill_color: int = 0xFF8000):
        """
        添加数据条条件格式

        Args:
            range_address: 范围地址
            fill_color: 数据条颜色（Excel 颜色值，注意 BGR 顺序，可用 helpers.rgb_to_excel_color）
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)
        bar = cell_range.com_object.FormatConditions.AddDatabar()
        bar.BarColor.Color = fill_color
        return bar

    def add_color_scale(self,
                        range_address: str,
                        min_color: int,
                        mid_color: int = None,
                        max_color: int = None):
        """
        添加色阶条件格式

        Args:
            range_address: 范围地址
            min_color: 最小值颜色
            mid_color: 中间值颜色（None 表示双色色阶）
            max_color: 最大值颜色（双色色阶时作为最大值颜色）
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)
        if mid_color is not None:
            scale = cell_range.com_object.FormatConditions.AddColorScale(3)
            scale.ColorScaleCriteria(1).FormatColor.Color = min_color
            scale.ColorScaleCriteria(2).FormatColor.Color = mid_color
            scale.ColorScaleCriteria(3).FormatColor.Color = max_color if max_color is not None else min_color
        else:
            scale = cell_range.com_object.FormatConditions.AddColorScale(2)
            scale.ColorScaleCriteria(1).FormatColor.Color = min_color
            scale.ColorScaleCriteria(2).FormatColor.Color = max_color if max_color is not None else min_color
        return scale

    def add_icon_set(self,
                     range_address: str,
                     icon_set: XlIconSet = XlIconSet.xl3TrafficLights1):
        """
        添加图标集条件格式

        Args:
            range_address: 范围地址
            icon_set: 图标集样式
        """
        validate_range_address(range_address)

        cell_range = self.ws.get_range(range_address)
        cf = cell_range.com_object.FormatConditions.AddIconSetCondition()
        # IconSets 挂在 Workbook 上（类型库定义），而非 Application
        cf.IconSet = self.ws.com_object.Parent.IconSets(icon_set.value)
        return cf

    def clear_conditional_formatting(self, range_address: str):
        """清除范围内所有条件格式"""
        validate_range_address(range_address)
        self.ws.get_range(range_address).clear_conditional_formats()

    def clear_data_validation(self, range_address: str):
        """清除范围内数据验证"""
        validate_range_address(range_address)
        self.ws.get_range(range_address).clear_validation()

    def apply_borders(self,
                      range_address: str,
                      line_style: BorderLineStyle = BorderLineStyle.CONTINUOUS,
                      weight: BorderWeight = BorderWeight.THIN,
                      color: int = None,
                      edge: str = "all"):
        """
        应用边框

        Args:
            range_address: 范围地址
            line_style: 线型
            weight: 粗细
            color: 颜色（Excel 颜色值）
            edge: 边框位置 "all"/"left"/"right"/"top"/"bottom"/"inside_vertical"/"inside_horizontal"
        """
        validate_range_address(range_address)

        spec = {"line_style": line_style, "weight": weight}
        if color is not None:
            spec["color"] = color
        self.ws.get_range(range_address).apply_style(RangeStyle(borders={edge: spec}))

    @staticmethod
    def _apply_cf_style(cf, style: Optional[RangeStyle]):
        """将 RangeStyle 应用到条件格式对象"""
        if style is None:
            return
        font = cf.Font
        interior = cf.Interior
        if style.font_bold is not None:
            font.Bold = style.font_bold
        if style.font_italic is not None:
            font.Italic = style.font_italic
        if style.font_color is not None:
            font.Color = style.font_color
        if style.fill_color is not None:
            interior.Color = style.fill_color