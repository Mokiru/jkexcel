import os
from typing import Optional

from jkexcel.models.enums import XlChartType
from jkexcel.models.exceptions import ExcelCOMError


class Chart:
    """Excel 图表封装类"""

    def __init__(self, com_chart, com_chart_object=None):
        """
        初始化 Chart

        Args:
            com_chart: COM Chart 对象
            com_chart_object: COM ChartObject 容器（嵌入式图表才有）
        """
        if com_chart is None:
            raise ExcelCOMError("COM Chart 对象不能为 None")
        self._chart = com_chart
        self._chart_object = com_chart_object

    def __repr__(self) -> str:
        return f"<Chart type={self.chart_type}>"

    @property
    def com_object(self):
        """获取底层 COM Chart 对象"""
        return self._chart

    @property
    def com_container(self):
        """获取底层 COM ChartObject 容器"""
        return self._chart_object

    @property
    def chart_type(self) -> XlChartType:
        """获取或设置图表类型"""
        try:
            return XlChartType(self._chart.ChartType)
        except ValueError:
            return self._chart.ChartType
        except Exception as e:
            raise ExcelCOMError(f"获取图表类型失败: {e}")

    @chart_type.setter
    def chart_type(self, value: XlChartType):
        try:
            self._chart.ChartType = value.value if isinstance(value, XlChartType) else value
        except Exception as e:
            raise ExcelCOMError(f"设置图表类型失败: {e}")

    def set_source_data(self, source_range, plot_by_rows: bool = False):
        """
        设置图表数据源

        Args:
            source_range: Range 对象（含表头/系列名）
            plot_by_rows: True 按行绘制系列，False 按列（1=xlRows, 2=xlColumns）
        """
        try:
            # 位置参数：Source, PlotBy
            self._chart.SetSourceData(source_range.com_object, 1 if plot_by_rows else 2)
        except Exception as e:
            raise ExcelCOMError(f"设置数据源失败: {e}")

    @property
    def title(self) -> Optional[str]:
        """获取或设置图表标题（设置时自动启用标题）"""
        try:
            if self._chart.HasTitle:
                return self._chart.ChartTitle.Text
            return None
        except Exception as e:
            raise ExcelCOMError(f"获取图表标题失败: {e}")

    @title.setter
    def title(self, text: str):
        try:
            self._chart.HasTitle = True
            self._chart.ChartTitle.Text = text
        except Exception as e:
            raise ExcelCOMError(f"设置图表标题失败: {e}")

    @property
    def has_legend(self) -> bool:
        """获取或设置是否显示图例"""
        try:
            return self._chart.HasLegend
        except Exception as e:
            raise ExcelCOMError(f"获取图例状态失败: {e}")

    @has_legend.setter
    def has_legend(self, value: bool):
        try:
            self._chart.HasLegend = value
        except Exception as e:
            raise ExcelCOMError(f"设置图例状态失败: {e}")

    def set_position(self, left: float, top: float,
                     width: float = None, height: float = None):
        """
        设置图表位置和大小（仅嵌入式图表）

        Args:
            left: 左边距（磅）
            top: 上边距（磅）
            width: 宽度（磅）
            height: 高度（磅）
        """
        if self._chart_object is None:
            raise ExcelCOMError("独立图表工作表不支持设置位置")
        try:
            self._chart_object.Left = left
            self._chart_object.Top = top
            if width is not None:
                self._chart_object.Width = width
            if height is not None:
                self._chart_object.Height = height
        except Exception as e:
            raise ExcelCOMError(f"设置图表位置失败: {e}")

    def export_image(self, file_path: str, filter_name: str = "PNG") -> bool:
        """
        导出图表为图片

        Args:
            file_path: 保存路径
            filter_name: 图片格式（PNG/JPG/GIF）
        """
        try:
            # 位置参数：Filename, FilterName
            result = self._chart.Export(file_path, filter_name)
            # WPS 的 Export 返回 False 但实际已写入文件，以文件是否生成为准
            return bool(result) or os.path.exists(file_path)
        except Exception as e:
            raise ExcelCOMError(f"导出图表图片失败: {e}")

    def delete(self):
        """删除图表"""
        try:
            if self._chart_object is not None:
                self._chart_object.Delete()
            else:
                self._chart.Delete()
        except Exception as e:
            raise ExcelCOMError(f"删除图表失败: {e}")