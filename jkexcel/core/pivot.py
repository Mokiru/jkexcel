from typing import Optional

from jkexcel.core.range import Range
from jkexcel.models.enums import XlPivotFieldOrientation, XlConsolidationFunction
from jkexcel.models.exceptions import ExcelCOMError


class PivotTable:
    """Excel 数据透视表封装类"""

    def __init__(self, com_pivot_table, com_pivot_cache=None):
        """
        初始化 PivotTable

        Args:
            com_pivot_table: COM PivotTable 对象
            com_pivot_cache: COM PivotCache 对象
        """
        if com_pivot_table is None:
            raise ExcelCOMError("COM PivotTable 对象不能为 None")
        self._pivot = com_pivot_table
        self._cache = com_pivot_cache

    def __repr__(self) -> str:
        return f"<PivotTable '{self.name}'>"

    @property
    def com_object(self):
        """获取底层 COM PivotTable 对象"""
        return self._pivot

    @property
    def com_cache(self):
        """获取底层 COM PivotCache 对象"""
        return self._cache

    @property
    def name(self) -> str:
        """获取透视表名称"""
        try:
            return self._pivot.Name
        except Exception as e:
            raise ExcelCOMError(f"获取透视表名称失败: {e}")

    @property
    def table_range(self) -> Range:
        """获取透视表占用的范围（含筛选字段区域）"""
        try:
            return Range(self._pivot.TableRange2)
        except Exception as e:
            raise ExcelCOMError(f"获取透视表范围失败: {e}")

    def add_row_field(self, field_name: str, position: int = None):
        """
        添加行字段

        Args:
            field_name: 源数据列名
            position: 字段顺序（从1开始）
        """
        self._set_field(field_name, XlPivotFieldOrientation.xlRowField, position)

    def add_column_field(self, field_name: str, position: int = None):
        """添加列字段"""
        self._set_field(field_name, XlPivotFieldOrientation.xlColumnField, position)

    def add_page_field(self, field_name: str, position: int = None):
        """添加筛选（页）字段"""
        self._set_field(field_name, XlPivotFieldOrientation.xlPageField, position)

    def _set_field(self, field_name: str, orientation: XlPivotFieldOrientation,
                   position: int = None):
        """设置字段方向"""
        try:
            field = self._pivot.PivotFields(field_name)
            field.Orientation = orientation.value
            if position is not None:
                field.Position = position
        except Exception as e:
            raise ExcelCOMError(f"设置透视字段失败: {e}")

    def add_data_field(self, field_name: str,
                       caption: str = None,
                       function: XlConsolidationFunction = XlConsolidationFunction.xlSum):
        """
        添加值字段

        Args:
            field_name: 源数据列名
            caption: 显示名称（如 "求和项:金额"）
            function: 汇总方式（求和/计数/平均等）
        """
        try:
            field = self._pivot.PivotFields(field_name)
            # 位置参数：Field, Caption, Function
            return self._pivot.AddDataField(
                field,
                caption if caption is not None else f"求和项:{field_name}",
                function.value
            )
        except Exception as e:
            raise ExcelCOMError(f"添加值字段失败: {e}")

    def refresh(self):
        """刷新透视表数据"""
        try:
            self._pivot.RefreshTable()
        except Exception as e:
            raise ExcelCOMError(f"刷新透视表失败: {e}")

    def delete(self):
        """删除透视表（清除其占用区域）"""
        try:
            self._pivot.TableRange2.Clear()
        except Exception as e:
            raise ExcelCOMError(f"删除透视表失败: {e}")