from typing import List, Dict, Any, Optional, Union
import pandas as pd

from jkexcel.core.range import Range
from jkexcel.core.worksheet import Worksheet
from jkexcel.utils.validators import validate_range_address


class DataOperations:
    """数据操作类"""

    def __init__(self, worksheet: Worksheet):
        """
        初始化

        Args:
            worksheet: Worksheet 对象
        """
        self.ws = worksheet

    def read_to_dataframe(self,
                          start_cell: str = "A1",
                          has_headers: bool = True) -> pd.DataFrame:
        """
        读取为 pandas DataFrame

        Args:
            start_cell: 起始单元格
            has_headers: 是否有表头

        Returns:
            DataFrame
        """
        validate_range_address(start_cell)

        data = self.ws.read_data(start_cell=start_cell, has_headers=has_headers)

        if has_headers and data['headers']:
            return pd.DataFrame(data['data'], columns=data['headers'])
        else:
            return pd.DataFrame(data['data'])

    def write_from_dataframe(self,
                             df: pd.DataFrame,
                             start_cell: str = "A1",
                             include_index: bool = False,
                             include_headers: bool = True):
        """
        从 DataFrame 写入

        Args:
            df: DataFrame
            start_cell: 起始单元格
            include_index: 是否包含索引
            include_headers: 是否包含表头
        """
        validate_range_address(start_cell)

        if include_index:
            df = df.reset_index()

        # NaN/NaT 转 None，避免 COM 写入报错；Timestamp 转原生 datetime
        df = df.astype(object).where(pd.notnull(df), None)
        data = [
            [v.to_pydatetime() if isinstance(v, pd.Timestamp) else v for v in row]
            for row in df.values.tolist()
        ]
        headers = df.columns.tolist() if include_headers else None

        self.ws.write_data(
            data=data,
            start_cell=start_cell,
            headers=headers
        )

    def filter_data(self,
                    criteria_range: str,
                    criteria: Dict[str, Any]) -> List[Range]:
        """
        筛选数据（内存中按条件匹配，不改变工作表显示）

        Args:
            criteria_range: 条件范围（首行为表头）
            criteria: 条件字典 {列: 值}，键可以是表头名称或列号（从1开始，相对范围）

        Returns:
            匹配的数据行范围列表（不含表头行）
        """
        validate_range_address(criteria_range)

        data_range = self.ws.get_range(criteria_range)
        values = data_range.get_values()
        if len(values) < 2:
            return []

        headers = values[0]
        # 将条件键统一解析为范围内的列下标（0基）
        col_indices = {}
        for key, expected in criteria.items():
            if isinstance(key, int):
                idx = key - 1
            else:
                idx = headers.index(key) if key in headers else -1
            if not 0 <= idx < len(headers):
                raise ValueError(f"筛选条件列不存在: {key}")
            col_indices[idx] = expected

        start_row = data_range.row
        start_col = data_range.column
        col_count = len(headers)

        result = []
        for offset, row_values in enumerate(values[1:], start=1):
            if all(row_values[idx] == expected for idx, expected in col_indices.items()):
                abs_row = start_row + offset
                result.append(self.ws.get_cells(abs_row, start_col, abs_row, start_col + col_count - 1))
        return result

    def find_duplicates(self,
                        search_range: str,
                        columns: Optional[List[int]] = None) -> List[Range]:
        """
        查找重复值（按行比较，返回第二次及之后出现的重复行）

        Args:
            search_range: 搜索范围（所有行均视为数据，不含表头假设）
            columns: 要检查的列号列表（从1开始，相对范围），None 表示比较所有列

        Returns:
            重复数据行的范围列表
        """
        validate_range_address(search_range)

        data_range = self.ws.get_range(search_range)
        values = data_range.get_values()
        if len(values) < 2:
            return []

        col_count = len(values[0])
        indices = [c - 1 for c in columns] if columns else list(range(col_count))

        start_row = data_range.row
        start_col = data_range.column

        seen = set()
        result = []
        for offset, row_values in enumerate(values):
            key = tuple(row_values[i] for i in indices)
            if key in seen:
                abs_row = start_row + offset
                result.append(self.ws.get_cells(abs_row, start_col, abs_row, start_col + col_count - 1))
            else:
                seen.add(key)
        return result

    def remove_duplicates(self,
                          search_range: str,
                          columns: Optional[List[int]] = None,
                          has_headers: bool = True):
        """
        删除重复值（调用 Excel 原生 RemoveDuplicates，直接修改工作表）

        Args:
            search_range: 搜索范围
            columns: 要检查的列号列表（从1开始，相对范围），None 表示所有列
            has_headers: 范围首行是否为表头
        """
        validate_range_address(search_range)

        data_range = self.ws.get_range(search_range)
        col_count = data_range.columns_count
        cols = tuple(columns) if columns else tuple(range(1, col_count + 1))
        # 位置参数：Columns, Header（1=xlYes, 2=xlNo）
        data_range.com_object.RemoveDuplicates(cols, 1 if has_headers else 2)