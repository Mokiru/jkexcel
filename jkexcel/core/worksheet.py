import os
from typing import List, Optional, Dict, Any, Union

import pythoncom

from jkexcel.core.range import Range
from jkexcel.models.config import SheetVisibility, RangeStyle
from jkexcel.models.enums import XlFindLookIn, XlLookAt, XlSearchOrder, XlSearchDirection, XlDirection, \
    PageOrientation, PaperSize, XlChartType
from jkexcel.models.exceptions import WorksheetNotFoundError, RangeError
from jkexcel.utils.helpers import cell_address_to_indices
from jkexcel.utils.validators import validate_file_exists


class Worksheet:
    """Excel 工作表封装类"""

    def __init__(self, com_worksheet, excel_app):
        """
        初始化 Worksheet

        Args:
            com_worksheet: COM Worksheet 对象
        """
        if com_worksheet is None:
            raise WorksheetNotFoundError("COM Worksheet 对象不能为 None")
        self._worksheet = com_worksheet
        self._excel = excel_app
        self._data_ops = None
        self._format_ops = None

    def __repr__(self) -> str:
        return f"<Worksheet '{self.name}'>"

    @property
    def com_object(self):
        """获取底层 COM 对象"""
        return self._worksheet

    @property
    def name(self) -> str:
        """获取或设置工作表名称"""
        try:
            return self._worksheet.Name
        except Exception as e:
            raise WorksheetNotFoundError(f"获取名称失败: {e}")

    @name.setter
    def name(self, name: str):
        """设置工作表名称"""
        try:
            self._worksheet.Name = name
        except Exception as e:
            raise WorksheetNotFoundError(f"设置名称失败: {e}")

    @property
    def index(self) -> int:
        """获取工作表索引"""
        try:
            return self._worksheet.Index
        except Exception as e:
            raise WorksheetNotFoundError(f"获取索引失败: {e}")

    @property
    def visible(self) -> bool:
        """获取或设置可见性"""
        try:
            return self._worksheet.Visible == SheetVisibility.VISIBLE.value
        except Exception as e:
            raise WorksheetNotFoundError(f"获取可见性失败: {e}")

    @visible.setter
    def visible(self, visible: bool):
        """设置可见性"""
        try:
            visibility = SheetVisibility.VISIBLE.value if visible else SheetVisibility.HIDDEN.value
            self._worksheet.Visible = visibility
        except Exception as e:
            raise WorksheetNotFoundError(f"设置可见性失败: {e}")

    def set_visibility(self, visibility: SheetVisibility):
        """
        设置可见性

        Args:
            visibility: 可见性枚举
        """
        try:
            self._worksheet.Visible = visibility.value
        except Exception as e:
            raise WorksheetNotFoundError(f"设置可见性失败: {e}")

    @property
    def used_range(self) -> Range:
        """获取使用的范围"""
        try:
            return Range(self._worksheet.UsedRange)
        except Exception as e:
            raise WorksheetNotFoundError(f"获取使用范围失败: {e}")

    @property
    def cells(self) -> Range:
        """获取所有单元格"""
        return Range(self._worksheet.Cells)

    def get_range(self, address: str) -> Range:
        """
        通过地址获取范围

        Args:
            address: 范围地址，如 "A1", "A1:B10"

        Returns:
            Range 对象
        """
        try:
            return Range(self._worksheet.Range(address))
        except Exception as e:
            raise RangeError(f"获取范围失败: {e}")

    def get_cell(self, row: int, column: int) -> Range:
        """
        通过行列获取单元格

        Args:
            row: 行号（从1开始）
            column: 列号（从1开始）

        Returns:
            Range 对象
        """
        try:
            return Range(self._worksheet.Cells(row, column))
        except Exception as e:
            raise RangeError(f"获取单元格失败: {e}")

    def get_cells(self, start_row: int, start_col: int,
                  end_row: int = None, end_col: int = None) -> Range:
        """
        获取单元格范围

        Args:
            start_row: 起始行
            start_col: 起始列
            end_row: 结束行（None 表示只取一行）
            end_col: 结束列（None 表示只取一列）

        Returns:
            Range 对象
        """
        try:
            if end_row is None:
                end_row = start_row
            if end_col is None:
                end_col = start_col

            return Range(self._worksheet.Range(
                self._worksheet.Cells(start_row, start_col),
                self._worksheet.Cells(end_row, end_col)
            ))
        except Exception as e:
            raise RangeError(f"获取单元格范围失败: {e}")

    def activate(self):
        """激活工作表"""
        try:
            self._worksheet.Activate()
        except Exception as e:
            raise WorksheetNotFoundError(f"激活工作表失败: {e}")

    def select(self):
        """选择工作表"""
        try:
            self._worksheet.Select()
        except Exception as e:
            raise WorksheetNotFoundError(f"选择工作表失败: {e}")

    def delete(self):
        """删除工作表"""
        try:
            self._worksheet.Delete()
        except Exception as e:
            raise WorksheetNotFoundError(f"删除工作表失败: {e}")

    def copy(self, before: Optional['Worksheet'] = None,
             after: Optional['Worksheet'] = None) -> None:
        """
        复制工作表

        Args:
            before: 复制到指定工作表之前
            after: 复制到指定工作表之后
        """
        try:
            # 注意：动态绑定下 Copy 的关键字参数会丢失，必须用位置参数 + None 占位
            if before:
                self._worksheet.Copy(before.com_object, None)
            elif after:
                self._worksheet.Copy(None, after.com_object)
            else:
                self._worksheet.Copy()
        except Exception as e:
            raise WorksheetNotFoundError(f"复制工作表失败: {e}")

    def move(self, before: Optional['Worksheet'] = None,
             after: Optional['Worksheet'] = None):
        """
        移动工作表

        Args:
            before: 移动到指定工作表之前
            after: 移动到指定工作表之后
        """
        try:
            # 与 Copy 同理：动态绑定下关键字参数会丢失，使用位置参数 + None 占位
            if before:
                self._worksheet.Move(before.com_object, None)
            elif after:
                self._worksheet.Move(None, after.com_object)
            else:
                self._worksheet.Move()
        except Exception as e:
            raise WorksheetNotFoundError(f"移动工作表失败: {e}")

    def protect(self, password: str = "",
                allow_formatting_cells: bool = True,
                allow_formatting_columns: bool = True,
                allow_formatting_rows: bool = True,
                allow_inserting_columns: bool = True,
                allow_inserting_rows: bool = True,
                allow_inserting_hyperlinks: bool = True,
                allow_deleting_columns: bool = True,
                allow_deleting_rows: bool = True,
                allow_sorting: bool = True,
                allow_filtering: bool = True,
                allow_using_pivot_tables: bool = True):
        """
        保护工作表

        Args:
            password: 密码
            allow_xxx: 允许的操作
        """
        try:
            self._worksheet.Protect(
                Password=password,
                DrawingObjects=True,
                Contents=True,
                Scenarios=True,
                AllowFormattingCells=allow_formatting_cells,
                AllowFormattingColumns=allow_formatting_columns,
                AllowFormattingRows=allow_formatting_rows,
                AllowInsertingColumns=allow_inserting_columns,
                AllowInsertingRows=allow_inserting_rows,
                AllowInsertingHyperlinks=allow_inserting_hyperlinks,
                AllowDeletingColumns=allow_deleting_columns,
                AllowDeletingRows=allow_deleting_rows,
                AllowSorting=allow_sorting,
                AllowFiltering=allow_filtering,
                AllowUsingPivotTables=allow_using_pivot_tables
            )
        except Exception as e:
            raise WorksheetNotFoundError(f"保护工作表失败: {e}")

    def unprotect(self, password: str = ""):
        """
        取消保护

        Args:
            password: 密码
        """
        try:
            self._worksheet.Unprotect(Password=password)
        except Exception as e:
            raise WorksheetNotFoundError(f"取消保护失败: {e}")

    def auto_fit_columns(self, start_col: int = 1, end_col: int = None):
        """
        自动调整列宽

        Args:
            start_col: 起始列
            end_col: 结束列
        """
        try:
            if end_col is None:
                col_range = self._worksheet.Columns(start_col)
            else:
                col_range = self._worksheet.Range(
                    self._worksheet.Cells(1, start_col),
                    self._worksheet.Cells(1, end_col)
                ).EntireColumn

            col_range.AutoFit()
        except Exception as e:
            raise WorksheetNotFoundError(f"自动调整列宽失败: {e}")

    def auto_fit_rows(self, start_row: int = 1, end_row: int = None):
        """
        自动调整行高

        Args:
            start_row: 起始行
            end_row: 结束行
        """
        try:
            if end_row is None:
                row_range = self._worksheet.Rows(start_row)
            else:
                row_range = self._worksheet.Range(
                    self._worksheet.Cells(start_row, 1),
                    self._worksheet.Cells(end_row, 1)
                ).EntireRow

            row_range.AutoFit()
        except Exception as e:
            raise WorksheetNotFoundError(f"自动调整行高失败: {e}")

    def insert_row(self, row: int, count: int = 1):
        """
        插入行

        Args:
            row: 插入位置
            count: 插入数量
        """
        try:
            self.get_cells(row, 1, row + count - 1, 1).com_object.EntireRow.Insert()
        except Exception as e:
            raise WorksheetNotFoundError(f"插入行失败: {e}")

    def insert_column(self, column: int, count: int = 1):
        """
        插入列

        Args:
            column: 插入位置
            count: 插入数量
        """
        try:
            col_letter = self._column_to_letter(column)
            end_col_letter = self._column_to_letter(column + count - 1)
            self._worksheet.Range(f"{col_letter}:{end_col_letter}").Insert()
        except Exception as e:
            raise WorksheetNotFoundError(f"插入列失败: {e}")

    def delete_row(self, row: int, count: int = 1):
        """
        删除行

        Args:
            row: 删除起始行
            count: 删除数量
        """
        try:
            self.get_cells(row, 1, row + count - 1, 1).com_object.EntireRow.Delete()
        except Exception as e:
            raise WorksheetNotFoundError(f"删除行失败: {e}")

    def delete_column(self, column: int, count: int = 1):
        """
        删除列

        Args:
            column: 删除起始列
            count: 删除数量
        """
        try:
            col_letter = self._column_to_letter(column)
            end_col_letter = self._column_to_letter(column + count - 1)
            self._worksheet.Range(f"{col_letter}:{end_col_letter}").Delete()
        except Exception as e:
            raise WorksheetNotFoundError(f"删除列失败: {e}")

    def write_data(self, data: List[List[Any]], start_cell: str = "A1",
                   headers: List[str] = None, style: RangeStyle = None):
        """
        写入数据

        Args:
            data: 二维数据列表
            start_cell: 起始单元格
            headers: 表头列表
            style: 样式配置
        """
        try:
            # 写入表头
            if headers:
                header_range = self.get_range(start_cell).resize(
                    rows=1,
                    cols=len(headers)
                )
                header_range.value = [headers]
                if style:
                    header_range.apply_style(style)

                # 计算数据起始位置
                start_row = self.get_range(start_cell).row + 1
                start_col = self.get_range(start_cell).column
                data_range = self.get_cells(start_row, start_col)
            else:
                data_range = self.get_range(start_cell)

            # 写入数据
            if data:
                data_range.resize(rows=len(data), cols=len(data[0])).value = data

        except Exception as e:
            raise WorksheetNotFoundError(f"写入数据失败: {e}")

    def read_data(self, start_cell: str = "A1",
                  has_headers: bool = False) -> Dict[str, Any]:
        """
        读取数据

        Args:
            start_cell: 起始单元格
            has_headers: 是否有表头

        Returns:
            字典，包含 headers 和 data
        """
        try:
            # 从起始单元格扩展到整个连续数据区域，而不是只读单个单元格
            data_range = self.get_range(start_cell).current_region
            values = data_range.get_values()

            if has_headers and values:
                headers = values[0]
                data = values[1:]
                return {"headers": headers, "data": data}
            else:
                return {"headers": [], "data": values}

        except Exception as e:
            raise WorksheetNotFoundError(f"读取数据失败: {e}")

    def add_table(self, range_address: str, name: str = None,
                  has_headers: bool = True):
        """
        添加表格

        Args:
            range_address: 范围地址
            name: 表格名称
            has_headers: 是否有表头
        """
        try:
            table_range = self.get_range(range_address)

            # 创建表
            table = self._worksheet.ListObjects.Add(
                SourceType=1,  # xlSrcRange
                Source=table_range.com_object,
                XlListObjectHasHeaders=1 if has_headers else 2
            )

            if name:
                table.Name = name

            return table
        except Exception as e:
            raise WorksheetNotFoundError(f"添加表格失败: {e}")

    def _column_to_letter(self, column: int) -> str:
        """列号转字母"""
        result = ""
        while column > 0:
            column, remainder = divmod(column - 1, 26)
            result = chr(65 + remainder) + result
        return result

    def _letter_to_column(self, letter: str) -> int:
        """字母转列号"""
        result = 0
        for char in letter.upper():
            result = result * 26 + (ord(char) - 64)
        return result

    @property
    def auto_filter_mode(self) -> bool:
        """获取或设置自动筛选下拉箭头状态（设为 False 可移除筛选）"""
        try:
            return self._worksheet.AutoFilterMode
        except Exception as e:
            raise WorksheetNotFoundError(f"获取筛选状态失败: {e}")

    @auto_filter_mode.setter
    def auto_filter_mode(self, value: bool):
        try:
            self._worksheet.AutoFilterMode = value
        except Exception as e:
            raise WorksheetNotFoundError(f"设置筛选状态失败: {e}")

    def show_all_data(self):
        """显示所有数据（清除筛选结果，保留筛选箭头）"""
        try:
            if self._worksheet.FilterMode:
                self._worksheet.ShowAllData()
        except Exception as e:
            raise WorksheetNotFoundError(f"显示所有数据失败: {e}")

    def hide_rows(self, row: int, count: int = 1, hidden: bool = True):
        """
        隐藏/显示行

        Args:
            row: 起始行
            count: 行数
            hidden: True 隐藏，False 显示
        """
        try:
            self.get_cells(row, 1, row + count - 1, 1).entire_row.hidden = hidden
        except Exception as e:
            raise WorksheetNotFoundError(f"设置行隐藏失败: {e}")

    def hide_columns(self, column: int, count: int = 1, hidden: bool = True):
        """
        隐藏/显示列

        Args:
            column: 起始列
            count: 列数
            hidden: True 隐藏，False 显示
        """
        try:
            self.get_cells(1, column, 1, column + count - 1).entire_column.hidden = hidden
        except Exception as e:
            raise WorksheetNotFoundError(f"设置列隐藏失败: {e}")

    @property
    def tab_color(self) -> Optional[int]:
        """获取或设置工作表标签颜色（None 表示无颜色）"""
        try:
            color = self._worksheet.Tab.Color
            return color if color is not False else None
        except Exception as e:
            raise WorksheetNotFoundError(f"获取标签颜色失败: {e}")

    @tab_color.setter
    def tab_color(self, color: Optional[int]):
        try:
            if color is None:
                self._worksheet.Tab.ColorIndex = -4142  # xlColorIndexNone
            else:
                self._worksheet.Tab.Color = color
        except Exception as e:
            raise WorksheetNotFoundError(f"设置标签颜色失败: {e}")

    def last_row(self, column: int = 1) -> int:
        """
        获取指定列的最后一个非空行号（等价 Ctrl+↑）

        Args:
            column: 参考列号（从1开始）
        """
        try:
            com = self._worksheet
            return com.Cells(com.Rows.Count, column).End(XlDirection.xlUp.value).Row
        except Exception as e:
            raise WorksheetNotFoundError(f"获取最后一行失败: {e}")

    def last_column(self, row: int = 1) -> int:
        """
        获取指定行的最后一个非空列号（等价 Ctrl+←）

        Args:
            row: 参考行号（从1开始）
        """
        try:
            com = self._worksheet
            return com.Cells(row, com.Columns.Count).End(XlDirection.xlToLeft.value).Column
        except Exception as e:
            raise WorksheetNotFoundError(f"获取最后一列失败: {e}")

    def clear(self, clear_contents: bool = True, clear_format: bool = True):
        """清除整个工作表"""
        self.cells.clear(clear_contents=clear_contents, clear_format=clear_format)

    def freeze_panes(self, cell: str = "A2"):
        """
        冻结窗格（冻结指定单元格上方的行和左侧的列）

        Args:
            cell: 冻结基准单元格，如 "A2" 冻结首行，"B2" 冻结首行首列

        Note:
            窗口操作要求应用可见，后台模式下会临时显示窗口并在完成后恢复
        """
        try:
            row, col = cell_address_to_indices(cell)
            if row <= 1 and col <= 1:
                raise ValueError("冻结基准单元格不能为 A1")
            app_com = self._excel.com_object
            old_visible = app_com.Visible
            try:
                app_com.Visible = True
                self.activate()
                window = app_com.ActiveWindow
                window.FreezePanes = False
                window.SplitRow = row - 1
                window.SplitColumn = col - 1
                window.FreezePanes = True
            finally:
                app_com.Visible = old_visible
        except Exception as e:
            raise WorksheetNotFoundError(f"冻结窗格失败: {e}")

    def unfreeze_panes(self):
        """取消冻结窗格"""
        try:
            app_com = self._excel.com_object
            old_visible = app_com.Visible
            try:
                app_com.Visible = True
                self.activate()
                window = app_com.ActiveWindow
                window.FreezePanes = False
                window.SplitRow = 0
                window.SplitColumn = 0
            finally:
                app_com.Visible = old_visible
        except Exception as e:
            raise WorksheetNotFoundError(f"取消冻结窗格失败: {e}")

    @property
    def shapes(self):
        """获取图形对象集合（图片/形状等，COM 对象）"""
        return self._worksheet.Shapes

    def add_picture(self, file_path: str,
                    left: float = 0, top: float = 0,
                    width: float = None, height: float = None,
                    cell: str = None):
        """
        插入图片

        Args:
            file_path: 图片文件路径
            left: 左边距（磅）
            top: 上边距（磅）
            width: 宽度（磅），None 保持原始尺寸
            height: 高度（磅），None 保持原始尺寸
            cell: 指定后以该单元格左上角为插入位置（覆盖 left/top）

        Returns:
            COM Shape 对象
        """
        validate_file_exists(file_path)
        try:
            if cell:
                anchor = self.get_range(cell).com_object
                left, top = anchor.Left, anchor.Top
            # 位置参数：Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height
            return self._worksheet.Shapes.AddPicture(
                os.path.abspath(file_path), 0, -1,
                left, top,
                -1 if width is None else width,
                -1 if height is None else height
            )
        except Exception as e:
            raise WorksheetNotFoundError(f"插入图片失败: {e}")

    def add_comment(self, cell: str, text: str):
        """给单元格添加批注"""
        self.get_range(cell).add_comment(text)

    def add_chart(self, chart_type: XlChartType = XlChartType.xlColumnClustered,
                  source_range: str = None,
                  left: float = 100, top: float = 100,
                  width: float = 400, height: float = 300):
        """
        添加嵌入式图表

        Args:
            chart_type: 图表类型
            source_range: 数据源范围地址（含表头），None 表示稍后设置
            left/top: 图表位置（磅）
            width/height: 图表大小（磅）

        Returns:
            Chart 对象
        """
        from jkexcel.core.chart import Chart
        try:
            # 位置参数：Left, Top, Width, Height
            chart_obj = self._worksheet.ChartObjects().Add(left, top, width, height)
            com_chart = chart_obj.Chart
            com_chart.ChartType = chart_type.value
            if source_range:
                com_chart.SetSourceData(self.get_range(source_range).com_object, 2)  # 2=xlColumns
            return Chart(com_chart, chart_obj)
        except Exception as e:
            raise WorksheetNotFoundError(f"添加图表失败: {e}")

    def create_pivot_table(self, source_range: str, dest_cell: str, name: str,
                           row_fields: List[str] = None,
                           column_fields: List[str] = None,
                           data_fields: List[str] = None):
        """
        创建数据透视表

        Args:
            source_range: 源数据范围地址（首行必须为表头）
            dest_cell: 透视表放置的起始单元格
            name: 透视表名称
            row_fields: 行字段列名列表
            column_fields: 列字段列名列表
            data_fields: 值字段列名列表（默认求和，更多汇总方式用 PivotTable.add_data_field）

        Returns:
            PivotTable 对象
        """
        from jkexcel.core.pivot import PivotTable
        try:
            src = self.get_range(source_range)
            r, c = src.row, src.column
            rows, cols = src.shape
            # R1C1 引用样式（xlR1C1=-4150），避免动态绑定下 Address 参数化调用问题
            source_ref = f"'{self.name}'!R{r}C{c}:R{r + rows - 1}C{c + cols - 1}"
            # 位置参数：SourceType(1=xlDatabase), SourceData
            cache = self._worksheet.Parent.PivotCaches().Create(1, source_ref)
            dest = self.get_range(dest_cell)
            # 位置参数：TableDestination, TableName
            com_pt = cache.CreatePivotTable(dest.com_object, name)
            pivot = PivotTable(com_pt, cache)
            for field in row_fields or []:
                pivot.add_row_field(field)
            for field in column_fields or []:
                pivot.add_column_field(field)
            for field in data_fields or []:
                pivot.add_data_field(field)
            return pivot
        except Exception as e:
            raise WorksheetNotFoundError(f"创建数据透视表失败: {e}")

    def add_hyperlink(self, cell: str, address: str,
                      text_to_display: str = None,
                      screen_tip: str = None,
                      sub_address: str = None):
        """给单元格添加超链接"""
        self.get_range(cell).add_hyperlink(address, text_to_display, screen_tip, sub_address)

    @property
    def print_area(self) -> str:
        """获取或设置打印区域（如 "A1:F20"，设为空字符串清除）"""
        try:
            return self._worksheet.PageSetup.PrintArea
        except Exception as e:
            raise WorksheetNotFoundError(f"获取打印区域失败: {e}")

    @print_area.setter
    def print_area(self, address: str):
        try:
            self._worksheet.PageSetup.PrintArea = address or ""
        except Exception as e:
            raise WorksheetNotFoundError(f"设置打印区域失败: {e}")

    def set_print_titles(self, rows: str = None, columns: str = None):
        """
        设置每页重复打印的标题行/列

        Args:
            rows: 标题行，如 "1:2" 或 "$1:$2"
            columns: 标题列，如 "A:B" 或 "$A:$B"
        """
        try:
            ps = self._worksheet.PageSetup
            if rows is not None:
                ps.PrintTitleRows = "" if not rows else "$" + rows.replace("$", "").replace(":", ":$")
            if columns is not None:
                ps.PrintTitleColumns = "" if not columns else "$" + columns.replace("$", "").replace(":", ":$")
        except Exception as e:
            raise WorksheetNotFoundError(f"设置打印标题失败: {e}")

    def set_page_setup(self,
                       orientation: PageOrientation = None,
                       paper_size: PaperSize = None,
                       zoom: Union[int, bool] = None,
                       fit_to_pages_wide: int = None,
                       fit_to_pages_tall: int = None,
                       left_margin: float = None,
                       right_margin: float = None,
                       top_margin: float = None,
                       bottom_margin: float = None,
                       header_margin: float = None,
                       footer_margin: float = None,
                       center_horizontally: bool = None,
                       center_vertically: bool = None):
        """
        页面设置

        Args:
            orientation: 纸张方向
            paper_size: 纸张大小
            zoom: 缩放百分比（10-400），与 fit_to_pages 互斥
            fit_to_pages_wide: 缩放为指定页宽（需同时设置 zoom=False 或本方法自动处理）
            fit_to_pages_tall: 缩放为指定页高
            left_margin/right_margin/top_margin/bottom_margin: 页边距（厘米）
            header_margin/footer_margin: 页眉/页脚边距（厘米）
            center_horizontally: 水平居中
            center_vertically: 垂直居中
        """
        try:
            ps = self._worksheet.PageSetup
            app = self._excel.com_object
            if orientation is not None:
                ps.Orientation = orientation.value
            if paper_size is not None:
                ps.PaperSize = paper_size.value
            if fit_to_pages_wide is not None or fit_to_pages_tall is not None:
                ps.Zoom = False
                # 0 表示自动（COM 中需传 False，直接传 0 会报错）
                if fit_to_pages_wide is not None:
                    ps.FitToPagesWide = False if fit_to_pages_wide == 0 else fit_to_pages_wide
                if fit_to_pages_tall is not None:
                    ps.FitToPagesTall = False if fit_to_pages_tall == 0 else fit_to_pages_tall
            elif zoom is not None:
                ps.Zoom = zoom
            for attr, cm in (("LeftMargin", left_margin), ("RightMargin", right_margin),
                             ("TopMargin", top_margin), ("BottomMargin", bottom_margin),
                             ("HeaderMargin", header_margin), ("FooterMargin", footer_margin)):
                if cm is not None:
                    setattr(ps, attr, app.CentimetersToPoints(cm))
            if center_horizontally is not None:
                ps.CenterHorizontally = center_horizontally
            if center_vertically is not None:
                ps.CenterVertically = center_vertically
        except Exception as e:
            raise WorksheetNotFoundError(f"页面设置失败: {e}")

    def set_header(self, left: str = None, center: str = None, right: str = None):
        """设置页眉（支持 &P 页码、&D 日期等 Excel 页眉代码）"""
        try:
            ps = self._worksheet.PageSetup
            if left is not None:
                ps.LeftHeader = left
            if center is not None:
                ps.CenterHeader = center
            if right is not None:
                ps.RightHeader = right
        except Exception as e:
            raise WorksheetNotFoundError(f"设置页眉失败: {e}")

    def set_footer(self, left: str = None, center: str = None, right: str = None):
        """设置页脚（支持 &P 页码、&N 总页数等 Excel 页脚代码）"""
        try:
            ps = self._worksheet.PageSetup
            if left is not None:
                ps.LeftFooter = left
            if center is not None:
                ps.CenterFooter = center
            if right is not None:
                ps.RightFooter = right
        except Exception as e:
            raise WorksheetNotFoundError(f"设置页脚失败: {e}")

    def print_out(self, copies: int = 1,
                  preview: bool = False,
                  active_printer: str = None,
                  print_to_file: bool = False,
                  collate: bool = True):
        """
        打印工作表

        Args:
            copies: 份数
            preview: 是否预览
            active_printer: 打印机
            print_to_file: 打印到文件
            collate: 是否逐份打印
        """
        try:
            if preview:
                self._worksheet.PrintPreview()
            else:
                # 位置参数：From, To, Copies, Preview, ActivePrinter, PrintToFile, Collate
                self._worksheet.PrintOut(
                    pythoncom.Missing, pythoncom.Missing, copies, False,
                    pythoncom.Missing if active_printer is None else active_printer,
                    print_to_file, collate
                )
        except Exception as e:
            raise WorksheetNotFoundError(f"打印工作表失败: {e}")

    def export_as_pdf(self, file_path: str,
                      quality: int = 0,  # xlQualityStandard
                      include_doc_props: bool = True,
                      ignore_print_areas: bool = False,
                      open_after_publish: bool = False):
        """
        将工作表导出为 PDF

        Args:
            file_path: 保存路径
            quality: 质量
            include_doc_props: 包含文档属性
            ignore_print_areas: 忽略打印区域
            open_after_publish: 导出后打开文件
        """
        try:
            # 位置参数：Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas, From, To, OpenAfterPublish
            self._worksheet.ExportAsFixedFormat(
                0, os.path.abspath(file_path), quality,
                include_doc_props, ignore_print_areas,
                pythoncom.Missing, pythoncom.Missing, open_after_publish
            )
        except Exception as e:
            raise WorksheetNotFoundError(f"导出 PDF 失败: {e}")

    # Range 的便捷方法代理
    def current_region(self, start_cell: str = "A1") -> Range:
        """获取当前区域"""
        return self.get_range(start_cell).current_region

    @property
    def data(self):
        """数据操作入口（DataOperations）"""
        if self._data_ops is None:
            from jkexcel.operations.data_operations import DataOperations
            self._data_ops = DataOperations(self)
        return self._data_ops

    @property
    def format(self):
        """格式操作入口（FormatOperations）"""
        if self._format_ops is None:
            from jkexcel.operations.format_operations import FormatOperations
            self._format_ops = FormatOperations(self)
        return self._format_ops

    def find_all(self, what: str, look_in: int = -4163) -> List[Range]:
        """查找所有匹配项"""
        return self.used_range.find_all(what, look_in)

    def find(self, what: str, after: Range = None, look_in: 'XlFindLookIn' = XlFindLookIn.xlValues,
             look_at: 'XlLookAt' = XlLookAt.xlWhole,
             search_order: 'XlSearchOrder' = XlSearchOrder.xlByRows,
             search_direction: 'XlSearchDirection' = XlSearchDirection.xlNext, match_case: bool = False,
             match_byte: bool = False, search_format: bool = False) -> Range:
        return self.cells.find(what, after, look_in, look_at, search_order, search_direction, match_case, match_byte,
                               search_format)
