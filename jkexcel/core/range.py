from typing import Any, List, Tuple, Optional, Union, Dict, Iterator
import pythoncom
import win32com.client

from jkexcel.models.config import RangeStyle
from jkexcel.models.enums import XlFindLookIn, XlLookAt, XlSearchDirection, XlSearchOrder, XlPasteType, \
    XlPasteSpecialOperation, XlDirection, XlInsertShiftDirection, XlDeleteShiftDirection, XlCellType, \
    XlAutoFilterOperator, XlBordersIndex, BorderLineStyle, BorderWeight
from jkexcel.models.exceptions import RangeError
from jkexcel.utils.helpers import rgb_to_excel_color


class Range:
    """Excel 范围封装类"""

    def __init__(self, com_range):
        """
        初始化 Range

        Args:
            com_range: COM Range 对象
        """
        if com_range is None:
            raise RangeError("COM Range 对象不能为 None")
        self._range = com_range

    def __repr__(self) -> str:
        return f"<Range '{self.address}'>"

    def __iter__(self) -> Iterator['Range']:
        """迭代范围内的每个单元格（行优先）"""
        for i in range(1, self.count + 1):
            yield Range(self._range.Item(i))

    def __getitem__(self, key: Union[int, Tuple[int, int], str]) -> 'Range':
        """
        获取单元格

        Args:
            key: 整数序号（行优先，从1开始，相对范围）、
                 (行, 列) 元组（从1开始，相对范围）或 A1 地址（工作表绝对地址）
        """
        try:
            if isinstance(key, str):
                return Range(self._range.Worksheet.Range(key))
            if isinstance(key, tuple):
                return Range(self._range.Item(key[0], key[1]))
            return Range(self._range.Item(key))
        except Exception as e:
            raise RangeError(f"获取单元格失败: {e}")

    @property
    def com_object(self):
        """获取底层 COM 对象"""
        return self._range

    @property
    def address(self) -> str:
        """获取范围地址（如 A1:B10）"""
        try:
            return self._range.Address.replace('$', '')
        except Exception as e:
            raise RangeError(f"获取地址失败: {e}")

    @property
    def value(self) -> Any:
        """获取或设置范围的值"""
        try:
            return self._range.Value
        except Exception as e:
            raise RangeError(f"获取值失败: {e}")

    @value.setter
    def value(self, value: Any):
        """设置范围的值"""
        try:
            self._range.Value = self._normalize_values(value)
        except Exception as e:
            raise RangeError(f"设置值失败: {e}")

    @property
    def formula(self) -> str:
        """获取或设置公式"""
        try:
            return self._range.Formula
        except Exception as e:
            raise RangeError(f"获取公式失败: {e}")

    @formula.setter
    def formula(self, formula: str):
        """设置公式"""
        try:
            self._range.Formula = formula
        except Exception as e:
            raise RangeError(f"设置公式失败: {e}")

    @property
    def text(self) -> str:
        """获取文本"""
        try:
            return self._range.Text
        except Exception as e:
            raise RangeError(f"获取文本失败: {e}")

    @property
    def row(self) -> int:
        """获取起始行号"""
        try:
            return self._range.Row
        except Exception as e:
            raise RangeError(f"获取行号失败: {e}")

    @property
    def column(self) -> int:
        """获取起始列号"""
        try:
            return self._range.Column
        except Exception as e:
            raise RangeError(f"获取列号失败: {e}")

    @property
    def row_height(self) -> float:
        """获取或设置行高"""
        try:
            return self._range.RowHeight
        except Exception as e:
            raise RangeError(f"获取行高失败: {e}")

    @row_height.setter
    def row_height(self, height: float):
        """设置行高"""
        try:
            self._range.RowHeight = height
        except Exception as e:
            raise RangeError(f"设置行高失败: {e}")

    @property
    def column_width(self) -> float:
        """获取或设置列宽"""
        try:
            return self._range.ColumnWidth
        except Exception as e:
            raise RangeError(f"获取列宽失败: {e}")

    @column_width.setter
    def column_width(self, width: float):
        """设置列宽"""
        try:
            self._range.ColumnWidth = width
        except Exception as e:
            raise RangeError(f"设置列宽失败: {e}")

    @property
    def current_region(self) -> 'Range':
        """获取当前区域（以空行空列为界的连续数据区域）"""
        try:
            return Range(self._range.CurrentRegion)
        except Exception as e:
            raise RangeError(f"获取当前区域失败: {e}")

    @property
    def count(self) -> int:
        """获取单元格总数"""
        try:
            return self._range.Count
        except Exception as e:
            raise RangeError(f"获取单元格数量失败: {e}")

    @property
    def rows_count(self) -> int:
        """获取行数"""
        try:
            return self._range.Rows.Count
        except Exception as e:
            raise RangeError(f"获取行数失败: {e}")

    @property
    def columns_count(self) -> int:
        """获取列数"""
        try:
            return self._range.Columns.Count
        except Exception as e:
            raise RangeError(f"获取列数失败: {e}")

    @property
    def shape(self) -> Tuple[int, int]:
        """获取范围形状 (行数, 列数)"""
        return self.rows_count, self.columns_count

    @property
    def entire_row(self) -> 'Range':
        """获取范围所在的整行"""
        try:
            return Range(self._range.EntireRow)
        except Exception as e:
            raise RangeError(f"获取整行失败: {e}")

    @property
    def entire_column(self) -> 'Range':
        """获取范围所在的整列"""
        try:
            return Range(self._range.EntireColumn)
        except Exception as e:
            raise RangeError(f"获取整列失败: {e}")

    @property
    def hidden(self) -> bool:
        """获取或设置隐藏状态（仅对整行/整列有效）"""
        try:
            return self._range.Hidden
        except Exception as e:
            raise RangeError(f"获取隐藏状态失败: {e}")

    @hidden.setter
    def hidden(self, value: bool):
        try:
            self._range.Hidden = value
        except Exception as e:
            raise RangeError(f"设置隐藏状态失败: {e}")

    @property
    def merge_cells(self) -> Optional[bool]:
        """获取合并状态（区域内状态不一致时返回 None）"""
        try:
            return self._range.MergeCells
        except Exception as e:
            raise RangeError(f"获取合并状态失败: {e}")

    @property
    def number_format(self) -> str:
        """获取或设置数字格式"""
        try:
            return self._range.NumberFormat
        except Exception as e:
            raise RangeError(f"获取数字格式失败: {e}")

    @number_format.setter
    def number_format(self, value: str):
        try:
            self._range.NumberFormat = value
        except Exception as e:
            raise RangeError(f"设置数字格式失败: {e}")

    @property
    def wrap_text(self) -> Optional[bool]:
        """获取或设置自动换行"""
        try:
            return self._range.WrapText
        except Exception as e:
            raise RangeError(f"获取自动换行失败: {e}")

    @wrap_text.setter
    def wrap_text(self, value: bool):
        try:
            self._range.WrapText = value
        except Exception as e:
            raise RangeError(f"设置自动换行失败: {e}")

    @property
    def locked(self) -> Optional[bool]:
        """获取或设置锁定状态（工作表保护后生效）"""
        try:
            return self._range.Locked
        except Exception as e:
            raise RangeError(f"获取锁定状态失败: {e}")

    @locked.setter
    def locked(self, value: bool):
        try:
            self._range.Locked = value
        except Exception as e:
            raise RangeError(f"设置锁定状态失败: {e}")

    @property
    def font_name(self) -> Optional[str]:
        """获取字体名称"""
        try:
            return self._range.Font.Name
        except Exception as e:
            raise RangeError(f"获取字体名称失败: {e}")

    @property
    def font_size(self) -> Optional[float]:
        """获取字体大小"""
        try:
            return self._range.Font.Size
        except Exception as e:
            raise RangeError(f"获取字体大小失败: {e}")

    @property
    def font_bold(self) -> Optional[bool]:
        """获取字体加粗状态"""
        try:
            return self._range.Font.Bold
        except Exception as e:
            raise RangeError(f"获取字体加粗状态失败: {e}")

    @property
    def font_color(self) -> Optional[int]:
        """获取字体颜色（Excel 颜色值）"""
        try:
            return self._range.Font.Color
        except Exception as e:
            raise RangeError(f"获取字体颜色失败: {e}")

    @property
    def fill_color(self) -> Optional[int]:
        """获取填充颜色（Excel 颜色值）"""
        try:
            return self._range.Interior.Color
        except Exception as e:
            raise RangeError(f"获取填充颜色失败: {e}")

    def get_style(self) -> RangeStyle:
        """读取当前样式快照（区域内格式不一致的字段为 None）"""
        try:
            font = self._range.Font
            return RangeStyle(
                font_name=font.Name,
                font_size=font.Size,
                font_bold=font.Bold,
                font_italic=font.Italic,
                font_color=font.Color,
                font_underline=font.Underline,
                font_strikethrough=font.Strikethrough,
                fill_color=self._range.Interior.Color,
                horizontal_alignment=self._range.HorizontalAlignment,
                vertical_alignment=self._range.VerticalAlignment,
                number_format=self._range.NumberFormat,
                wrap_text=self._range.WrapText,
            )
        except Exception as e:
            raise RangeError(f"读取样式失败: {e}")

    def offset(self, row_offset: int = 0, column_offset: int = 0) -> 'Range':
        """
        获取偏移后的范围（大小不变）

        Args:
            row_offset: 行偏移（正数向下，负数向上）
            column_offset: 列偏移（正数向右，负数向左）
        """
        try:
            # 早期绑定（gen_py）下 Offset 是参数化属性，需通过 GetOffset 调用
            get_offset = getattr(self._range, "GetOffset", None)
            if get_offset is not None:
                return Range(get_offset(row_offset, column_offset))
            return Range(self._range.Offset(row_offset, column_offset))
        except Exception as e:
            raise RangeError(f"偏移范围失败: {e}")

    def end(self, direction: XlDirection) -> 'Range':
        """
        获取区域边缘的单元格（等价 Ctrl+方向键）

        Args:
            direction: 移动方向，如 XlDirection.xlUp 可用于定位最后一行
        """
        try:
            return Range(self._range.End(direction.value))
        except Exception as e:
            raise RangeError(f"获取边缘单元格失败: {e}")

    def insert(self, shift: XlInsertShiftDirection = XlInsertShiftDirection.xlShiftDown):
        """
        插入单元格（原单元格按 shift 方向移动）

        Args:
            shift: 移动方向（下移/右移）
        """
        try:
            self._range.Insert(shift.value)
        except Exception as e:
            raise RangeError(f"插入单元格失败: {e}")

    def delete(self, shift: XlDeleteShiftDirection = XlDeleteShiftDirection.xlShiftUp):
        """
        删除单元格（周围单元格按 shift 方向移动）

        Args:
            shift: 移动方向（上移/左移）
        """
        try:
            self._range.Delete(shift.value)
        except Exception as e:
            raise RangeError(f"删除单元格失败: {e}")

    def special_cells(self, cell_type: XlCellType, value: Any = None) -> Optional['Range']:
        """
        定位特殊单元格

        Args:
            cell_type: 单元格类型（空值/公式/可见单元格等）
            value: 配合 xlCellTypeConstants/xlCellTypeFormulas 使用的值类型筛选

        Returns:
            匹配的 Range，无匹配单元格时返回 None
        """
        try:
            com = self._range.SpecialCells(cell_type.value,
                                           pythoncom.Missing if value is None else value)
            return Range(com)
        except Exception:
            # 无匹配单元格时 COM 报错，视为空结果
            return None

    def auto_filter(self, field: int = None,
                    criteria1: Any = None,
                    operator: XlAutoFilterOperator = XlAutoFilterOperator.xlAnd,
                    criteria2: Any = None,
                    visible_dropdown: bool = True):
        """
        自动筛选

        Args:
            field: 筛选列号（从1开始，相对范围），None 表示仅切换筛选下拉箭头
            criteria1: 筛选条件1（如 ">100"、"苹果"）
            operator: 两个条件的组合方式或 Top10 等筛选类型
            criteria2: 筛选条件2
            visible_dropdown: 是否显示下拉箭头
        """
        try:
            if field is None:
                # 显式全 Missing：gen_py 早期绑定下无参调用会被填入默认 Operator 导致失败
                self._range.AutoFilter(pythoncom.Missing, pythoncom.Missing,
                                       pythoncom.Missing, pythoncom.Missing, pythoncom.Missing)
            else:
                self._range.AutoFilter(
                    field,
                    pythoncom.Missing if criteria1 is None else criteria1,
                    operator.value,
                    pythoncom.Missing if criteria2 is None else criteria2,
                    visible_dropdown
                )
        except Exception as e:
            raise RangeError(f"自动筛选失败: {e}")

    def resize(self, rows: int = None, cols: int = None) -> 'Range':
        """
        调整范围大小

        Args:
            rows: 行数（None 保持原行数）
            cols: 列数（None 保持原列数）

        Returns:
            调整后的新 Range 对象
        """
        try:
            if rows is None:
                rows = self._range.Rows.Count
            if cols is None:
                cols = self._range.Columns.Count
            # 不使用 COM Resize：部分 pywin32 动态绑定环境下会被误解析为属性，
            # 改用 Cells/Range 按行列号计算新区域
            start_row = self._range.Row
            start_col = self._range.Column
            sheet = self._range.Worksheet
            return Range(sheet.Range(
                sheet.Cells(start_row, start_col),
                sheet.Cells(start_row + rows - 1, start_col + cols - 1)
            ))
        except Exception as e:
            raise RangeError(f"调整范围大小失败: {e}")

    def get_values(self) -> List[List[Any]]:
        """
        获取范围的值（二维列表）

        Returns:
            二维列表的值
        """
        try:
            values = self._range.Value
            if values is None:
                return []

            # 处理单个单元格
            if not isinstance(values, (list, tuple)):
                return [[values]]

            # 处理单行或单列
            if isinstance(values[0], (list, tuple)):
                return [list(row) for row in values]
            else:
                return [list(values)]

        except Exception as e:
            raise RangeError(f"获取值失败: {e}")

    @staticmethod
    def _normalize_values(value: Any) -> Any:
        """将一维列表规范化为二维列表，避免 COM 按列写入或只写首个元素"""
        if isinstance(value, (list, tuple)):
            if len(value) == 0:
                return value
            if not isinstance(value[0], (list, tuple)):
                # 一维列表视为单行数据
                return [list(value)]
            return [list(row) if isinstance(row, (list, tuple)) else [row] for row in value]
        return value

    def set_values(self, values: Union[List[List[Any]], List[Any], Any]):
        """
        设置范围的值

        Args:
            values: 值，可以是二维列表、一维列表或单个值
        """
        try:
            self._range.Value = self._normalize_values(values)
        except Exception as e:
            raise RangeError(f"设置值失败: {e}")

    def clear(self, clear_contents: bool = True, clear_format: bool = True):
        """
        清除范围内容

        Args:
            clear_contents: 是否清除内容
            clear_format: 是否清除格式
        """
        try:
            if clear_contents and clear_format:
                self._range.Clear()
            elif not clear_contents and clear_format:
                self._range.ClearFormats()
            elif clear_contents and not clear_format:
                self._range.ClearContents()
        except Exception as e:
            raise RangeError(f"清除失败: {e}")

    def copy(self, destination: Optional['Range'] = None):
        """
        复制范围

        Args:
            destination: 目标范围，None 表示复制到剪贴板
        """
        try:
            if destination is not None:
                self._range.Copy(Destination=destination.com_object)
            else:
                self._range.Copy()
        except Exception as e:
            raise RangeError(f"复制失败: {e}")

    def paste(self, paste: XlPasteType = XlPasteType.xlPasteAll,
              operation: XlPasteSpecialOperation = XlPasteSpecialOperation.xlPasteSpecialOperationNone,
              skip_blanks: bool = False, transpose: bool = False):
        """
        粘贴（需要在 copy 之后调用）
        paste 粘贴的区域部分
        operation 粘贴操作
        skip_blanks 如果true则不将剪贴板上区域中的空白单元格粘贴到目标区域中
        transpose 如果true则表示在粘贴区域时转置行和列
        """
        try:
            self._range.PasteSpecial(Paste=paste.value, Operation=operation.value,
                                     SkipBlanks=skip_blanks, Transpose=transpose)
        except Exception as e:
            raise RangeError(f"粘贴失败: {e}")

    def autofit(self, columns: bool = True, rows: bool = False):
        """
        自动调整大小

        Args:
            columns: 是否自动调整列宽
            rows: 是否自动调整行高
        """
        try:
            if columns:
                self._range.Columns.AutoFit()
            if rows:
                self._range.Rows.AutoFit()
        except Exception as e:
            raise RangeError(f"自动调整失败: {e}")

    def merge(self, across: bool = False):
        """
        合并单元格

        Args:
            across: 是否跨行合并
        """
        try:
            self._range.Merge(Across=across)
        except Exception as e:
            raise RangeError(f"合并单元格失败: {e}")

    def unmerge(self):
        """取消合并"""
        try:
            self._range.UnMerge()
        except Exception as e:
            raise RangeError(f"取消合并失败: {e}")

    def apply_style(self, style: RangeStyle):
        """
        应用样式

        Args:
            style: 样式配置
        """
        try:
            if style.font_name:
                self._range.Font.Name = style.font_name
            if style.font_size:
                self._range.Font.Size = style.font_size
            if style.font_bold is not None:
                self._range.Font.Bold = style.font_bold
            if style.font_italic is not None:
                self._range.Font.Italic = style.font_italic
            if style.font_underline is not None:
                # 兼容 XlUnderlineStyle 枚举与原始 int 值
                self._range.Font.Underline = getattr(style.font_underline, "value", style.font_underline)
            if style.font_strikethrough is not None:
                self._range.Font.Strikethrough = style.font_strikethrough
            if style.font_color is not None:
                self._range.Font.Color = style.font_color
            elif style.font_color_rgb is not None:
                self._range.Font.Color = rgb_to_excel_color(*style.font_color_rgb)
            if style.fill_color is not None:
                self._range.Interior.Color = style.fill_color
            elif style.fill_color_rgb is not None:
                self._range.Interior.Color = rgb_to_excel_color(*style.fill_color_rgb)
            if style.horizontal_alignment is not None:
                self._range.HorizontalAlignment = style.horizontal_alignment
            if style.vertical_alignment is not None:
                self._range.VerticalAlignment = style.vertical_alignment
            if style.number_format:
                self._range.NumberFormat = style.number_format
            if style.wrap_text is not None:
                self._range.WrapText = style.wrap_text
            if style.borders:
                self._apply_borders(style.borders)
        except Exception as e:
            raise RangeError(f"应用样式失败: {e}")

    def _apply_borders(self, borders: Dict[str, Any]):
        """应用边框配置（见 RangeStyle.borders 说明）"""
        edge_map = {
            "left": XlBordersIndex.xlEdgeLeft,
            "right": XlBordersIndex.xlEdgeRight,
            "top": XlBordersIndex.xlEdgeTop,
            "bottom": XlBordersIndex.xlEdgeBottom,
            "inside_vertical": XlBordersIndex.xlInsideVertical,
            "inside_horizontal": XlBordersIndex.xlInsideHorizontal,
            "diagonal_down": XlBordersIndex.xlDiagonalDown,
            "diagonal_up": XlBordersIndex.xlDiagonalUp,
        }
        rows, cols = self.shape
        for edge, spec in borders.items():
            if not isinstance(spec, dict):
                spec = {"line_style": spec}
            line_style = spec.get("line_style", BorderLineStyle.CONTINUOUS)
            line_style = line_style.value if isinstance(line_style, BorderLineStyle) else line_style
            weight = spec.get("weight", BorderWeight.THIN)
            weight = weight.value if isinstance(weight, BorderWeight) else weight
            color = spec.get("color")

            if edge == "all":
                indices = [XlBordersIndex.xlEdgeLeft, XlBordersIndex.xlEdgeTop,
                           XlBordersIndex.xlEdgeBottom, XlBordersIndex.xlEdgeRight]
                if cols > 1:
                    indices.append(XlBordersIndex.xlInsideVertical)
                if rows > 1:
                    indices.append(XlBordersIndex.xlInsideHorizontal)
            elif edge in edge_map:
                indices = [edge_map[edge]]
            else:
                raise RangeError(f"无效的边框位置: {edge}")

            for idx in indices:
                border = self._range.Borders(idx.value)
                border.LineStyle = line_style
                if line_style != BorderLineStyle.NONE.value:
                    border.Weight = weight
                    if color is not None:
                        border.Color = color

    def clear_conditional_formats(self):
        """清除范围内所有条件格式"""
        try:
            self._range.FormatConditions.Delete()
        except Exception as e:
            raise RangeError(f"清除条件格式失败: {e}")

    def clear_validation(self):
        """清除范围内数据验证"""
        try:
            self._range.Validation.Delete()
        except Exception as e:
            raise RangeError(f"清除数据验证失败: {e}")

    def add_comment(self, text: str):
        """
        添加批注（已存在时覆盖内容）

        Args:
            text: 批注内容
        """
        try:
            existing = self._range.Comment
            if existing is not None:
                try:
                    existing.Text(text)  # 早期绑定：Text 是方法
                except TypeError:
                    existing.Text = text  # 动态绑定：Text 是属性
            else:
                self._range.AddComment(text)
        except Exception as e:
            raise RangeError(f"添加批注失败: {e}")

    @property
    def comment(self) -> Optional[str]:
        """获取或设置批注内容（无批注时为 None，设置时自动创建）"""
        try:
            existing = self._range.Comment
            if existing is None:
                return None
            text = existing.Text
            # 早期绑定下 Text 是方法，动态绑定下是属性
            return text() if callable(text) else text
        except Exception as e:
            raise RangeError(f"获取批注失败: {e}")

    @comment.setter
    def comment(self, text: str):
        self.add_comment(text)

    def delete_comment(self):
        """删除批注"""
        try:
            existing = self._range.Comment
            if existing is not None:
                existing.Delete()
        except Exception as e:
            raise RangeError(f"删除批注失败: {e}")

    def add_hyperlink(self, address: str,
                      text_to_display: str = None,
                      screen_tip: str = None,
                      sub_address: str = None):
        """
        添加超链接

        Args:
            address: 链接目标（URL、文件路径或 mailto:）
            text_to_display: 单元格显示文本
            screen_tip: 鼠标悬停提示
            sub_address: 文档内位置（如 "Sheet2!A1"）
        """
        try:
            # 位置参数：Anchor, Address, SubAddress, ScreenTip, TextToDisplay
            # 注意：SubAddress/ScreenTip 传 Missing 会导致 TextToDisplay 被忽略，需用空串
            self._range.Hyperlinks.Add(
                self._range,
                address,
                sub_address if sub_address is not None else "",
                screen_tip if screen_tip is not None else "",
                pythoncom.Missing if text_to_display is None else text_to_display
            )
        except Exception as e:
            raise RangeError(f"添加超链接失败: {e}")

    def clear_hyperlinks(self):
        """清除范围内所有超链接"""
        try:
            self._range.Hyperlinks.Delete()
        except Exception as e:
            raise RangeError(f"清除超链接失败: {e}")

    def find(self, what: str, after: 'Range' = None, look_in: 'XlFindLookIn' = XlFindLookIn.xlValues,
             look_at: 'XlLookAt' = XlLookAt.xlWhole,
             search_order: 'XlSearchOrder' = XlSearchOrder.xlByRows,
             search_direction: 'XlSearchDirection' = XlSearchDirection.xlNext, match_case: bool = False,
             match_byte: bool = False, search_format: bool = False) -> Optional['Range']:
        """
        查找内容
        Returns:
            找到的范围或 None
        """
        try:
            found = self._range.Find(
                What=what,
                After=pythoncom.Missing if after is None else after.com_object,
                LookIn=look_in.value,
                LookAt=look_at.value,
                SearchOrder=search_order.value,
                SearchDirection=search_direction.value,
                MatchCase=match_case,
                MatchByte=match_byte,
                SearchFormat=search_format
            )
            if found:
                return Range(found)
            return None
        except Exception as e:
            raise RangeError(f"查找失败: {e}")

    def find_all(self, what: str, look_in: int = -4163) -> List['Range']:
        """
        查找所有匹配项

        Args:
            what: 要查找的内容
            look_in: 查找范围

        Returns:
            找到的范围列表
        """
        try:
            first = self._range.Find(
                What=what,
                LookIn=look_in
            )
            if not first:
                return []

            found_ranges = [Range(first)]
            current = first

            while True:
                current = self._range.FindNext(current)
                if not current or current.Address == first.Address:
                    break
                found_ranges.append(Range(current))

            return found_ranges
        except Exception as e:
            raise RangeError(f"查找所有失败: {e}")

    def replace(self, what: str, replacement: str, look_at: 'XlLookAt' = XlLookAt.xlWhole):
        """
        替换内容

        Args:
            what: 要查找的内容
            replacement: 替换为的内容
            look_at: 匹配方式（完全匹配/部分匹配）
        """
        try:
            # Replace 不支持 LookIn 参数，传入会报 TypeError
            self._range.Replace(
                What=what,
                Replacement=replacement,
                LookAt=look_at.value
            )
        except Exception as e:
            raise RangeError(f"替换失败: {e}")

    def sort(self, key_range: 'Range',
             order: int = 1,  # 1=升序, 2=降序
             header: int = 1,  # 1=有标题, 2=无标题
             key2: 'Range' = None,
             order2: int = 1,
             key3: 'Range' = None,
             order3: int = 1):
        """
        排序（支持最多三个关键字）

        Args:
            key_range: 第一排序关键字范围
            order: 第一关键字排序顺序
            header: 是否有标题
            key2: 第二排序关键字范围
            order2: 第二关键字排序顺序
            key3: 第三排序关键字范围
            order3: 第三关键字排序顺序
        """
        try:
            # 使用现代 Sort 对象 + SortFields，传统 Range.Sort 的多关键字 Order2/Order3
            # 在部分绑定环境下不生效
            com_sort = self._range.Worksheet.Sort
            com_sort.SortFields.Clear()
            # 位置参数：Key, SortOn(0=xlSortOnValues), Order
            com_sort.SortFields.Add(key_range.com_object, 0, order)
            if key2 is not None:
                com_sort.SortFields.Add(key2.com_object, 0, order2)
            if key3 is not None:
                com_sort.SortFields.Add(key3.com_object, 0, order3)
            com_sort.SetRange(self._range)
            com_sort.Header = header
            com_sort.Apply()
        except Exception as e:
            raise RangeError(f"排序失败: {e}")

    def to_list(self) -> List[List[Any]]:
        """转换为二维列表（别名）"""
        return self.get_values()

    def to_flat_list(self) -> List[Any]:
        """转换为一维列表"""
        values = self.get_values()
        flat_list = []
        for row in values:
            if isinstance(row, list):
                flat_list.extend(row)
            else:
                flat_list.append(row)
        return flat_list
