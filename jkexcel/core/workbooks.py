import os
from typing import List, Optional, Iterator, Union, Dict

import pythoncom

from jkexcel.core.worksheet import Worksheet
from jkexcel.core.worksheets import Worksheets
from jkexcel.models.config import SaveFormat, SaveAsAccessMode, SaveConflictResolution
from jkexcel.models.enums import Platform, SeparatorFormat, CorruptLoad
from jkexcel.models.exceptions import WorkbookNotFoundError


class Workbook:
    """Excel 工作簿封装类"""

    def __init__(self, com_workbook, excel_app):
        """
        初始化 Workbook

        Args:
            com_workbook: COM Workbook 对象
            excel_app: ExcelApp 实例
        """
        if com_workbook is None:
            raise WorkbookNotFoundError("COM Workbook 对象不能为 None")
        self._workbook = com_workbook
        self._excel = excel_app
        self._worksheets = None

    def __repr__(self) -> str:
        return f"<Workbook '{self.name}'>"

    def __enter__(self):
        """上下文管理器进入"""
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器退出"""
        self.close(save_changes=False)
        return False

    @property
    def com_object(self):
        """获取底层 COM 对象"""
        return self._workbook

    @property
    def excel_app(self):
        """获取 ExcelApp 实例"""
        return self._excel

    @property
    def name(self) -> str:
        """获取工作簿名称"""
        try:
            return self._workbook.Name
        except Exception as e:
            raise WorkbookNotFoundError(f"获取名称失败: {e}")

    @property
    def full_name(self) -> str:
        """获取完整路径"""
        try:
            return self._workbook.FullName
        except Exception as e:
            raise WorkbookNotFoundError(f"获取完整路径失败: {e}")

    @property
    def path(self) -> str:
        """获取路径"""
        try:
            return self._workbook.Path
        except Exception as e:
            raise WorkbookNotFoundError(f"获取路径失败: {e}")

    @property
    def saved(self) -> bool:
        """是否已保存"""
        try:
            return self._workbook.Saved
        except Exception as e:
            raise WorkbookNotFoundError(f"获取保存状态失败: {e}")

    @property
    def read_only(self) -> bool:
        """是否只读"""
        try:
            return self._workbook.ReadOnly
        except Exception as e:
            raise WorkbookNotFoundError(f"获取只读状态失败: {e}")

    @property
    def file_format(self) -> int:
        """获取文件格式"""
        try:
            return self._workbook.FileFormat
        except Exception as e:
            raise WorkbookNotFoundError(f"获取文件格式失败: {e}")

    @property
    def worksheets(self) -> Worksheets:
        """获取工作表集合"""
        if self._worksheets is None:
            self._worksheets = Worksheets(self._workbook.Worksheets, self._excel)
        return self._worksheets

    @property
    def sheets(self) -> Worksheets:
        """获取所有表（包括图表等）（别名）"""
        return Worksheets(self._workbook.Sheets, self._excel)

    def activate(self):
        """激活工作簿"""
        try:
            self._workbook.Activate()
        except Exception as e:
            raise WorkbookNotFoundError(f"激活工作簿失败: {e}")

    def save(self):
        """保存工作簿"""
        try:
            self._workbook.Save()
        except Exception as e:
            raise WorkbookNotFoundError(f"保存工作簿失败: {e}")

    def save_as(self, file_path: str, file_format: Optional[SaveFormat] = None,
                password: Optional[str] = pythoncom.Missing,
                write_res_password: Optional[str] = pythoncom.Missing,
                read_only_recommended: Optional[bool] = pythoncom.Missing,
                create_backup: Optional[bool] = pythoncom.Missing,
                access_mode: Optional[SaveAsAccessMode] = None,
                conflict_resolution: Optional[SaveConflictResolution] = None,
                add_to_mru: Optional[bool] = pythoncom.Missing,
                text_code_page: Optional[str] = pythoncom.Missing,
                text_visual_layout: Optional[object] = pythoncom.Missing,
                local: Optional[bool] = pythoncom.Missing):
        """
        另存为
        :param file_path:
        :param file_format:
        :param password:
        :param write_res_password:
        :param read_only_recommended:
        :param create_backup:
        :param access_mode:
        :param conflict_resolution:
        :param add_to_mru:
        :param text_code_page:
        :param text_visual_layout:
        :param local:
        :return:
        """
        if file_format:
            name_without_ext, _ = os.path.splitext(file_path)
            file_name = name_without_ext + file_format.value[1]
        full_path = os.path.abspath(file_path)
        dir_path = os.path.dirname(full_path)
        os.makedirs(dir_path, exist_ok=True)
        file_name = os.path.normpath(file_path)
        try:
            self._workbook.SaveAs(file_name, file_format.value[0] if file_format else pythoncom.Missing, password,
                                  write_res_password,
                                  read_only_recommended, create_backup,
                                  access_mode.value[0] if access_mode else pythoncom.Missing,
                                  conflict_resolution.value[0] if conflict_resolution else pythoncom.Missing,
                                  add_to_mru,
                                  text_code_page,
                                  text_visual_layout, local)
        except Exception as e:
            raise WorkbookNotFoundError(f"另存为失败: {e}")

    def save_copy_as(self, file_path: str):
        """
        保存副本

        Args:
            file_path: 副本路径
        """
        try:
            self._workbook.SaveCopyAs(file_path)
        except Exception as e:
            raise WorkbookNotFoundError(f"保存副本失败: {e}")

    def close(self, save_changes: bool = True,
              file_path: str = None):
        """
        关闭工作簿

        Args:
            save_changes: 是否保存更改
            file_path: 保存路径（如果为None则使用原路径）
        """
        try:
            if save_changes and file_path:
                self.save_as(file_path)
            elif save_changes:
                self.save()
            self._excel.workbooks.on_workbook_closed(self)
            self._workbook.Close(SaveChanges=False)
            self._worksheets = None
        except Exception as e:
            raise WorkbookNotFoundError(f"关闭工作簿失败: {e}")

    def protect(self, password: str = "",
                structure: bool = True,
                windows: bool = False):
        """
        保护工作簿

        Args:
            password: 密码
            structure: 保护结构
            windows: 保护窗口
        """
        try:
            self._workbook.Protect(
                Password=password,
                Structure=structure,
                Windows=windows
            )
        except Exception as e:
            raise WorkbookNotFoundError(f"保护工作簿失败: {e}")

    def unprotect(self, password: str = ""):
        """
        取消保护

        Args:
            password: 密码
        """
        try:
            self._workbook.Unprotect(Password=password)
        except Exception as e:
            raise WorkbookNotFoundError(f"取消保护失败: {e}")

    def refresh_all(self):
        """刷新所有数据"""
        try:
            self._workbook.RefreshAll()
        except Exception as e:
            raise WorkbookNotFoundError(f"刷新数据失败: {e}")

    def calculate(self):
        """计算所有公式"""
        try:
            self._workbook.Application.Calculate()
        except Exception as e:
            raise WorkbookNotFoundError(f"计算失败: {e}")

    def get_active_sheet(self) -> Worksheet:
        """获取活动工作表"""
        try:
            return Worksheet(self._workbook.ActiveSheet, self._excel)
        except Exception as e:
            raise WorkbookNotFoundError(f"获取活动工作表失败: {e}")

    def add_worksheet(self, before: Union[int, str, Worksheet] = None,
                      after: Union[int, str, Worksheet] = None) -> Worksheet:
        """
        添加新工作表

        Args:
            before: 插入到指定工作表之前
            after: 插入到指定工作表之后

        Returns:
            Worksheet 对象
        """
        return self.worksheets.add(before=before, after=after)

    def get_worksheet(self, key: Union[int, str]) -> Worksheet:
        """
        获取工作表

        Args:
            key: 索引或名称

        Returns:
            Worksheet 对象
        """
        return self.worksheets.get(key)

    def print_out(self, copies: int = 1,
                  preview: bool = False,
                  active_printer: str = None,
                  print_to_file: bool = False,
                  collate: bool = True):
        """
        打印

        Args:
            copies: 份数
            preview: 是否预览
            active_printer: 打印机
            print_to_file: 打印到文件
            collate: 是否逐份打印
        """
        try:
            if preview:
                self._workbook.PrintPreview()
            else:
                self._workbook.PrintOut(
                    Copies=copies,
                    ActivePrinter=active_printer,
                    PrintToFile=print_to_file,
                    Collate=collate
                )
        except Exception as e:
            raise WorkbookNotFoundError(f"打印失败: {e}")

    def export_as_pdf(self, file_path: str,
                      quality: int = 0,  # xlQualityStandard
                      include_doc_props: bool = True,
                      ignore_print_areas: bool = False):
        """
        导出为 PDF

        Args:
            file_path: 保存路径
            quality: 质量
            include_doc_props: 包含文档属性
            ignore_print_areas: 忽略打印区域
        """
        try:
            self._workbook.ExportAsFixedFormat(
                Type=0,  # xlTypePDF
                Filename=file_path,
                Quality=quality,
                IncludeDocProperties=include_doc_props,
                IgnorePrintAreas=ignore_print_areas
            )
        except Exception as e:
            raise WorkbookNotFoundError(f"导出 PDF 失败: {e}")

    def export_as_xps(self, file_path: str):
        """
        导出为 XPS

        Args:
            file_path: 保存路径
        """
        try:
            self._workbook.ExportAsFixedFormat(
                Type=1,  # xlTypeXPS
                Filename=file_path
            )
        except Exception as e:
            raise WorkbookNotFoundError(f"导出 XPS 失败: {e}")


class Workbooks:
    """Excel 工作簿集合封装类"""

    def __init__(self, com_workbooks, excel_app):
        """
        初始化 Workbooks

        Args:
            com_workbooks: COM Workbooks 对象
            excel_app: ExcelApp 实例
        """
        if com_workbooks is None:
            raise WorkbookNotFoundError("COM Workbooks 对象不能为 None")
        self._workbooks = com_workbooks
        self._workbook_register: Dict[str, Optional[Workbook]] = {}
        self._excel = excel_app

    def __repr__(self) -> str:
        return f"<Workbooks count={self.count}>"

    def __len__(self) -> int:
        """获取工作簿数量"""
        return self.count

    def __getitem__(self, key: Union[int, str]) -> Workbook:
        """通过索引或名称获取工作簿"""
        return self.get(key)

    def __iter__(self) -> Iterator[Workbook]:
        """迭代工作簿"""
        for item in list(self._workbook_register.values()):
            yield item

    @property
    def com_object(self):
        """获取底层 COM 对象"""
        return self._workbooks

    @property
    def excel_app(self):
        """获取 ExcelApp 实例"""
        return self._excel

    @property
    def count(self) -> int:
        """获取工作簿数量"""
        try:
            return self._workbooks.Count
        except Exception as e:
            raise WorkbookNotFoundError(f"获取工作簿数量失败: {e}")

    @property
    def names(self) -> List[str]:
        """获取所有工作簿名称"""
        return [wb.name for wb in self]

    def register_workbook(self, wb: Workbook):
        """
        注册工作簿

        Args:
            wb: 工作簿对象
        """
        self._workbook_register[wb.name] = wb

    def get(self, key: Union[int]) -> Workbook:
        """
        获取工作簿

        Args:
            key: 索引（从1开始）

        Returns:
            Workbook 对象
        """
        try:
            if isinstance(key, int):
                return list(self._workbook_register.values())[key - 1]
            elif isinstance(key, str):
                return self._workbook_register[key]
            else:
                raise WorkbookNotFoundError("索引或名称不能为 None")
        except Exception as e:
            raise WorkbookNotFoundError(f"获取工作簿失败: {e}")

    def add(self, *args, **kwargs) -> Workbook:
        """
        添加新工作簿 并保存

        Returns:
            Workbook 对象
        """
        try:
            com_wb = self._workbooks.Add()
            wb = Workbook(com_wb, self._excel)
            if 'file_path' in kwargs:
                wb.save_as(*args, **kwargs)
            self._workbook_register[wb.name] = wb
            return wb
        except Exception as e:
            raise WorkbookNotFoundError(f"添加工作簿失败: {e}")

    def open(self, file_path: str,
             update_links: bool = True,
             read_only: bool = False,
             sep_format: Optional[SeparatorFormat] = None,
             password: str = pythoncom.Missing,
             write_res_password: str = pythoncom.Missing,
             ignore_read_only_recommended: bool = True,
             origin: Optional[Platform] = None,
             delimiter: str = pythoncom.Missing,
             editable: bool = pythoncom.Missing,
             notify: bool = pythoncom.Missing,
             converter: int = pythoncom.Missing,
             add_to_mru: bool = False,
             local: bool = False,
             corrupt_load: Optional[CorruptLoad] = CorruptLoad.xlNormalLoad
             ) -> Workbook:
        """
        打开工作簿

        Args:
            file_path: 文件路径
            update_links: 打开工作簿时是否更新外部引用
            read_only: 是否只读
            sep_format: 如果 Microsoft Excel 打开文本文件，则此参数指定分隔符字符。 如果省略此参数，则使用当前分隔符
            password: 包含打开受保护工作簿所需密码的字符串。 如果省略此参数并且工作簿需要密码，则会提示用户输入密码。
            write_res_password: 包含写入写保护的工作簿所需密码的字符串。 如果省略此参数并且工作簿需要密码，则将提示用户输入密码。
            ignore_read_only_recommended: 如果为 True，则不让 Microsoft Excel 显示只读的建议消息（如果该工作簿以建议只读选项保存）。
            origin: 如果文件是文本文件，则此参数表示其来源，这样就可正确映射代码页和回车/换行 (CR/LF)。 可以是以下 `Platform` 常量之一： `xlMacintosh` `xlWindows` 或 `xlMSDOS` 如果省略此参数，则使用当前操作系统。
            delimiter: 如果sep_format设置为 `SeparatorFormat.Custom` 则该表示使用自定义的字符串，注意这里只会选择字符串第一个字符作为分隔符
            editable: 如果文件为 Microsoft Excel 4.0 外接程序，则此参数为 True 时可打开该外接程序以使其成为可见窗口。 如果此参数为 False 或被省略，则以隐藏方式打开外接程序，并且无法设为可见。 此选项不适用于在 Microsoft Excel 5.0 或更高版本中创建的加载项。如果文件是 Excel 模板，则为 True，可打开指定的模板进行编辑。 如果为 False，则可根据指定的模板打开新工作簿
            notify: 当文件不能以可读写模式打开时，如果此参数为 `True`，则可将该文件添加到文件通知列表。 Microsoft Excel 将以只读模式打开该文件并轮询文件通知列表，并在文件可用时向用户发出通知。 如果此参数为 `False` 或省略，则不会请求通知，并且任何打开不可用文件的尝试都将失败。
            converter: 打开文件时要尝试的第一个文件转换器的索引。 首先尝试指定的文件转换器
            add_to_mru: 如果为 `True`，则将该工作簿添加到最近使用的文件列表中。
            local:如果为 True，则以 Microsoft Excel（包括控制面板设置）的语言保存文件。 如果为 False（默认值），则以 Visual Basic for Applications (VBA) 的语言保存文件，其中 Visual Basic for Applications (VBA) 通常为美国英语版本，除非从中运行 Workbooks.Open 的 VBA 项目是旧的已国际化的 XL5/95 VBA 项目
            corrupt_load: 可为以下常量之一：`xlNormalLoad`、`xlRepairFile` 和 `xlExtractData`。 如果未指定值，则默认行为为 `xlNormalLoad`，并且不会在通过 OM 启动时尝试恢复。
        Returns:
            Workbook 对象
        """
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"文件不存在: {file_path}")

        try:
            com_wb = self._workbooks.Open(
                Filename=file_path,
                UpdateLinks=0 if update_links else 3,  # 0=更新, 3=不更新
                ReadOnly=read_only,
                Format=sep_format.value if sep_format else pythoncom.Missing,
                Password=password,
                WriteResPassword=write_res_password,
                IgnoreReadOnlyRecommended=ignore_read_only_recommended,
                Origin=origin.value if origin else pythoncom.Missing,
                Delimiter=delimiter,
                Editable=editable,
                Notify=notify,
                Converter=converter,
                AddToMru=add_to_mru,
                Local=local,
                CorruptLoad=corrupt_load.value
            )
            wb = Workbook(com_wb, self._excel)
            self._workbook_register[wb.name] = wb
            return wb
        except Exception as e:
            raise WorkbookNotFoundError(f"打开工作簿失败: {e}")

    def close_all(self, save_changes: bool = False):
        """
        关闭所有工作簿

        Args:
            save_changes: 是否保存更改
        """
        for workbook in self:
            try:
                workbook.close(save_changes=save_changes)
            except:
                pass

    def exists(self, name: str) -> bool:
        """
        检查工作簿是否存在

        Args:
            name: 工作簿名称

        Returns:
            bool
        """
        return name in self._workbook_register

    def on_workbook_closed(self, wb: Workbook):
        self._workbook_register.pop(wb.name)
