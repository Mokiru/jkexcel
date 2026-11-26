"""Excel 应用级事件监听

用法：
    class MyEvents(ExcelEvents):
        def on_sheet_change(self, sheet, target):
            print("单元格变化:", target.Address)

    app.attach_events(MyEvents())
    # 事件通过 COM 消息泵分发，脚本中需定期调用：
    app.pump_events()
    # 结束时：
    app.detach_events()

注意：
    - 回调参数为原始 COM 对象，可自行用 Worksheet/Range 包装
    - 回调在 pump_events()/消息循环中被触发，需保证 COM 已初始化（同一线程）
    - BeforeSave/BeforeClose 的 cancel 参数仅供读取，暂不支持取消操作
"""
import win32com.client


class ExcelEvents:
    """应用事件处理器基类：子类化并重写关心的 on_* 方法"""

    def on_sheet_change(self, sheet, target):
        """单元格内容变化"""

    def on_sheet_selection_change(self, sheet, target):
        """选区变化"""

    def on_sheet_activate(self, sheet):
        """工作表被激活"""

    def on_workbook_open(self, workbook):
        """工作簿被打开"""

    def on_new_workbook(self, workbook):
        """新建工作簿"""

    def on_workbook_activate(self, workbook):
        """工作簿被激活"""

    def on_workbook_before_close(self, workbook, cancel):
        """工作簿关闭前"""

    def on_workbook_before_save(self, workbook, save_as_ui, cancel):
        """工作簿保存前"""

    def on_sheet_before_delete(self, sheet, cancel):
        """工作表删除前"""


class _AppEventSink:
    """COM 事件接收器：将 Application 事件转发给 ExcelEvents 处理器"""

    _handler = None

    def _dispatch(self, method_name, *args):
        handler = self._handler
        if handler is not None:
            wrapped = [self._wrap_arg(a) for a in args]
            getattr(handler, method_name)(*wrapped)

    @staticmethod
    def _wrap_arg(arg):
        """WPS 传入的事件参数是原始 PyIDispatch，需包装为 Dispatch 对象才能访问属性"""
        if arg is None or isinstance(arg, (bool, int, float, str, tuple, list)):
            return arg
        try:
            return win32com.client.Dispatch(arg)
        except Exception:
            return arg

    def OnSheetChange(self, sh, target):
        self._dispatch("on_sheet_change", sh, target)

    def OnSheetSelectionChange(self, sh, target):
        self._dispatch("on_sheet_selection_change", sh, target)

    def OnSheetActivate(self, sh):
        self._dispatch("on_sheet_activate", sh)

    def OnWorkbookOpen(self, wb):
        self._dispatch("on_workbook_open", wb)

    def OnNewWorkbook(self, wb):
        self._dispatch("on_new_workbook", wb)

    def OnWorkbookActivate(self, wb):
        self._dispatch("on_workbook_activate", wb)

    def OnWorkbookBeforeClose(self, wb, cancel):
        self._dispatch("on_workbook_before_close", wb, cancel)

    def OnWorkbookBeforeSave(self, wb, save_as_ui, cancel):
        self._dispatch("on_workbook_before_save", wb, save_as_ui, cancel)

    def OnSheetBeforeDelete(self, sh, cancel):
        self._dispatch("on_sheet_before_delete", sh, cancel)