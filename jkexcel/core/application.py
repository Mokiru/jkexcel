import pythoncom
import psutil
import time
import ctypes
import ctypes.wintypes
import win32com.client
from contextlib import contextmanager
from typing import Optional, List, Dict, Any, Union
import logging

from jkexcel.core.excel_driver import ExcelApplicationService
from jkexcel.core.workbooks import Workbooks, Workbook
from jkexcel.models.config import ExcelConfig
from jkexcel.models.enums import ExcelType, XlCalculation
from jkexcel.models.exceptions import ExcelNotRunningError, ExcelCOMError

logger = logging.getLogger(__name__)


class ExcelApp:
    """Excel 应用程序封装类"""

    _instances = {}

    _ref_count = 0

    def __new__(cls, *args, **kwargs):
        if kwargs.get('config', None):
            driver = kwargs.get('config').driver
        else:
            driver = ExcelType.OFFICE
        """单例模式"""
        if driver not in cls._instances:
            cls._instances[driver] = super().__new__(cls)
        return cls._instances[driver]

    def __init__(self, config: ExcelConfig = None):
        """
        初始化 ExcelApp
        Args:
            config: Excel 配置
        """
        if not hasattr(self, '_initialized'):
            self._config = config or ExcelConfig()
            self._excel = None
            self._workbooks = None
            self._initialized = False
            self._pid = None
            self._closed_by_workbook = False
            self._event_sink = None

    def __enter__(self):
        """上下文管理器进入"""
        self.start()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        """上下文管理器退出"""
        if self.is_running:
            self.quit()
        return False

    def __repr__(self) -> str:
        status = "运行中" if self.is_running else "已停止"
        return f"<ExcelApp {status} PID={self._pid}>"

    def _cleanup(self):
        """清理资源"""
        if self.is_running:
            try:
                self.quit()
            except:
                pass

    def start(self, visible: bool = None,
              display_alerts: bool = None,
              screen_updating: bool = None) -> 'ExcelApp':
        """
        启动 Excel

        Args:
            visible: 是否可见
            display_alerts: 是否显示警报
            screen_updating: 是否更新屏幕

        Returns:
            self
        """
        if self.is_running:
            return self
        try:
            ExcelApp._ref_count += 1
            if ExcelApp._ref_count <= 1:
                pythoncom.CoInitialize()
            # 创建 Excel 实例
            self._excel = ExcelApplicationService.create_application(excel_type=self._config.driver,
                                                                     throw_exception=True)
            self._pid = self._get_process_id()

            # 应用配置
            if visible is not None:
                self._excel.Visible = visible
            else:
                self._excel.Visible = self._config.visible

            if display_alerts is not None:
                self._excel.DisplayAlerts = display_alerts
            else:
                self._excel.DisplayAlerts = self._config.display_alerts

            if screen_updating is not None:
                self._excel.ScreenUpdating = screen_updating
            else:
                self._excel.ScreenUpdating = self._config.screen_updating

            self._excel.EnableEvents = self._config.enable_events

            # 初始化 Workbooks
            self._workbooks = Workbooks(self)
            self._initialized = True

            logger.info(f"Excel 启动成功 (PID: {self._pid})")
            return self

        except Exception as e:
            ExcelApp._ref_count -= 1
            self._initialized = False
            raise ExcelNotRunningError(f"启动 Excel 失败: {e}")

    def _get_process_id(self) -> Optional[int]:
        """获取进程 ID"""
        try:
            # Hwnd 是窗口句柄而非 PID，需通过窗口句柄反查进程 ID
            hwnd = self._excel.Hwnd
            pid = ctypes.wintypes.DWORD()
            ctypes.windll.user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                return pid.value
        except:
            pass
        # 通过驱动对应的进程名查找
        if self._config.driver == ExcelType.WPS:
            # WPS 表格实际进程名可能为 et.exe/ket.exe，wps.exe 为集成主进程
            process_names = {'wps.exe', 'et.exe', 'ket.exe'}
        else:
            # Microsoft Office Excel
            process_names = {'excel.exe'}
        for proc in psutil.process_iter(['pid', 'name']):
            if proc.info['name'] and proc.info['name'].lower() in process_names:
                return proc.info['pid']
        return None

    @property
    def is_running(self) -> bool:
        """Excel 是否在运行"""
        if not self._excel or not self._initialized:
            return False

        try:
            # 尝试访问属性来检查是否存活
            _ = self._excel.Version
            return True
        except:
            return False

    @property
    def version(self) -> str:
        """获取 Excel 版本"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")

        try:
            return self._excel.Version
        except Exception as e:
            raise ExcelCOMError(f"获取版本失败: {e}")

    @property
    def com_object(self):
        """获取底层 COM 对象"""
        return self._excel

    @property
    def count(self):
        """获取工作簿数量"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")
        return self._workbooks.count

    @property
    def workbooks(self) -> Workbooks:
        """获取工作簿集合"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")
        return self._workbooks

    @property
    def active_workbook(self) -> Optional[Workbook]:
        """获取活动工作簿"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")

        try:
            com_wb = self._excel.ActiveWorkbook
            if com_wb:
                return Workbook(com_wb, self)
            return None
        except Exception as e:
            raise ExcelCOMError(f"获取活动工作簿失败: {e}")

    @property
    def active_sheet(self):
        """获取活动工作表（通过活动工作簿）"""
        wb = self.active_workbook
        if wb:
            return wb.get_active_sheet()
        return None

    def calculate(self):
        """强制计算所有打开的工作簿"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")

        try:
            self._excel.Calculate()
        except Exception as e:
            raise ExcelCOMError(f"计算失败: {e}")

    def _get_app_attr(self, attr: str):
        """读取应用级 COM 属性"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")
        try:
            return getattr(self._excel, attr)
        except Exception as e:
            raise ExcelCOMError(f"获取 {attr} 失败: {e}")

    def _set_app_attr(self, attr: str, value):
        """设置应用级 COM 属性"""
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")
        try:
            setattr(self._excel, attr, value)
        except Exception as e:
            raise ExcelCOMError(f"设置 {attr} 失败: {e}")

    @property
    def visible(self) -> bool:
        """获取或设置 Excel 窗口可见性"""
        return self._get_app_attr("Visible")

    @visible.setter
    def visible(self, value: bool):
        self._set_app_attr("Visible", value)

    @property
    def screen_updating(self) -> bool:
        """获取或设置屏幕刷新（批量操作时关闭可大幅提升性能）"""
        return self._get_app_attr("ScreenUpdating")

    @screen_updating.setter
    def screen_updating(self, value: bool):
        self._set_app_attr("ScreenUpdating", value)

    @property
    def display_alerts(self) -> bool:
        """获取或设置警告弹窗"""
        return self._get_app_attr("DisplayAlerts")

    @display_alerts.setter
    def display_alerts(self, value: bool):
        self._set_app_attr("DisplayAlerts", value)

    @property
    def enable_events(self) -> bool:
        """获取或设置事件触发"""
        return self._get_app_attr("EnableEvents")

    @enable_events.setter
    def enable_events(self, value: bool):
        self._set_app_attr("EnableEvents", value)

    @property
    def calculation(self) -> XlCalculation:
        """获取或设置计算模式"""
        return XlCalculation(self._get_app_attr("Calculation"))

    @calculation.setter
    def calculation(self, value: Union[XlCalculation, int]):
        # Excel 限制：没有打开的工作簿时无法设置 Calculation，此时跳过
        if self._excel.Workbooks.Count == 0:
            return
        self._set_app_attr("Calculation", value.value if isinstance(value, XlCalculation) else value)

    @contextmanager
    def suspend(self, screen_updating: bool = False,
                enable_events: bool = False,
                display_alerts: bool = False,
                calculation: XlCalculation = XlCalculation.xlCalculationManual):
        """
        批量操作上下文：临时关闭屏幕刷新/事件/弹窗并切换计算模式，退出时自动恢复

        用法：
            with app.suspend():
                ws.write_data(big_data)
            # 退出后自动恢复原状态并重新计算

        Args:
            screen_updating: 上下文内的屏幕刷新状态
            enable_events: 上下文内的事件触发状态
            display_alerts: 上下文内的警告弹窗状态
            calculation: 上下文内的计算模式
        """
        saved = {
            "ScreenUpdating": self.screen_updating,
            "EnableEvents": self.enable_events,
            "DisplayAlerts": self.display_alerts,
            "Calculation": self._get_app_attr("Calculation"),
        }
        try:
            self.screen_updating = screen_updating
            self.enable_events = enable_events
            self.display_alerts = display_alerts
            self.calculation = calculation
            yield self
        finally:
            # 先恢复计算模式再恢复屏幕刷新，避免恢复时大量重算闪烁
            for attr in ("Calculation", "EnableEvents", "DisplayAlerts", "ScreenUpdating"):
                try:
                    setattr(self._excel, attr, saved[attr])
                except Exception:
                    pass

    def wait(self, seconds: float):
        """
        等待指定时间

        Args:
            seconds: 等待秒数
        """
        time.sleep(seconds)

    def quit(self):
        """
        退出 Excel
        """
        if not self.is_running or not self._excel:
            return
        # 先断开事件监听，避免退出过程中回调访问已释放对象
        try:
            self.detach_events()
        except Exception:
            pass
        # 关闭所有工作簿
        for wb in self.workbooks:
            try:
                wb.close(save_changes=False)
            except:
                pass
        if not self._excel:
            return
        self._excel.Quit()
        ExcelApp._ref_count -= 1
        logger.info("Excel 已退出")
        self._excel = None
        self._workbooks = None
        self._initialized = False
        self._pid = None
        if ExcelApp._ref_count <= 0:
            try:
                pythoncom.CoUninitialize()
            except:
                pass

    def create_workbook(self, *args, **kwargs) -> Workbook:
        """
        创建新工作簿

        Returns:
            Workbook 对象
        """
        return self.workbooks.add(*args, **kwargs)

    def open_workbook(self, *args, **kwargs) -> Workbook:
        """
        打开工作簿

        Args:
        Returns:
            Workbook 对象
        """
        return self.workbooks.open(*args, **kwargs)

    def close_all_workbooks(self, save_changes: bool = False):
        """
        关闭所有工作簿

        Args:
            save_changes: 是否保存更改
        """
        self.workbooks.close_all(save_changes=save_changes)

    def run_macro(self, macro_name: str, *args):
        """
        运行宏

        Args:
            macro_name: 宏名称
            *args: 宏参数
        """
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")

        try:
            return self._excel.Run(macro_name, *args)
        except Exception as e:
            raise ExcelCOMError(f"运行宏失败: {e}")

    def attach_events(self, handler):
        """
        监听应用级事件

        Args:
            handler: ExcelEvents 子类实例，重写关心的 on_* 方法

        Note:
            事件通过 COM 消息泵分发，脚本中需定期调用 pump_events()
        """
        if not self.is_running:
            raise ExcelNotRunningError("Excel 未运行")
        if self._event_sink is not None:
            raise ExcelCOMError("已监听事件，请先 detach_events")
        try:
            from jkexcel.core.events import _AppEventSink
            _AppEventSink._handler = handler
            self._event_sink = win32com.client.WithEvents(self._excel, _AppEventSink)
            return self._event_sink
        except Exception as e:
            raise ExcelCOMError(f"监听事件失败: {e}")

    def detach_events(self):
        """断开事件监听"""
        if self._event_sink is not None:
            try:
                self._event_sink.close()
            finally:
                self._event_sink = None
                from jkexcel.core.events import _AppEventSink
                _AppEventSink._handler = None

    @staticmethod
    def pump_events():
        """分发待处理的 COM 事件（在消息循环/等待中定期调用）"""
        pythoncom.PumpWaitingMessages()

    def get_open_workbook_names(self) -> List[str]:
        """
        获取所有打开的工作簿名称

        Returns:
            工作簿名称列表
        """
        return self.workbooks.names

    def find_workbook_by_name(self, name: str) -> Optional[Workbook]:
        """
        通过名称查找工作簿

        Args:
            name: 工作簿名称

        Returns:
            Workbook 对象或 None
        """
        for wb in self.workbooks:
            if wb.name.lower() == name.lower():
                return wb
        return None

    def get_process_info(self) -> Dict[str, Any]:
        """
        获取进程信息

        Returns:
            进程信息字典
        """
        if not self._pid:
            return {}

        try:
            proc = psutil.Process(self._pid)
            return {
                'pid': self._pid,
                'name': proc.name(),
                'status': proc.status(),
                'cpu_percent': proc.cpu_percent(),
                'memory_percent': proc.memory_percent(),
                'create_time': proc.create_time(),
                'exe': proc.exe(),
                'cwd': proc.cwd(),
                'cmdline': proc.cmdline(),
                'username': proc.username(),
            }
        except:
            return {'pid': self._pid}

    @classmethod
    def get_all_instances(cls) -> List['ExcelApp']:
        """获取所有 ExcelApp 实例"""
        instances = list(cls._instances.values())
        return instances

    @classmethod
    def cleanup_all(cls):
        """清理所有实例"""
        for instance in cls.get_all_instances():
            try:
                instance.quit()
            except:
                pass
        cls._instances.clear()
