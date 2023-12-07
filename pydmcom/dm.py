import os
import struct
import ctypes
from .enums import KeyCode, FindDir

if os.name != 'posix':
    from win32com import client
else:
    from . import dm_mac as client


DLL_PATH = None
REGISTER_DLL_PATH = None

TAG = '[DM]'


def handleRet(ret, ok_value=1):
    h_ret, *values = ret
    # print(f"values {h_ret} {values}")
    if h_ret != ok_value:
        return None
    return values


class DM:
    @staticmethod
    def setDllPath(dll_path: str):
        """
         Args:
            dll_path: dm.dll的绝对路径
        """
        DLL_PATH = dll_path

    @staticmethod
    def setRegisterDllPath(register_dll_path: str):
        """
         Args:
            register_dll_path: RegDll.dll的绝对路径
        """
        REGISTER_DLL_PATH = register_dll_path

    @staticmethod
    def checkEnv() -> bool:
        """
        检查环境是否准备就绪
        """
        if struct.calcsize("P") * 8 == 64:
            print(f"{TAG} not support 64bit python")
            return False
        if DLL_PATH is None:
            print(f"{TAG} please set dll path")
            return False
        if REGISTER_DLL_PATH is None:
            print(f"{TAG} please set register dll path")
            return False
        return True

    def __init__(self) -> None:
        """
        初始化
        """
        print(f"{TAG} 开始初始化 dm.dmsoft")
        try:
            self.__dm = client.Dispatch('dm.dmsoft')
        except Exception:
            print(f"初始化 dm.dmsoft 失败,尝试注册 {DLL_PATH}")
            register_dll = ctypes.cdll.LoadLibrary(REGISTER_DLL_PATH)
            register_dll.SetDllPathW(DLL_PATH, 0)
            self.__dm = client.Dispatch('dm.dmsoft')

        print(f"{TAG} 初始化成功: " + 'VER:', self.ver(), ',ID:', self.GetID(), ',PATH:',
              os.path.join(self.GetBasePath(), 'dm.dll'))

    def SetShowErrorMsg(self, show=1):
        return self.__dm.SetShowErrorMsg(show)

    # ----------------------------基本设置----------------------------------
    def GetBasePath(self):
        return self.__dm.GetBasePath()

    def GetDmCount(self):
        return self.__dm.GetDmCount()

    def GetID(self):
        return self.__dm.GetID()

    def GetLastError(self):
        return self.__dm.GetLastError()

    def GetPath(self):
        return self.__dm.GetPath()

    # 获取版本信息
    def Ver(self):
        return self.__dm.Ver()

    # 注册大漠收费功能
    def Reg(self, key: str, code: str):
        """
        -1 : 无法连接网络,(可能防火墙拦截,如果可以正常访问大漠插件网站，那就可以肯定是被防火墙拦截)
        -2 : 进程没有以管理员方式运行. (出现在win7 vista 2008.建议关闭uac)
        0 : 失败 (未知错误)
        1 : 成功
        2 : 余额不足
        3 : 绑定了本机器，但是账户余额不足50元.
        4 : 注册码错误
        5 : 你的机器或者IP在黑名单列表中或者不在白名单列表中.
        -8 : 版本附加信息长度超过了10
        -9 : 版本附加信息里包含了非法字母.
        """
        ret: int = self.__dm.Reg(key, code)
        return ret

    # ----------------------------窗口----------------------------------

    # 获取窗口信息
    def EnumWindow(self, parent, title, class_name, filter) -> list[int]:
        """
        根据指定条件,枚举系统中符合条件的窗口,可以枚举到按键自带的无法枚举到的窗口

        parent 整形数: 获得的窗口句柄是该窗口的子窗口的窗口句柄,取0时为获得桌面句柄

        title 字符串: 窗口标题. 此参数是模糊匹配.

        class_name 字符串: 窗口类名. 此参数是模糊匹配.

        filter整形数: 取值定义如下

        1 : 匹配窗口标题,参数title有效

        2 : 匹配窗口类名,参数class_name有效.

        4 : 只匹配指定父窗口的第一层孩子窗口

        8 : 匹配所有者窗口为0的窗口,即顶级窗口

        16 : 匹配可见的窗口

        32 : 匹配出的窗口按照窗口打开顺序依次排列 <收费功能，具体详情点击查看>

        这些值可以相加,比如4+8+16就是类似于任务管理器中的窗口列表
        """
        ret: str = self.__dm.EnumWindow(parent or 0, title, class_name, filter)
        if ret == '':
            return []
        return list(map(lambda item: int(item), ret.split(',')))

    # 移动窗口
    def MoveWindow(self, hwnd, x, y):
        return self.__dm.MoveWindow(hwnd, x, y)

    def GetForegroundWindow(self):
        """
        获取顶层活动窗口,可以获取到按键自带插件无法获取到的句柄\n
        返回值:
            整形数:
            返回整型表示的窗口句柄
        """
        return self.__dm.GetForegroundWindow()

    def SetWindowState(self, hwnd, flag):
        """
        设置窗口的状态\n
        hwnd 整形数: 指定的窗口句柄

        flag 整形数: 取值定义如下

        0 : 关闭指定窗口

        1 : 激活指定窗口

        2 : 最小化指定窗口,但不激活

        3 : 最小化指定窗口,并释放内存,但同时也会激活窗口.

        4 : 最大化指定窗口,同时激活窗口.

        5 : 恢复指定窗口 ,但不激活

        6 : 隐藏指定窗口

        7 : 显示指定窗口

        8 : 置顶指定窗口

        9 : 取消置顶指定窗口

        10 : 禁止指定窗口

        11 : 取消禁止指定窗口

        12 : 恢复并激活指定窗口

        13 : 强制结束窗口所在进程.
        """
        return self.__dm.SetWindowState(hwnd, flag)

    def FindWindowEx(self, parent: int, class_name: str, title='') -> int:
        """
        查找符合类名或者标题名的顶层可见窗口,如果指定了parent,则在parent的第一层子窗口中查找.
        parent 整形数: 父窗口句柄，如果为空，则匹配所有顶层窗口

        class 字符串: 窗口类名，如果为空，则匹配所有. 这里的匹配是模糊匹配.

        title 字符串: 窗口标题,如果为空，则匹配所有. 这里的匹配是模糊匹配.
        """
        return self.__dm.findWindowEx(parent, class_name, title)

    # 设置窗口尺寸
    def SetWindowSize(self, hwnd, width, height) -> int:
        return self.__dm.SetWindowSize(hwnd, width, height)

    def GetWindowTitle(self, hwnd) -> str:
        return self.__dm.GetWindowTitle(hwnd)

    # 后台绑定窗口
    def BindWindowEx(self,
                     hwnd,
                     display='dx2',
                     mouse='windows',
                     keypad='windows',
                     public='dx',
                     mode=101) -> int:
        '''
    绑定dx会比较耗时间,请不要频繁调用此函数.
    如果绑定的是dx,要注意不可连续操作dx,中间至少加个10ms的延时,否则可能会导致操作失败.比如绑定图色DX,那么不要连续取色等,键鼠也是一样.
    有些窗口绑定之后必须加一定的延时,否则后台也无效.一般1秒到2秒的延时就足够.
    '''
        self.hwnd = hwnd
        return self.__dm.BindWindowEx(hwnd, display, mouse, keypad, public,
                                      mode)

    # 解绑窗口
    def UnBindWindow(self) -> int:
        return self.__dm.UnBindWindow()

    def EnableFakeActive(self, enabled=1):
        return self.__dm.EnableFakeActive(enabled)

    def EnableIme(self, enabled=1):
        return self.__dm.EnableIme(enabled)

    # ----------------------------文字识别----------------------------------
    def UseDict(self, index) -> int:
        return self.__dm.UseDict(index)

    # 找字
    def FindStr(
        self,
        x1,
        y1,
        x2,
        y2,
        string,
        color='ffffff-000000',
        sim=1.0,
    ) -> int:
        intX = client.VARIANT(-1, 'byref')
        intY = client.VARIANT(-1, 'byref')
        ret = self.__dm.FindStr(x1, y1, x2, y2, string, color, sim, intX, intY)
        return handleRet(ret, 0)

    def Ocr(self,
            x1,
            y1,
            x2,
            y2,
            color_format='ffffff-000000',
            sim=1.0) -> str:

        ret = self.__dm.Ocr(x1, y1, x2, y2, color_format, sim)
        if not ret:
            return None
        return ret

    # ----------------------------图色----------------------------------

    def CapturePng(self, x1, y1, x2, y2, file) -> int:
        return self.__dm.CapturePng(x1, y1, x2, y2, file)

    def EnableDisplayDebug(self, enabled=0):
        return self.__dm.EnableDisplayDebug(enabled)

    def CapturePre(self, file) -> int:
        return self.__dm.CapturePre(file)

    def CmpColor(self, x, y, color, sim) -> int:
        return self.__dm.CmpColor(x, y, color, sim)

    def FindColor(self,
                  x1,
                  y1,
                  x2,
                  y2,
                  color,
                  sim=1.0,
                  dir=FindDir.LeftToRightAndTopToBottom):
        intX = client.VARIANT(-1, 'byref')
        intY = client.VARIANT(-1, 'byref')
        ret = self.__dm.FindColor(x1, y1, x2, y2, color, sim, dir.value, intX,
                                  intY)
        return handleRet(ret)

    def GetColorNum(self, x1,
                    y1,
                    x2,
                    y2,
                    color,
                    sim=1.0):
        return self.__dm.GetColorNum(x1,
                                     y1,
                                     x2,
                                     y2,
                                     color,
                                     sim)

    def FindPic(self,
                x1,
                y1,
                x2,
                y2,
                pic_name,
                delta_color='000000',
                sim=1.0,
                dir=FindDir.LeftToRightAndTopToBottom) -> int:
        intX = client.VARIANT(-1, 'byref')
        intY = client.VARIANT(-1, 'byref')
        # pic_name = f"images/{pic_name}.bmp"
        ret = self.__dm.FindPic(x1, y1, x2, y2, pic_name, delta_color, sim,
                                dir.value, intX, intY)
        return handleRet(ret, 0)

    # ----------------------------键盘----------------------------------
    def KeyDown(self, code: KeyCode) -> int:
        return self.__dm.KeyDown(code.value)

    def KeyPress(self, code: KeyCode) -> int:
        return self.__dm.KeyPress(code.value)

    def KeyUp(self, code: KeyCode) -> int:
        return self.__dm.KeyUp(code.value)

    def KeyPressStr(self, key_str, delay) -> int:
        return self.__dm.KeyPressStr(key_str, delay)

    # ----------------------------鼠标----------------------------------
    def GetCursorPos(self, ) -> int:
        #  const x = new client.Variant(-1, 'byref');
        # const y = new client.Variant(-1, 'byref');
        # return self.__dm.GetCursorPos(x, y)
        # return (1, 2)
        pass

    def LeftClick(self) -> int:
        return self.__dm.LeftClick()

    def RightClick(self) -> int:
        return self.__dm.RightClick()

    def LeftDown(self) -> int:
        return self.__dm.LeftDown()

    def RightDown(self) -> int:
        return self.__dm.RightDown()

    def LeftUp(self) -> int:
        return self.__dm.LeftUp()

    def RightUp(self) -> int:
        return self.__dm.RightUp()

    def WheelDown(self) -> int:
        return self.__dm.WheelDown()

    def WheelUp(self) -> int:
        return self.__dm.WheelUp()

    def MoveTo(self, x, y) -> int:
        return self.__dm.MoveTo(x, y)

    # ----------------------------系统----------------------------------
    def SetScreen(self, width, height, depth) -> int:
        return self.__dm.SetScreen(width, height, depth)

    def GetScreenHeight(self, ) -> int:
        return self.__dm.GetScreenHeight()

    def GetScreenWidth(self, ) -> int:
        return self.__dm.GetScreenWidth()

    def GetScreenDepth(self, ) -> int:
        return self.__dm.GetScreenDepth()

    def GetDiskSerial(self, ) -> int:
        return self.__dm.GetDiskSerial()

    def GetMachineCode(self, ) -> int:
        return self.__dm.GetMachineCode()

    def GetMachineCodeNoMac(self, ) -> int:
        return self.__dm.GetMachineCodeNoMac()

    def GetOsType(self, ) -> int:
        return self.__dm.GetOsType()

    def CheckFontSmooth(self, ) -> int:
        return self.__dm.CheckFontSmooth()

    def DisableFontSmooth(self, ) -> int:
        return self.__dm.DisableFontSmooth()
