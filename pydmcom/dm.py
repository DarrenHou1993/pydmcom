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

TAG = '[pydmcom]'


def handleRet(ret, ok_value=1):
    h_ret, *values = ret
    print(f"values {h_ret} {values}")
    if h_ret != ok_value:
        return None
    return values


class DM:
    @staticmethod
    def setDllPath(dll_path: str):
        """
        设置Dll文件的路径

        参数:
            dll_path: dm.dll的绝对路径
        """
        DLL_PATH = dll_path

    @staticmethod
    def setRegisterDllPath(register_dll_path: str):
        """
        设置注册Dll文件的路径

        参数:
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

        print(
            f"{TAG} 初始化成功: + VER: {self.Ver()}, ID: {self.GetID()}, PATH: {self.GetBasePath()}")

    def SetShowErrorMsg(self, show=1) -> int:
        """
        简介:
            设置是否弹出错误信息,默认是打开.
        返回值:
            0 : 失败
            1 : 成功
        """
        return self.__dm.SetShowErrorMsg(show)

    # ----------------------------基本设置----------------------------------

    def SetPath(self, path):
        """
        简介:
            设置全局路径,设置了此路径后,所有接口调用中,相关的文件都相对于此路径. 比如图片,字库等.
        参数定义:
            path: 路径,可以是相对路径,也可以是绝对路径.
        返回值:
            0 : 失败
            1 : 成功
        """
        return self.__dm.SetPath(path)

    def GetBasePath(self) -> str:
        """
        获取注册在系统中的dm.dll的路径

        返回值:
            str: dm.dll的绝对路径.
        """
        return self.__dm.GetBasePath()

    def GetDmCount(self) -> int:
        """
        返回当前进程已经创建的dm对象个数.

        返回值:
            int: 当前进程已经创建的dm对象个数.
        """
        return self.__dm.GetDmCount()

    def GetID(self) -> int:
        """
        返回当前大漠对象的ID值，这个值对于每个对象是唯一存在的。可以用来判定两个大漠对象是否一致.

        返回值:
            int: 大漠对象的ID值.
        """
        return self.__dm.GetID()

    def GetLastError(self):
        """
        获取插件命令的最后错误

        返回值:
            int: 错误码.

            返回值表示错误值。 0表示无错误.

            -1 : 表示你使用了绑定里的收费功能，但是没注册，无法使用.

            -2 : 使用模式0 2 4 6时出现，因为目标窗口有保护，或者目标窗口没有以管理员权限打开. 常见于win7以上系统.或者有安全软件拦截插件.解决办法: 关闭所有安全软件，并且关闭系统UAC,然后再重新尝试. 如果还不行就可以肯定是目标窗口有特殊保护.

            -3 : 使用模式0 2 4 6时出现，可能目标窗口有保护，也可能是异常错误.

            -4 : 使用模式1 3 5 7 101 103时出现，这是异常错误.

            -5 : 使用模式1 3 5 7 101 103时出现, 这个错误的解决办法就是关闭目标窗口，重新打开再绑定即可. 也可能是运行脚本的进程没有管理员权限.

            -6 -7 -9 : 使用模式1 3 5 7 101 103时出现,异常错误. 还有可能是安全软件的问题，比如360等。尝试卸载360.

            -8 -10 : 使用模式1 3 5 7 101 103时出现, 目标进程可能有保护,也可能是插件版本过老，试试新的或许可以解决.

            -11 : 使用模式1 3 5 7 101 103时出现, 目标进程有保护. 告诉我解决。

            -12 : 使用模式1 3 5 7 101 103时出现, 目标进程有保护. 告诉我解决。

            -13 : 使用模式1 3 5 7 101 103时出现, 目标进程有保护. 或者是因为上次的绑定没有解绑导致。 尝试在绑定前调用ForceUnBindWindow.

            -14 : 使用模式0 1 4 5时出现, 有可能目标机器兼容性不太好. 可以尝试其他模式. 比如2 3 6 7

            -16 : 可能使用了绑定模式 0 1 2 3 和 101，然后可能指定了一个子窗口.导致不支持.可以换模式4 5 6 7或者103来尝试. 另外也可以考虑使用父窗口或者顶级窗口.来避免这个错误。还有可能是目标窗口没有正常解绑 然后再次绑定的时候.

            -17 : 模式1 3 5 7 101 103时出现. 这个是异常错误. 告诉我解决.

            -18 : 句柄无效.

            -19 : 使用模式0 1 2 3 101时出现,说明你的系统不支持这几个模式. 可以尝试其他模式.


        """
        return self.__dm.GetLastError()

    def GetPath(self):
        """
        获取全局路径.(可用于调试)
        """
        return self.__dm.GetPath()

    # 获取版本信息
    def Ver(self) -> str:
        """
        简介
            返回当前插件版本号
        返回值:
            str: 当前插件的版本描述字符串.
        """
        return self.__dm.Ver()

    # 注册大漠收费功能
    def Reg(self, key: str, code: str):
        """
        简介:
            非简单游平台使用，调用此函数来注册，从而使用插件的高级功能.推荐使用此函数.
        参数定义:
            reg_code 字符串: 注册码. (从大漠插件后台获取)

            ver_info 字符串: 版本附加信息. 可以在后台详细信息查看. 可以任意填写. 可留空. 长度不能超过10. 并且只能包含数字和字母以及小数点. 这个版本信息不是插件版本.
        返回值:
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
    def EnumWindow(self, parent, title, class_name, filter=2) -> list[int]:
        """
        简介:
            根据指定条件,枚举系统中符合条件的窗口,可以枚举到按键自带的无法枚举到的窗口
        参数定义:
            parent 整形数: 获得的窗口句柄是该窗口的子窗口的窗口句柄,取0时为获得桌面句柄

            title 字符串: 窗口标题. 此参数是模糊匹配.

            class_name 字符串: 窗口类名. 此参数是模糊匹配.

            filter 整形数: 取值定义如下

            1 : 匹配窗口标题,参数title有效

            2 : 匹配窗口类名,参数class_name有效.

            4 : 只匹配指定父窗口的第一层孩子窗口

            8 : 匹配所有者窗口为0的窗口,即顶级窗口

            16 : 匹配可见的窗口

            32 : 匹配出的窗口按照窗口打开顺序依次排列 <收费功能，具体详情点击查看>

            这些值可以相加,比如4+8+16就是类似于任务管理器中的窗口列表
        返回值:
            list[int]. 返回窗口句柄列表.
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
        简介:
            获取顶层活动窗口,可以获取到按键自带插件无法获取到的句柄
        返回值:
            int: 返回整型表示的窗口句柄
        """
        return self.__dm.GetForegroundWindow()

    def SetWindowState(self, hwnd, flag):
        """
        简介:
            设置窗口的状态
        参数定义:
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
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.SetWindowState(hwnd, flag)

    def FindWindowEx(self, parent: int, class_name: str, title='') -> int:
        """
        简介:
            查找符合类名或者标题名的顶层可见窗口,如果指定了parent,则在parent的第一层子窗口中查找.
        参数定义:
            parent 整形数: 父窗口句柄，如果为空，则匹配所有顶层窗口

            class 字符串: 窗口类名，如果为空，则匹配所有. 这里的匹配是模糊匹配.

            title 字符串: 窗口标题,如果为空，则匹配所有. 这里的匹配是模糊匹配.
        返回值:
            int: 整形数表示的窗口句柄，没找到返回0
        """
        return self.__dm.findWindowEx(parent, class_name, title)

    # 设置窗口尺寸
    def SetWindowSize(self, hwnd, width, height) -> int:
        """
        简介:
            设置窗口的大小
        参数定义:
            hwnd 整形数: 指定的窗口句柄

            width 整形数: 宽度

            height 整形数: 高度
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.SetWindowSize(hwnd, width, height)

    def GetWindowTitle(self, hwnd) -> str:
        """
        简介:
            获取窗口的标题
        参数定义:
            hwnd 整形数: 指定的窗口句柄
        返回值:
            str: 窗口的标题
        """
        return self.__dm.GetWindowTitle(hwnd)

    # 后台绑定窗口
    def BindWindowEx(self,
                     hwnd,
                     display='dx2',
                     mouse='windows',
                     keypad='windows',
                     public='dx',
                     mode=101) -> int:
        """
        简介:
            绑定指定的窗口,并指定这个窗口的屏幕颜色获取方式,鼠标仿真模式,键盘仿真模式 高级用户使用.
        返回值:
            0: 失败
            1: 成功
        注意:
            绑定dx会比较耗时间,请不要频繁调用此函数.\n
            如果绑定的是dx,要注意不可连续操作dx,中间至少加个10ms的延时,否则可能会导致操作失败.比如绑定图色DX,那么不要连续取色等,键鼠也是一样.\n
            有些窗口绑定之后必须加一定的延时,否则后台也无效.一般1秒到2秒的延时就足够.
        """
        return self.__dm.BindWindowEx(hwnd, display, mouse, keypad, public,
                                      mode)

    # 解绑窗口
    def UnBindWindow(self) -> int:
        """
        简介:
            解除绑定窗口,并释放系统资源.一般在OnScriptExit调用
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.UnBindWindow()

    def EnableFakeActive(self, enabled=1):
        """
        简介:
            设置是否开启后台假激活功能. 默认是关闭. 一般用不到. 除非有人有特殊需求.
        参数定义:
            enable 整形数: 0 关闭 1 开启
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.EnableFakeActive(enabled)

    # ----------------------------文字识别----------------------------------
    def UseDict(self, index) -> int:
        """
        简介:
            表示使用哪个字库文件进行识别(index范围:0-9) 设置之后，永久生效，除非再次设定
        参数定义:
            index 整形数:字库编号(0-9)
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.UseDict(index)

    def SetDict(self, index, file) -> int:
        """
        简介:
            设置字库文件
        参数定义:
            index 整形数:字库的序号,取值为0-9,目前最多支持10个字库

            file 字符串:字库文件名
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.SetDict(index, file)

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
        """
        简介:
            在屏幕范围(x1,y1,x2,y2)内,查找string(可以是任意个字符串的组合),并返回符合color_format的坐标位置,相似度sim同Ocr接口描述.
        """
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
        """
        简介:
            识别屏幕范围(x1,y1,x2,y2)内符合color_format的字符串,并且相似度为sim,sim取值范围(0.1-1.0),

            这个值越大越精确,越大速度越快,越小速度越慢,请斟酌使用!
        """
        ret = self.__dm.Ocr(x1, y1, x2, y2, color_format, sim)
        if not ret:
            return None
        return ret

    # ----------------------------图色----------------------------------

    def CapturePng(self, x1, y1, x2, y2, file) -> int:
        """
        简介:
            同Capture函数，只是保存的格式为PNG.
        参数定义:
            x1 整形数:区域的左上X坐标
            y1 整形数:区域的左上Y坐标
            x2 整形数:区域的右下X坐标
            y2 整形数:区域的右下Y坐标
            file 字符串:保存的文件名,保存的地方一般为SetPath中设置的目录 当然这里也可以指定全路径名.
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.CapturePng(x1, y1, x2, y2, file)

    def EnableDisplayDebug(self, enabled=0):
        """
        简介:
            开启图色调试模式，此模式会稍许降低图色和文字识别的速度.默认不开启.
        参数定义:
            enabled 整形数: 0:关闭调试模式,1:开启调试模式
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.EnableDisplayDebug(enabled)

    def CapturePre(self, file) -> int:
        """
        简介:
            抓取上次操作的图色区域，保存为file(24位位图)
        参数定义:
            file 字符串: 保存的图片文件名
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.CapturePre(file)

    def CmpColor(self, x, y, color, sim) -> int:
        """
        简介:
            比较指定坐标点(x,y)的颜色
        参数定义:
            x 整形数: X坐标

            y 整形数: Y坐标

            color 字符串: 颜色字符串,可以支持偏色,多色,例如 "ffffff-202020|000000-000000" 这个表示白色偏色为202020,和黑色偏色为000000.颜色最多支持10种颜色组合. 注意，这里只支持RGB颜色.

            sim 双精度浮点数: 相似度(0.1-1.0)
        返回值:
            0: 匹配
            1: 不匹配
        """
        return self.__dm.CmpColor(x, y, color, sim)

    def FindColor(self,
                  x1,
                  y1,
                  x2,
                  y2,
                  color,
                  sim=1.0,
                  dir=FindDir.LeftToRightAndTopToBottom):
        """
        简介:
            查找指定区域内的颜色,颜色格式"RRGGBB-DRDGDB",注意,和按键的颜色格式相反
        """
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
        """
        简介:
            获取屏幕的高度.
        参数定义:
            width 整形数: 屏幕的宽度
            height 整形数: 屏幕的高度
            depth 整形数: 屏幕的色深度.(16或者32等)
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.SetScreen(width, height, depth)

    def GetScreenHeight(self, ) -> int:
        """
        简介:
            获取屏幕的高度.
        返回值:
            int 返回屏幕的高度
        """
        return self.__dm.GetScreenHeight()

    def GetScreenWidth(self, ) -> int:
        """
        简介:
            获取屏幕的宽度.
        返回值:
            int 返回屏幕的宽度
        """
        return self.__dm.GetScreenWidth()

    def GetScreenDepth(self, ) -> int:
        """
        简介:
            获取屏幕的色深.
        返回值:
            int 返回系统颜色深度.(16或者32等)
        """
        return self.__dm.GetScreenDepth()

    def GetDiskSerial(self, ) -> str:
        """
        简介:
            获取本机的硬盘序列号.支持ide scsi硬盘. 要求调用进程必须有管理员权限. 否则返回空串.
        返回值:
            str 字符串表达的硬盘序列号
        """
        return self.__dm.GetDiskSerial()

    def GetMachineCode(self, ) -> str:
        """
        简介:
            获取本机的机器码.(带网卡). 此机器码用于插件网站后台. 要求调用进程必须有管理员权限. 否则返回空串.
        返回值:
            str 字符串表达的机器机器码
        """
        return self.__dm.GetMachineCode()

    def GetMachineCodeNoMac(self, ) -> int:
        """
        简介:
            获取本机的机器码.(不带网卡) 要求调用进程必须有管理员权限. 否则返回空串.
        返回值:
            str 字符串表达的机器机器码
        """
        return self.__dm.GetMachineCodeNoMac()

    def GetOsType(self, ) -> int:
        """
        简介:
            得到操作系统的类型
        返回值:
            0 : win95/98/me/nt4.0

            1 : xp/2000

            2 : 2003

            3 : win7/vista/2008
        """
        return self.__dm.GetOsType()

    def CheckFontSmooth(self, ) -> int:
        """
        简介:
            检测当前系统是否有开启屏幕字体平滑
        返回值:
           0 : 系统没开启平滑字体.

           1 : 系统有开启平滑字体.
        """
        return self.__dm.CheckFontSmooth()

    def DisableFontSmooth(self, ) -> int:
        """
        简介:
            关闭当前系统屏幕字体平滑.同时关闭系统的ClearType功能.
        返回值:
            0: 失败
            1: 成功
        """
        return self.__dm.DisableFontSmooth()
