TAG = '[Dispatch for MAC]'


class Dispatch(object):

    def __init__(self, name):
        print(f'{TAG} 对接COM组件 {name}')

    def SetDict(self, index, path):
        pass

    def SetPath(self, path):
        pass

    def Ver(self):
        return '7.2'

    def GetID(self):
        return 123321

    def GetBasePath(self):
        return '/usr/base'

    def Reg(self, arg1, arg2):
        return 1

    def GetScreenDepth(self):
        return 32

    def GetScreenWidth(self):

        return 1920

    def GetScreenHeight(self):
        return 1080

    def BindWindowEx(self, hwnd,
                     display='dx2',
                     mouse='windows',
                     keypad='windows',
                     public='dx',
                     mode=101):
        return 1

    def EnableDisplayDebug(self):
        return 1


class VARIANT():
    pass
