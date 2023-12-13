from pydmcom import DM
import os

DM.setDllPath('D:/Program Files/DMSoft/dmsoft.dll')
DM.setRegisterDllPath('D:/Program Files/DMSoft/dmreg.dll')
dm = DM()
print(dm.GetScreenWidth())


regRet = dm.Reg('6008885359e4a0926d85f7483f16735d9fb71b', 'JINXIN')
print(f"注册结果: {regRet}")


RESOURCE_PATH = os.path.abspath(os.path.join(__file__, '../test-resources'))

print(f"资源路径: {RESOURCE_PATH}")
dm.SetPath(RESOURCE_PATH)
dm.SetDict(0, r'fonts\font1.txt')
dm.UseDict(0)
windows = dm.FindWindowEx(None, 'Chrome_RenderWidgetHostHWND')
print(f"windows {windows}")

bindRet = dm.BindWindowEx(windows[0])
print(f"绑定结果: {bindRet}")
dm.SetWindowSize(windows[0], 800, 600)

dm.MoveTo(100, 100)
dm.LeftDown()
dm.MoveTo(100, 200)
dm.MoveTo(200, 200)
dm.MoveTo(200, 100)
dm.MoveTo(100, 100)
dm.LeftUp()

position = dm.FindStr(348, 31, 380, 47, '清除')
print(f"找字的坐标 {position}")
# 348,31,380,47,宽高(32,16) 清除
# 高度80后可以绘画
