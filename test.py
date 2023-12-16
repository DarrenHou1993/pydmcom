from pydmcom import DM
import time
import os
APP_DATA_PATH = os.path.join(
    os.path.expanduser('~'),
    'tiger_debug')
DLL_LIB_PATH = os.path.join(APP_DATA_PATH, 'libs')
RESOURCE_PATH = os.path.abspath(os.path.join(__file__, '../test-resources'))

print(f"资源路径: {RESOURCE_PATH}")
print(f"libs路径: {DLL_LIB_PATH}")

DM.setDllPath(os.path.join(DLL_LIB_PATH, 'dm.dll'))
DM.setRegisterDllPath(os.path.join(DLL_LIB_PATH, 'RegDll.dll'))
dm = DM()
regRet = dm.Reg('6008885359e4a0926d85f7483f16735d9fb71b', 'JINXIN')
print(f"注册结果: {regRet}")

print(f"当前屏幕的分辨率是 {dm.GetScreenWidth()}, {dm.GetScreenHeight()}")

dm.SetPath(RESOURCE_PATH)
set_dict_ret = dm.SetDict(0, r'fonts\font1.txt')
print(f"设置字库结果：{set_dict_ret}")
use_dict_ret = dm.UseDict(0)
print(f"使用字库结果：{use_dict_ret}")

windows = dm.EnumWindow(0,'WhiteBoard - Google Chrome',None,1)
child = dm.FindWindowEx(windows[0], 'Chrome_RenderWidgetHostHWND')
print(f"windows {windows} child {child}")

bindRet = dm.BindWindowEx(child)
print(f"绑定结果: {bindRet}")
dm.SetWindowSize(windows[0], 900, 800)

time.sleep(1)
dm.MoveTo(100, 100)
time.sleep(0.5)

dm.LeftDown()
dm.MoveTo(100, 200)
time.sleep(0.5)

dm.MoveTo(200, 200)
time.sleep(0.5)

dm.MoveTo(200, 100)
time.sleep(0.5)

dm.MoveTo(100, 100)
time.sleep(0.5)

dm.LeftUp()

position = dm.FindStr(304,26,350,49, '清除','b9b9b9-222222')
print(f"找字的坐标 {position}")


image_position = dm.FindPic(309,4,344,30,"images/clear.bmp")
print(f"找图的坐标 {image_position}")

filename = os.path.join(RESOURCE_PATH,"capture",'pre.bmp')
print(f"file path {filename}")
path = dm.CapturePre(filename)
print(f"capture path ret {path}")
# 348,31,380,47,宽高(32,16) 清除
# 高度80后可以绘画
