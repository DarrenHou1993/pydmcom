from pydmcom import DM

DM.setDllPath('D:/Program Files/DMSoft/dmsoft.dll')
DM.setRegisterDllPath('D:/Program Files/DMSoft/dmreg.dll')
dm = DM()
print(dm.GetScreenWidth())


regRet = dm.Reg('6008885359e4a0926d85f7483f16735d9fb71b', 'JINXIN')
print(f"注册结果: {regRet}")
# Chrome_RenderWidgetHostHWND
# 348,31,380,47,宽高(32,16) 清除
# 高度80后可以绘画

bindRet = dm.BindWindowEx(131320)

print(f"绑定结果: {bindRet}")
