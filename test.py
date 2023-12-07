from pydmcom import DM

DM.setDllPath('D:/Program Files/DMSoft/dmsoft.dll')
DM.setRegisterDllPath('D:/Program Files/DMSoft/dmreg.dll')
dm = DM()
print(dm.GetScreenWidth())


regRet = dm.Reg('6008885359e4a0926d85f7483f16735d9fb71b', 'JINXIN')
print(f"注册结果: {regRet}")

bindRet = dm.BindWindowEx(123321)

print(f"绑定结果: {bindRet}")
