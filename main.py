from pydmcom import DM

DM.setDllPath('D:/Program Files/DMSoft/dmsoft.dll')
DM.setRegisterDllPath('D:/Program Files/DMSoft/dmreg.dll')
dm = DM()
print(dm.GetScreenWidth())
