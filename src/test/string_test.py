path = 'D:/Program Files (x86)/nfsd/RestartService.bat'
pos = path.rfind('/')
print(path[0: pos])
print(path[pos+1:])