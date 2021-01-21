import os


def getCurrentDirPath(file):
    l = file.rfind('\\')
    path = file[:l]
    return path

def check(tomonth, file, name):
    l=file.rfind('\\')
    r=file.rfind('月')
    find =file.find(name)
    file=file[l+1:r]
    if(tomonth==-1 and find>0):
        return True
    elif(find>0 and ((int(file)-tomonth)<=1 or (int(file)-tomonth)==11)):
        return True
    else: return False

def getOutPutName(file):

    l = file.rfind('\\')
    r = file.rfind('月')
    path=file[:l]
    mo = file[l+1:r]
    path=os.path.join(path,'result')
    return str(mo),str(path)