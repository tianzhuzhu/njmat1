#输入为综调系统导出表与用户关系表
#输出为生成的表单
import datetime
import os
import openpyxl
import xlrd

import pandas as pd
def checkin():
    tomonth = datetime.datetime.now().month
    today = datetime.datetime.now().day
    while(True):
        mode=input("请输入模式：1 市政工程 2 整治工单")
        mode= "市政工程" if mode=='1' else "整治工单" if mode=='2' else ''
        print( mode+str(tomonth) + "-" + str(today) )
        file = input("今天是{},请输入需要归档清单原表格绝对路径（格式为{}mm-dd.xlsx）):".format(datetime.date.today().strftime(r"%y-%m-%d"),mode))
        while(file.find(str(tomonth)+"-"+str(today))<0 and file.find(mode)<0):
            file = input("今天是{},请输入需要归档清单原表格绝对路径（格式为{}mm-dd.xlsx）):".format(datetime.date.today().strftime(r"%y-%m-%d"),mode))
        data=pd.read_excel(file,sheet_name='data')
        matchup=input('工单归属关系表绝对路径（格式为*归属关系.xlsx）:')
        while(matchup.find('归属关系')<0):
            matchup = input('工单归属关系表绝对路径（格式为*归属关系.xlsx）:')
        matchup=pd.read_excel(matchup,sheet_name="部门")
        path=file[:file.rfind('\\')]
        return data,mode,matchup,path


def IsMoreThanSevenDay(x):
    x=str(x)
    if(x.find('剩余')>0):
        return False
    right=x.find('天')
    if(right==-1):
        return False
    left=x.find('：') if x.find('：')!=-1 else x.find(':')
    if(int(x[left + 1: right])>=7):
        return True
    return False


def pviotAndOutput(data, name,path,mode):
    count=data.value_counts('部门',sort=True)
    count=pd.DataFrame(count,columns=['汇总'])
    count.reindex()
    count.loc['总计','汇总']=count['汇总'].sum()

    if(not os.path.exists(path)):
        os.mkdir(path)
    if(mode=='整治工单'):
        path=os.path.join(path,str(datetime.datetime.now().month)+"-"+str(datetime.datetime.now().day)+mode+"在途汇总"+'.xlsx')
    elif(mode=='市政工程'):
        path = os.path.join(path, str(datetime.datetime.now().month) + "-" + str(
            datetime.datetime.now().day) + mode + '尚未归档.xlsx')

    if(name=='总清单'):
        count.to_excel(path,sheet_name=name+"汇总")
    else:
        with pd.ExcelWriter(path, mode='a', engine='openpyxl') as writer:
            count.to_excel(writer, sheet_name=name+"汇总表")
    with pd.ExcelWriter(path, mode='a', engine='openpyxl') as writer:
        data.to_excel(writer, sheet_name=name,index=False)


def findDepartmentBygroup(x):
    x=str(x)
    list=['溧水','六合','江宁','高淳','浦口']
    for i in list:
        print(i)
        if x.find(i)>=0:
            return i
    list=['维护岗','PON','光缆','设备']
    for i in list:
        if x.find(i)>=0:
            return '综维'
    list=['建设','有线接入']
    for i in list:
        if x.find(i)>=0:
            return '建设'
    list = ['秦淮', '鼓楼','雨花台','玄武','栖霞','化工园','建业']
    for i in list:
        if x.find(i) >= 0:
            return i
    if(x.find('雨花')>=0):
        return '雨花台'



def handleData(data, matchup,path,mode):
    matchup.rename(columns={"行标签":"处理组"},inplace=True)
    print(matchup)
    df=data['组/处理人'].str.split('/',expand=True)
    data['处理人'] = df[0]
    data['处理组']=df[1]
    print(data['处理组'])
    mergeData=pd.merge(left=data,right=matchup,on='处理组',how='left')
    columns=mergeData.columns.tolist()
    toFifth=['处理人','处理组','部门']
    for v in columns:
       if(v in toFifth):
           columns.remove(v)
           columns.insert(4,v)
    columns.remove('组/处理人')
    mergeData=mergeData.reindex(columns=columns)
    #根据处理组头匹配
    na=mergeData.loc[mergeData['部门'].isna()].apply(lambda x: findDepartmentBygroup(x['处理组']), axis=1)
    print(mergeData.loc[mergeData['部门'].isna()])
    print(' ---------自动填充--------')
    print(na)
    print(mergeData.loc[mergeData['部门'].isna()])
    mergeData.loc[mergeData['部门'].isna(),'部门']=na
    print(mergeData)
    # 需要丢弃 todo
#数据已经融合，接下来根据历时进行拆分
    NADATA=mergeData.loc[mergeData['部门'].isna()].copy()
    mergeData.dropna(subset=['部门'],inplace=True)
    mergeData.sort_values(by='截止时间',inplace=True,ascending=True)
    MoreThanSevenDay=mergeData.loc[mergeData['剩余历时'].map(lambda x:IsMoreThanSevenDay(x))==True]

    appendData=mergeData.append(MoreThanSevenDay)
    lessThanSevenData=appendData.drop_duplicates(keep=False)
    mergeData['排序']=mergeData['剩余历时'].map(lambda x:IsMoreThanSevenDay(x))
    mergeData.sort_values(by='排序',ascending=False)
    mergeData=mergeData.drop(columns='排序')
    path = os.path.join(path, 'result')
    pviotAndOutput(mergeData,'总清单',path,mode)
    pviotAndOutput(MoreThanSevenDay,'超时七天',path,mode)
    pviotAndOutput(lessThanSevenData,'未超时七天',path,mode)
    if(not NADATA.empty):
        NADATA.to_excel(os.path.join(path,mode+'未匹配到部门.xlsx'))

#完成拆分分别在 mergedata appendData less thansevendata中 ，mergedata加入一列

if __name__ == '__main__':
    data,mode,matchup,path= checkin()
    handleData(data,matchup,path,mode)