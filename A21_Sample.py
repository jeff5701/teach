"""
Created on Fri Jun 18 08:23:29 2021

@author: Jeff Huang
因為資料改架構更改，所以程式改寫
"""

import sys
import re 
from os import listdir
from os.path import isfile, isdir, join
import pandas as pd
import pyodbc
from datetime import datetime



server = 'My Server'
db='MyDataBase1'
db1 = 'MyDataBase2'
user = 'MyUserId'
password = 'MyPassWord'


ce = 1091 #學年學期比較用
table = 'A21_姊妹校資料'
table1="A21_姊妹校資料_簽約狀態"
table2 = 'A21_姊妹校資料_Temp'

#cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+db +';Trusted_Connection=yes;')
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+db +';UID='+user+';PWD='+ password)
cnxn1 = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+db1 +';UID='+user+';PWD='+ password)
cnxn2 = pyodbc.connect('DRIVER={SQL Server};SERVER='+server+';DATABASE='+db1 +';UID='+user+';PWD='+ password)


cursor = cnxn.cursor()
cursor1 = cnxn1.cursor()
cursor2 = cnxn2.cursor()



filepath = r"D:\BI\file"
files = r'姊妹校.xlsx'





#12：處理續約學校processupdate(M_Seq_No)
def processupdate(M_Seq_No):
    Compare = 1091 #ToDo 每次都要改
    yms_year=Compare[0:3]
    yms_sms=Compare[-1:4]
    
    while int(str(yms_year) + str(yms_sms)) < Compare:                            
        tmp_Year = (int(str(yms_year) + str(yms_sms)))
        
        cmd = "insert into [IR_Analysis].[dbo].[A21_姊妹校資料_簽約狀態] ([M_Seq_No],[yms_year],[yms_sms],[Status],[tempterm]) values({},{},{},N'{}',N'{}')".format(
            M_Seq_No,yms_year,yms_sms,"續約",tmp_Year)        
        try:
            cursor1.execute(cmd )
            cnxn1.commit()
        except Exception as e:
            print(e)
            print(M_Seq_No)
            
        #cursor1.execute(cmd )     
        #cnxn1.commit()
        
        if yms_sms ==2 :
            yms_sms = 1
            yms_year +=1
        else:
            yms_sms = 2
        
        #print (int(str(yms_year) + str(yms_sms)))


    
            




#11：取得M_Seq_No：
def GetMSeqNo(datas):
    #todo
    #cmd="INSERT INTO  [IR_Analysis].[dbo].[A21_姊妹校資料] (country,School,Region) VALUES(N'{}',N'{}',N'{}') ; select scope_identity() as id ".format(datas[0],datas[1],datas[2])
    cmd="SELECT [M_Seq_No] FROM A21_姊妹校資料 WHERE country =N'{}' AND school=N'{}' AND Region=N'{}'".format(datas[0],datas[1],datas[2])
    cursor.execute(cmd)
    row = cursor.fetchone()
    if row:        
        MSeq_No=str(row[0])
        #if (MSeq_No=="436"):
        #    print(school)
        #    print(tmp_year)
        #    print()
        #MSeq_No=""
        #print(MSeq_No)
    else:
        #print("Master 查無資料")
        MSeq_No="0"
        print(cmd)
        #先取消20210609
        #ErrOutPut("Master 查無資料；可能交流的學年期不在合約的學年期範圍內",ErrData)        
            
    #print(book_id[0])
    cursor.commit()
    return(MSeq_No)
    



#10:將錯誤資料存入Error table
def Erroutput(datas,descp):
    origen=(datas[0])
    country=(datas[1])
    school=(datas[2])
    #if origen =='歐洲':
    #    print(school)
        
    cmd="Insert Into A21_姊妹校資料_Error ([country],[school],[region],[descp]) VALUES ('{}', N'{}','{}','{}')".format(country,school.replace("'","''"),origen,descp.replace("'","''"))     
    cursor1.execute(cmd )     
    cnxn1.commit()





#9:Process to inser into A21_姊妹校資料_簽約狀態
"""
1.此功能是將每學期的EXCEL檔存入A21_姊妹校資料_Temp，再與主檔A21_姊妹校進行資料比對
2.狀態只存首次簽約
"""
def  processinseert(term,M_Seq_No):
    
    yms_year=term[0]
    yms_sms=term[1]
    
    #Compare = GetCompare()
    #這個要改
    Compare = 1091
    #Compare =getnowtern() #取得要匯到那個學期為止
              
    #if school=="北京外國語大學  Beijing Foreign Studies University":
    #    print(yms_year)
    #    print(yms_sms)
        
        
        
        
    
    """
    1.修改
    2.原程式：while int(str(yms_year) + str(yms_sms)) < Compare :
    """
    i=0
    while int(str(yms_year) + str(yms_sms)) < Compare :
        if i == 0:
            status="首次簽約"
        else:
            status="續約"            
            
        i +=1
        #print(i)
        
            
        tmp_Year = (int(str(yms_year) + str(yms_sms)))
        
        cmd = "insert into [IR_Analysis].[dbo].[A21_姊妹校資料_簽約狀態] ([M_Seq_No],[yms_year],[yms_sms],[Status],[tempterm]) values({},{},{},N'{}',N'{}')".format(
            M_Seq_No,yms_year,yms_sms,status,tmp_Year)
        
        try:
            cursor1.execute(cmd )
            cnxn1.commit()
        except Exception as e:
            print(e)
            print(M_Seq_No)
            
        #cursor1.execute(cmd )     
        #cnxn1.commit()
        
        if yms_sms ==2 :
            yms_sms = 1
            yms_year +=1
        else:
            yms_sms = 2
        
        #print (int(str(yms_year) + str(yms_sms)))








#8Insert New School

def InsertNewSchool(datas):
    #todo
    #cmd="INSERT INTO  [IR_Analysis].[dbo].[A21_姊妹校資料] (country,School,Region) VALUES(N'{}',N'{}',N'{}') ; select scope_identity() as id ".format(datas[0],datas[1],datas[2])
    cmd="INSERT INTO  [IR_Analysis].[dbo].[A21_姊妹校資料] (country,School,Region) VALUES(N'{}',N'{}',N'{}') ".format(datas[0],datas[1],datas[2]) 
    #cmd="INSERT INTO  [IR_Analysis].[dbo].[A21_姊妹校資料] (country,School,Region) VALUES(N'{}',N'{}',N'{}') ; SELECT @@IDENTITY AS book_id; ".format(datas[0],datas[1],datas[2])
    #print(cmd)
    cursor.execute(cmd )
    cursor.execute('select scope_identity() as fred')
    book_id = cursor.fetchone()
    #print(book_id[0])
    cursor.commit()
    return(book_id[0])


















#7:處理學年學期
def processSYearTerm(string,fun):
    
    #print(string)
    #print(fun)
    
    tmpTerm=string[0].split('/')
    m=0
    y=0
    y=int(int(tmpTerm[0]))
    m=(int(tmpTerm[1]))
    #print(m)
    
    if (m > 8):
        tempTerm=1
        tempYear = y -1911
    elif (m == 1): #1月
        tempTerm=1
        tempYear = y -1912
    else:
        tempTerm=2
        tempYear = y -1912
    
    return [tempYear,tempTerm]
    

            





#6:取得首次簽約日期
def getSignDate(string):
    #temp_signdate = re.findall(r'\d{4}\/\d{2}',string) #原始程式         
    temp_signdate = re.findall(r'\d+\/\d+',string)          
    
    if len(temp_signdate) > 0:
        first_signdate = processSYearTerm(temp_signdate,"getSignDate")
    else:
        first_signdate=''
        #print("getSignDate " +' ' + 'first_signdate len:' + str(len(first_signdate)))
        
    
    return first_signdate












"""
0:學校
1:國家
2:洲別
"""

#5檢核學校是否已存在
def ChkExistence(existence):
    #ToDoSomething
    
    try:
        
        cmd="SELECT DISTINCT [country] ,[school] ,[region] FROM [IR_Analysis].[dbo].[A21_姊妹校資料] WHERE school = N'{}' AND country = N'{}' AND region =N'{}'".format(existence[0],existence[1],existence[2])
        cursor.execute(cmd)
        row=cursor.fetchone()
        #print('execue')
        if row:
            exist=True           
        else:    
            #表示新學校
            exist=False            
        return exist
    
    except Exception as e:
        print(e)
        print('ChkExistence Error')


   


#4:保留中文字移除英文字
def Remove_Eng(string):
    try:
        result = re.sub(r'[^\u4e00-\u9fa5]','',string)
    
        return result
    except:
        return string    
    
    
       


#3:處理Excel主要Function
def Migration(file):
    count =0
    #亞洲地區
    data1 = pd.read_excel(file ,sheet_name = 0,skiprows = 3 , header = None )
    data1 = data1.rename(columns = {0:"國家",1:"學校",6:"簽約資料",13:"學生出訪",
                                    14:"學生來訪",15:"教師出訪",16:"教師來訪",17:'其他交流'})
    data1 = data1[["國家","學校","簽約資料","學生出訪","學生來訪","教師出訪","教師來訪","其他交流"]]
    for index, row in data1.iterrows():
        country=Remove_Eng(row['國家']) #移除英文字       
        school = str(row['學校'].replace('\n',' ' ) )
        sign2= str(row['簽約資料'])
        sign = str(row['簽約資料']).split('\n')
        so = str(row['學生出訪']).split('\n')
        si = str(row['學生來訪']).split('\n')
        to= str(row['教師出訪']).split('\n')
        ti= str(row['教師來訪']).split('\n')
        oh= str(row['其他交流']).split('\n')
                                        
        #檢核學校是否存在
        #datas=[school,country,"亞洲"]
        datas=[country,school.replace("'","''"),"亞洲"]
        existence=ChkExistence(datas)
        if existence :
            #已簽約學校
            #ToDoSomthing
            #print('old')            
            M_Seq_No=GetMSeqNo(datas)
            processupdate(M_Seq_No)
        else:
            #新簽約學校
            #ToDoSomthing
            #print('new')
            #新增學校
            #InsertNewSchool(datas)
            M_Seq_No=InsertNewSchool(datas)
            #print(M_Seq_No)                        
            strtmp = sign[len(sign)-1]
            signdate = getSignDate(strtmp.replace(".","/"))
            #print(signdate)            
            if (len(signdate) > 0):
                processinseert(signdate,M_Seq_No)
                datas = ['亞洲',country,school]
                
            else:
                datas = ['亞洲',country,school]
                descp="無法取得簽約日期"
                Erroutput(datas,descp)
            
                
        
            

        #不能用首次簽當成關鍵字,只能從串列最大值取後首次簽約資料                
        #strtmp = sign[len(sign)-1]
        
        #debug用
        #if school =="德山高專 National Institute of Technology Tokuyama College":
        #    print(country)
        #    signdate = getSignDate(strtmp.replace(".","/")) 
        #    break
 

       
        #取得簽約的學年與學期串列
        #print("school:" + school + ' ' + 'strtmp:' + strtmp)
        #signdate = getSignDate(strtmp.replace(".","/"))
        #print(signdate)
        
       
        #if (len(signdate) > 0):
        #    processinseert(signdate,'亞洲',country,school)
        #    datas = ['亞洲',country,school]
        #    
        #else:
        #    datas = ['亞洲',country,school]
        #    descp="無法取得簽約日期"
        #    Erroutput(datas,descp)
        #    #print("Error")
           
    count =0
    #歐洲地區
    data1 = pd.read_excel(file ,sheet_name = 1,skiprows = 3 , header = None )
    data1 = data1.rename(columns = {0:"國家",2:"學校",6:"簽約資料",13:"學生出訪",
                                    14:"學生來訪",15:"教師出訪",16:"教師來訪",17:'其他交流'})
    data1 = data1[["國家","學校","簽約資料","學生出訪","學生來訪","教師出訪","教師來訪","其他交流"]]
    for index, row in data1.iterrows():
        country=Remove_Eng(row['國家'])        
        school = str(row['學校'].replace('\n',' ' ) )
        sign2= str(row['簽約資料'])
        sign = str(row['簽約資料']).split('\n')
        so = str(row['學生出訪']).split('\n')
        si = str(row['學生來訪']).split('\n')
        to= str(row['教師出訪']).split('\n')
        ti= str(row['教師來訪']).split('\n')
        oh= str(row['其他交流']).split('\n')
        
        
        
        
        
        #檢核學校是否存在
        #datas=[school,country,"亞洲"]
        datas=[country,school.replace("'","''"),"歐洲"]
        existence=ChkExistence(datas)
        if existence :
            #已簽約學校
            #ToDoSomthing
            #print('old')            
            M_Seq_No=GetMSeqNo(datas)
            processupdate(M_Seq_No)
        else:
            #新簽約學校
            #ToDoSomthing
            #print('new')
            #新增學校
            #InsertNewSchool(datas)
            M_Seq_No=InsertNewSchool(datas)
            #print(M_Seq_No)                        
            strtmp = sign[len(sign)-1]
            signdate = getSignDate(strtmp.replace(".","/"))
            #print(signdate)            
            if (len(signdate) > 0):
                processinseert(signdate,M_Seq_No)
                datas = ['歐洲',country,school]
                
            else:
                datas = ['歐洲',country,school]
                descp="無法取得簽約日期"
                Erroutput(datas,descp)
        
        #不能用首次簽當成關鍵字,只能從串列最大值取後首次簽約資料        
        
        #strtmp = sign[len(sign)-1]
        
        #if school=="聖安東尼奧大學Universidad Catolica San Antonio de Murcia":
        #    print("oh:" + str(len(oh)))
        #    print("oh:" + str(len(oh[0])))
        #    print("strtmp:" + strtmp)
        #print(school)
             

        
        
        #取得簽約的學年與學期串列
        #print("school:" + school + ' ' + 'strtmp:' + strtmp)
        #signdate = getSignDate(strtmp.replace(".","/"))        
        #if (len(signdate) > 0):
        #    processinseert(signdate,'歐洲',country,school)
        #    datas = ['歐洲',country,school]
        #   
        #    
        #else:
        #    datas = ['歐洲',country,school]
        #    descp="無法取得簽約日期"
        #    Erroutput(datas,descp)
        #    #print("Error")
            
            
            
    count =0
    #美洲地區
    data1 = pd.read_excel(file ,sheet_name = 2,skiprows = 3 , header = None )
    data1 = data1.rename(columns = {0:"國家",2:"學校",6:"簽約資料",13:"學生出訪",
                                    14:"學生來訪",15:"教師出訪",16:"教師來訪",17:'其他交流'})
    data1 = data1[["國家","學校","簽約資料","學生出訪","學生來訪","教師出訪","教師來訪","其他交流"]]
    for index, row in data1.iterrows():
        country=Remove_Eng(row['國家'])        
        school = str(row['學校'].replace('\n',' ' ) )
        sign2= str(row['簽約資料'])
        sign = str(row['簽約資料']).split('\n')
        so = str(row['學生出訪']).split('\n')
        si = str(row['學生來訪']).split('\n')
        to= str(row['教師出訪']).split('\n')
        ti= str(row['教師來訪']).split('\n')
        oh= str(row['其他交流']).split('\n')
        
        
        
        
        
        #檢核學校是否存在
        #datas=[school,country,"亞洲"]
        datas=[country,school.replace("'","''"),"美洲"]
        existence=ChkExistence(datas)
        if existence :
            #已簽約學校
            #ToDoSomthing
            #print('old')            
            M_Seq_No=GetMSeqNo(datas)
            processupdate(M_Seq_No)
        else:
            #新簽約學校
            #ToDoSomthing
            #print('new')
            #新增學校
            #InsertNewSchool(datas)
            M_Seq_No=InsertNewSchool(datas)
            #print(M_Seq_No)                        
            strtmp = sign[len(sign)-1]
            signdate = getSignDate(strtmp.replace(".","/"))
            #print(signdate)            
            if (len(signdate) > 0):
                processinseert(signdate,M_Seq_No)
                datas = ['美洲',country,school]
                
            else:
                datas = ['美洲',country,school]
                descp="無法取得簽約日期"
                Erroutput(datas,descp)
        
        #不能用首次簽當成關鍵字,只能從串列最大值取後首次簽約資料        
        
        #strtmp = sign[len(sign)-1]
        
        #if school=="聖安東尼奧大學Universidad Catolica San Antonio de Murcia":
        #    print("oh:" + str(len(oh)))
        #    print("oh:" + str(len(oh[0])))
        #    print("strtmp:" + strtmp)
        #print(school)
             

        
        
        #取得簽約的學年與學期串列
        #print("school:" + school + ' ' + 'strtmp:' + strtmp)
        #signdate = getSignDate(strtmp.replace(".","/"))        
        #if (len(signdate) > 0):
        #    processinseert(signdate,'美洲',country,school)
        #    datas = ['美洲',country,school]
       # 
       #     
        #else:
        #    datas = ['美洲',country,school]
        #    descp="無法取得簽約日期"
        #    Erroutput(datas,descp)
        #    #print("Error")
            
            
            
    count =0
    #非洲地區
    data1 = pd.read_excel(file ,sheet_name = 3,skiprows = 3 , header = None )
    data1 = data1.rename(columns = {0:"國家",2:"學校",5:"簽約資料",11:"學生出訪",
                                    12:"學生來訪",13:"教師出訪",14:"教師來訪",15:'其他交流'})
    data1 = data1[["國家","學校","簽約資料","學生出訪","學生來訪","教師出訪","教師來訪","其他交流"]]
    for index, row in data1.iterrows():
        country=Remove_Eng(row['國家'])        
        school = str(row['學校'].replace('\n',' ' ) )
        sign2= str(row['簽約資料'])
        sign = str(row['簽約資料']).split('\n')
        so = str(row['學生出訪']).split('\n')
        si = str(row['學生來訪']).split('\n')
        to= str(row['教師出訪']).split('\n')
        ti= str(row['教師來訪']).split('\n')
        oh= str(row['其他交流']).split('\n')
                        
        
        
        #檢核學校是否存在
        #datas=[school,country,"亞洲"]
        datas=[country,school.replace("'","''"),"非洲"]
        existence=ChkExistence(datas)
        if existence :
            #已簽約學校
            #ToDoSomthing
            #print('old')            
            M_Seq_No=GetMSeqNo(datas)
            processupdate(M_Seq_No)
        else:
            #新簽約學校
            #ToDoSomthing
            #print('new')
            #新增學校
            #InsertNewSchool(datas)
            M_Seq_No=InsertNewSchool(datas)
            #print(M_Seq_No)                        
            strtmp = sign[len(sign)-1]
            signdate = getSignDate(strtmp.replace(".","/"))
            #print(signdate)            
            if (len(signdate) > 0):
                processinseert(signdate,M_Seq_No)
                datas = ['非洲',country,school]
                
            else:
                datas = ['非洲',country,school]
                descp="無法取得簽約日期"
                Erroutput(datas,descp)
        
        
        #print(sign)
       
        #不能用首次簽當成關鍵字,只能從串列最大值取後首次簽約資料        
        
        #strtmp = sign[len(sign)-1]
        
        #if school=="聖安東尼奧大學Universidad Catolica San Antonio de Murcia":
        #    print("oh:" + str(len(oh)))
        #    print("oh:" + str(len(oh[0])))
        #    print("strtmp:" + strtmp)
        #print(school)
             

        
        
        #取得簽約的學年與學期串列
        #print("school:" + school + ' ' + 'strtmp:' + strtmp)
        #signdate = getSignDate(strtmp.replace(".","/"))        
        
        #if (len(signdate) > 0):
        #    processinseert(signdate,'非洲',country,school)
        #    datas = ['非洲',country,school]                   
        #else:
        #    datas = ['非洲',country,school]
        #    descp="無法取得簽約日期"
        #    Erroutput(datas,descp)
        #    #print("Error")
            
            
            
    count =0
    #大洋洲地區
    data1 = pd.read_excel(file ,sheet_name = 4,skiprows = 3 , header = None )
    data1 = data1.rename(columns = {0:"國家",2:"學校",6:"簽約資料",12:"學生出訪",
                                    13:"學生來訪",14:"教師出訪",15:"教師來訪",16:'其他交流'})
    data1 = data1[["國家","學校","簽約資料","學生出訪","學生來訪","教師出訪","教師來訪","其他交流"]]
    for index, row in data1.iterrows():
        country=Remove_Eng(row['國家'])        
        school = str(row['學校'].replace('\n',' ' ) )
        sign2= str(row['簽約資料'])
        sign = str(row['簽約資料']).split('\n')
        so = str(row['學生出訪']).split('\n')
        si = str(row['學生來訪']).split('\n')
        to= str(row['教師出訪']).split('\n')
        ti= str(row['教師來訪']).split('\n')
        oh= str(row['其他交流']).split('\n')
                        
        
        
        #檢核學校是否存在
        #datas=[school,country,"亞洲"]
        datas=[country,school.replace("'","''"),"大洋洲"]
        existence=ChkExistence(datas)
        if existence :
            #已簽約學校
            #ToDoSomthing
            #print('old')            
            M_Seq_No=GetMSeqNo(datas)
            processupdate(M_Seq_No)
        else:
            #新簽約學校
            #ToDoSomthing
            #print('new')
            #新增學校
            #InsertNewSchool(datas)
            M_Seq_No=InsertNewSchool(datas)
            #print(M_Seq_No)                        
            strtmp = sign[len(sign)-1]
            signdate = getSignDate(strtmp.replace(".","/"))
            #print(signdate)            
            if (len(signdate) > 0):
                processinseert(signdate,M_Seq_No)
                datas = ['大洋洲',country,school]
                
            else:
                datas = ['大洋洲',country,school]
                descp="無法取得簽約日期"
                Erroutput(datas,descp)
        
        
        #不能用首次簽當成關鍵字,只能從串列最大值取後首次簽約資料        
        
        #strtmp = sign[len(sign)-1]
        
        #if school=="聖安東尼奧大學Universidad Catolica San Antonio de Murcia":
        #    print("oh:" + str(len(oh)))
        #    print("oh:" + str(len(oh[0])))
        #    print("strtmp:" + strtmp)
        #print(school)
             

        
        
        #取得簽約的學年與學期串列
        #print("school:" + school + ' ' + 'strtmp:' + strtmp)
        #signdate = getSignDate(strtmp.replace(".","/"))        
        #if (len(signdate) > 0):
        #    processinseert(signdate,'大洋洲',country,school)
        #    datas = ['大洋洲',country,school]                   
        #else:
        #    datas = ['大洋洲',country,school]
        #    descp="無法取得簽約日期"
        #    Erroutput(datas,descp)
        #    #print("Error")              

     
 
        
    

    

    
    print('OK')
    




















#2:處理錯誤資料表
def Clear_ErrorTable():
    cmd ="truncate table A21_姊妹校資料_Error"    
    cursor1.execute(cmd)
    cnxn1.commit()


#1:清除A21_姊妹校資料_Temp資料
def Clear_Data():
    cmd = 'truncate table {}'.format(table2) 
    cursor1.execute(cmd )     
    cnxn1.commit()



#0:初始
if __name__ =='__main__':
    #files = listdir(filepath)
    #Migration(filepath + '\\' + files[0])
    Clear_Data() #清除暫存資料表 'A21_姊妹校資料_Temp'
    Clear_ErrorTable()
    Migration(filepath  +'\\' +files)
    #getrecordset()