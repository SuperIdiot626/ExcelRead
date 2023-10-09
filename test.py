# -*- coding: utf-8 -*-
import datetime
import openpyxl
import csv
from openpyxl.styles import PatternFill     # 导入填充模块
from openpyxl.styles import Font

encode_type_alipay='gb18030'
encode_type_wechat='UTF-8'



except_alipay=["等待付款","交易关闭"]

color_gren="92D050"
color_grey="E7E6E6"
color_yellow="FFFF00"
color_red="FF0000"
color_orange="FFC000"
color_blue="00B0F0"
color_none="FFFFFF"


def read_data(file_path,input_type):
    if(input_type=="alipay"):               #配置输入相关属性
        encoding_type=encode_type_alipay
        time_format=r"%Y-%m-%d %H:%M:%S"
    if(input_type=="wechat"):
        encoding_type=encode_type_wechat
        time_format=r"%Y-%m-%d %H:%M:%S"

    data=[]
    with open(file_path, 'r',encoding=encoding_type,) as file:  #读取数据写入data
        reader = csv.reader(file)
        for row in reader:
            data.append(row)    
    
    for i in range(len(data)):                                  #找到真正需要的部分
        if data[i]==[]:
            continue
        if data[i][0]=="交易时间":
            break
    data=data[:i:-1]                                            #然后切除并从按日期从前往后排列

    if(input_type=="alipay"):
        i=0
        while(i<len(data)):
            data[i][0]= datetime.datetime.strptime(data[i][0], time_format) #修改为统一的时间格式
            if(data[i][8] in except_alipay or float(data[i][6])==0):        #交易状态以及交易金额判断
                del data[i]
                continue
            if(data[i][4][-4:]=="收益发放"):    #处理余额宝利息
                data[i][2]=""
                data[i][4]="余额宝利息"
                data[i][5]="收入"
            if(data[i][10][0]=="T"):            #处理淘宝商品
                data[i][2]=""
            if(data[i][2]=="美团"):             #处理美团商品
                data[i][2]=""
            if(data[i][4]=="余额宝-转出到银行卡"):
                data[i][4]="余额宝转出到"+data[i][2]
                data[i][5]="资金周转"
                data[i][2]=""
            if(data[i][4]=="余额宝-单次转入" ):
                data[i][5]="资金周转"
                data[i][2]=""

            data[i][4]=data[i][2]+data[i][4]    #讲其他无需进一步处理的交易记录的交易方和商品对象合在一起
            data[i][2]=""

            temp=[data[i][0],data[i][4],data[i][5],data[i][6],"支付宝",]
            data[i]=temp
            i+=1

    if(input_type=="wechat"):
        i=0
        while(i<len(data)):
            try:
                data[i][0]= datetime.datetime.strptime(data[i][0], time_format) #修改为统一的时间格式
            except:
                continue
            if(data[i][1]=="群收款" and data[i][4]=="收入"):
                data[i][2]="群收款_"+data[i][2]
                data[i][3]=""
                data[i][4]="群收款收入"

            if(data[i][1]=="微信红包"):
                data[i][2]="红包来自_"+data[i][2]
                data[i][3]=""
            
            if(data[i][1]=="微信红包（单发）"):
                data[i][3]=""

            if(data[i][1]=="微信红包-退款"):
                data[i][2]=""
                data[i][3]="微信红包-退款"

            if(data[i][1]=="商户消费"):
                data[i][3]=""
            
            if(data[i][1]=="零钱充值"):
                data[i][3]=data[i][1]+data[i][2][:-6]
                data[i][4]="资金周转"
                data[i][2]=""


            if(data[i][1]=="零钱提现"):
                data[i][3]=data[i][1]+data[i][2][:-6]
                data[i][2]=""
                data[i][4]="资金周转"
                service_fee=data[i+1][10][4:]
                data[i][5]="￥"+str(float(data[i][5][1:])-float(data[i][10][4:]))
                data.insert(i,[data[i][0],"提现手续费","支出",data[i][10][4:],"微信"])
                i=i+1

            data[i][3]=data[i][2]+data[i][3]    #讲其他无需进一步处理的交易记录的交易方和商品对象合在一起
            

            temp=[data[i][0],data[i][3],data[i][4],data[i][5][1:],"微信",]
            data[i]=temp
            i+=1



    return data

def format_data(data):
    data.sort()
    for i in data:
        if i[-1]=="微信":
            if i[-3]=="收入":
                i.append(color_orange)
            elif i[-3]=="资金周转":
                i.append(color_gren)
            elif i[-3]=="群收款收入":
                i.append(color_orange)
            else:
                i.append(color_grey)
        if i[-1]=="支付宝":
            if i[-3]=="收入":
                i.append(color_yellow)
            elif i[-3]=="资金周转":
                i.append(color_gren)
            else:
                i.append(color_none)

    return data

def write_excel_format():
    wb=openpyxl.Workbook()
    sheet=wb.worksheets[0]
    sheet.column_dimensions["A"].width=20
    sheet.column_dimensions["B"].width=30
    sheet.column_dimensions["C"].width=20
    sheet.column_dimensions["D"].width=20
    sheet.column_dimensions["E"].width=20
    sheet.column_dimensions["F"].width=20
    sheet.column_dimensions["G"].width=20
    sheet.column_dimensions["H"].width=20

    sheet['A1'].value='日期'
    sheet['B1'].value='项目'
    sheet['C1'].value='金额'
    sheet['D1'].value='可报销支出'
    sheet['E1'].value='其他支出'
    sheet['F1'].value='群收款'
    sheet['G1'].value='收入'
    sheet['H1'].value='备注'
    
    sheet['A2'].value='开始时间'
    sheet['B2'].value='截止时间'

    sheet['A5'].value='收入'
    sheet['A6'].value='群收款收入'
    sheet['A7'].value='支出'

    sheet['A9'].value='可报销支出'
    sheet['A10'].value='其他支出'
    sheet['A11'].value='实际个人支出'

    sheet['A20'].value='日期'
    sheet['B20'].value='项目'
    sheet['C20'].value='金额'
    sheet['D20'].value='可报销支出'
    sheet['E20'].value='其他支出'
    sheet['F20'].value='群收款'
    sheet['G20'].value='收入'
    sheet['H20'].value='备注'

    return wb

def write_excel_data(wb,data):
    startline=20
    week_day=-1
    sheet=wb.worksheets[0]
    for i in data:
        temp_week_day=i[0].weekday()+1
        if temp_week_day!=week_day:
            week_day=temp_week_day
            startline+=1
        sheet['B'+str(startline)].value=i[1]
        sheet['B'+str(startline)].fill=PatternFill('solid', fgColor=i[-1])
        sheet['B'+str(startline)].font=Font(name="等线",size=11)

        if i[2]=="收入" :
            sheet['G'+str(startline)].value=str(i[3])
        elif i[2]=="资金周转":
            sheet['B'+str(startline)].value=str(i[1]+i[3])
        elif i[2]=="群收款收入":
            sheet['F'+str(startline)].value=str(i[3])
        else:
            sheet['C'+str(startline)].value=str(i[3])
        startline+=1



data_wechat=read_data("副本 微信支付账单(20230917-20230924) - 副本.csv","wechat")
data_alipay=read_data("alipay_record_20230924_235400.csv","alipay")
data_total=data_alipay+data_wechat
del data_alipay
del data_wechat
data_total=format_data(data_total)


wb=write_excel_format()

write_excel_data(wb,data_total)
excenName="账单"
i=0
while(1):
    try:
        wb.save(excenName+'.xlsx')
    except:
        i=i+1
        excenName="账单"+str(i)
    else:
        break