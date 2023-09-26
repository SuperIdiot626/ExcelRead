import datetime
import csv

encode_type_alipay='GB18030'
encode_type_wechat='UTF-8'




def read_data(file_path,type):
    data=[]
    with open(file_path, 'r',encoding=type,) as file:    #读取支付宝文件
        reader = csv.reader(file)
        for row in reader:
            data.append(row)
    for i in range(len(data)):
        if data[i]=="交易时间":
            break
    data=data[i:]
    print(data)
    return data



data_wechat=[]      #空列表
data_alipay=[]      #空列表

with open('alipay_record_20230924_235400.csv', 'r',encoding='GB18030',) as file:    #读取支付宝文件
    reader = csv.reader(file)
    for row in reader:
        data_alipay.append(row)

for i in data_alipay



with open('微信支付账单(20230917-20230924).csv', 'r',encoding='UTF-8') as file:     #读取微信文件
    reader = csv.reader(file)
    for row in reader:
        data_wechat.append(row)




data_wechat=(data_wechat[17:])
data_alipay=(data_alipay[25:])
data_wechat.reverse()
data_alipay.reverse()


date_str = '2023-09-18 00:02:50'
date_obj = datetime.datetime.strptime(date_str, "%Y-%m-%d %H:%M:%S" )
print(date_obj)

print("!!!")