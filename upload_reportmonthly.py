


import glob
import os
import pandas as pd
import numpy as np
import re
#pd.set_option('display.max_columns', None) 
#path = r'C:/Users/Windows 10 Version 2/Google Drive/python/pythoncourse/T01.2020'
data_Rec =pd.DataFrame(columns=['State','Receipt Number','Type','Date','Amount','Customer Name','Receipt Method','Activity','Voucher Number','M_Y','Y'])
data_Pay =pd.DataFrame(columns=['Type','Supplier','Sup Num','Supplier Site','Date','Description','Payment Currency','Bank Account','Payment Amount','Payment Rate','Amount','Doc Num','Status','Voucher Number','Rec_CF','Acc_Num','M_Y','Y'])
path = r'C:/Users/Asus/OneDrive/Google Drive/python/REPORT'
extension = 'xls'
os.chdir(path)
result = glob.glob('*.{}'.format(extension))

def Rec_CF(tmp):
    if tmp['Type'] == 'CASH' and tmp[''] == '':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    elif tmp[''] == ''and tmp['Type'] == 'CASH':
        val = ''
    else:
        val = tmp['Activity']
    return val
def Rec_CF_P(data_Pay):
    if data_Pay['Type'] == 'Refund' and data_Pay['']=='':
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    elif data_Pay['Type'] == 'Refund' :
        val = ''
    else:
        val = ''
    return val
def convert_type(x):
    import re
    match_Des =[number for number in re.findall('CM/d{9}[a-zA-Z]{2}',str(x))]
    return match_Des
def convert_type_VTP(x):
    import re
    match_Des =[number for number in re.findall('',str(x))]
    if len(match_Des) == 1:
        res = ''
    else:
        res = ""
    return res
def convert_acc_num(x):
    import re
    num =[number for number in re.findall("/d+", str(x))]
    return num
df1 = pd.read_excel('',sheet_name='',skiprows=1)
df1 = df1.iloc[:,0:11]
df1['Payment Date'] = pd.to_datetime(df1["Payment Date"],dayfirst=True).dt.strftime('%d/%m/%Y')
df1['M_Y'] = df1['Payment Date'].apply(lambda x: str(x)[3:])
pv = pd.pivot_table(df1,index=["M_Y","Type"],values=["Amount"],aggfunc=[np.sum],fill_value=0)
for filename in result:
    df2 = pd.read_excel(filename,index_col=None)
    if df2.iloc[0,0] == '':
        tmp = df2.copy()
        tmp = tmp.iloc[3:,1:-13]
        new_header = tmp.iloc[0] #grab the first row for the header
        tmp = tmp[1:] #take the data less the header row
        tmp.columns = new_header #set the header row as the df header
        tmp= tmp.drop(['Currency','Functional Amount','Unapplied Amount','Deposit Date','GL Date','Status'], axis = 1) #inplace = True)
        tmp['M_Y'] = tmp['Receipt Date'].apply(lambda x: x[3:])
        tmp['Y'] = tmp['Receipt Date'].apply(lambda x: x[6:])
        tmp['Receipt Method'] = tmp['Receipt Method'].str.replace('','')
        tmp['Rec_CF'] = tmp.apply(Rec_CF, axis=1)
        tmp['Rec_CF'] = tmp['Rec_CF'].str.replace('','')
        tmp['Acc_Num'] = tmp['Receipt Method'].apply(convert_acc_num)
        tmp['Acc_Num'] = tmp['Acc_Num'].apply(lambda x: "".join(x))
        tmp = tmp.drop_duplicates(subset=['Receipt Number'])
        tmp = tmp[tmp["State"].str.contains("REV")==False]
        tmp.rename(columns = {'Receipt Date':'Date','Receipt Amount':'Amount'},inplace = True)
        data_Rec = data_Rec.append(tmp)
        # data_Rec = data_Rec.iloc[:,:12]
        # data_Rec['M_Y'] = data_Rec['Date'].apply(lambda x: x[3:])
        # datar = datar.append(data_Rec)
    if df2.iloc[0,0] == '':
        tmp = df2.copy()
        tmp = tmp.iloc[3:,:-9]
        new_header = tmp.iloc[0] #grab the first row for the header
        tmp = tmp[1:] #take the data less the header row
        tmp.columns = new_header #set the header row as the df header
        tmp= tmp.drop(['Rate Type','Rate Date','Operating Unit','Payment Method'], axis = 1) #inplace = True)
        tmp['Payment Amount'] =tmp['Payment Amount'] *-1 
        tmp = tmp[tmp["Bank Account"].str.contains("")==False]
        tmp = tmp[tmp["Status"].str.contains("VOIDED")==False]
        tmp["Bank Account"] = tmp["Bank Account"].str.replace('','')
        tmp['Rec_CF1'] = tmp.apply(Rec_CF_P, axis=1)
        tmp['Rec_CF2'] = tmp['Description'].apply(convert_type)
        tmp['Rec_CF2'] = tmp['Rec_CF2'].apply(lambda x: "".join(x))
        tmp['Rec_CF3'] = tmp['Description'].apply(convert_type_VTP)
        tmp['Rec_CF'] = tmp['Rec_CF2']+tmp['Rec_CF1']+tmp['Rec_CF3']
        tmp= tmp.drop(['Rec_CF1','Rec_CF2','Rec_CF3'], axis = 1) #inplace = True)
        tmp['Acc_Num'] = tmp['Bank Account'].apply(convert_acc_num)
        tmp['Acc_Num'] = tmp['Acc_Num'].apply(lambda x: "".join(x))
        tmp = tmp.fillna({'Payment Rate':1})
        tmp['Functional Amount']= tmp['Payment Amount'] * tmp['Payment Rate']
        tmp.rename(columns = {'Functional Amount':'Amount','Payment Date':'Date'}, inplace = True)
        tmp['M_Y'] = tmp['Date'].apply(lambda x: x[3:])
        tmp['Y'] = tmp['Date'].apply(lambda x: x[6:])
        decimals = 1    
        tmp['Amount'] = tmp['Amount'].apply(lambda x: round(x, decimals))
        data_Pay = data_Pay.append(tmp)
soPhatSinh= pd.concat([data_Rec, data_Pay]).groupby(["Acc_Num"], as_index=False)["Amount"].sum()
data_Rec_thu = data_Rec[data_Rec["Amount"]>0]
data_Rec_thu = data_Rec_thu[['Date','Amount','Acc_Num']]
data_Rec_chi = data_Rec[data_Rec["Amount"]<0]
data_Rec_chi = data_Rec_chi[['Date','Amount','Acc_Num']]
data_Pay_thu = data_Pay[data_Pay['Amount']>0]
data_Pay_thu = data_Pay_thu[['Date','Amount','Acc_Num']]
data_Pay_chi = data_Pay[data_Pay['Amount']<0]
data_Pay_chi = data_Pay_chi[['Date','Amount','Acc_Num']]
TONG_THU = pd.concat([data_Rec_thu, data_Pay_thu])
TONG_THU.rename(columns = {'Amount':'C_Amount'}, inplace = True)
TONG_CHI = pd.concat([data_Rec_chi, data_Pay_chi])
TONG_CHI.rename(columns = {'Amount':'D_Amount'}, inplace = True)
TONG_THUCHI = pd.concat([TONG_CHI, TONG_THU])
TONG_THUCHI =TONG_THUCHI.fillna(0)
TONG_THUCHI['Amount'] = TONG_THUCHI['C_Amount'] + TONG_THUCHI['D_Amount']
TONG_THUCHI['M_Y'] = TONG_THUCHI['Date'].apply(lambda x: x[3:])
PIVOT_TONG_THUCHI = pd.pivot_table(TONG_THUCHI,index=["M_Y","Acc_Num"],values=["D_Amount","C_Amount","Amount"],aggfunc=[np.sum],fill_value=0)
Agg = pd.concat([data_Rec, data_Pay]).groupby(["Rec_CF"], as_index=False)["Amount"].sum().sort_values(['Amount'], ascending=False)
AggT = pd.concat([data_Rec, data_Pay])
Agg_Month= AggT[AggT["M_Y"].str.contains("11/2022",na = False)].groupby(["Rec_CF"], as_index=False)["Amount"].sum().sort_values(['Amount'], ascending=False)

# PIVOT_Agg = pd.pivot_table(Agg,index=["M_Y","Rec_CF"],values=["Amount"],aggfunc=[np.sum],fill_value=0)
# result_agg = PIVOT_Agg.sort_values(('Amount'),ascending = False)
BALANCE = TONG_THUCHI.copy()
BALANCE[["Acc_Num","C_Amount","D_Amount","Amount","Date","M_Y"]]
BALANCE["Bank"] = BALANCE["Acc_Num"]
BALANCE["Bank"].unique()
res = BALANCE.groupby(['Bank','M_Y'])['Amount'].sum().reset_index()
# initialize list of lists
data = [[]]
  
# Create the pandas DataFrame
sddk = pd.DataFrame(data, columns=['Bank', 'Amount'])
sddk['M_Y'] = '12/2021'
sddk = sddk[['Bank','M_Y','Amount']]
import datetime as dt
SD = pd.concat([sddk,res]).sort_values(['Bank','M_Y'])
SD['M_Y'] = pd.to_datetime(SD['M_Y'], format='%m/%Y')
SD = SD.sort_values(['Bank','M_Y'])
SD['M_Y'] = pd.to_datetime(SD['M_Y'], format='%m/%Y').dt.strftime('%m-%Y')
pd.options.display.float_format = "{:,.0f}".format
SD['Balance'] = SD.groupby(['Bank'])['Amount'].cumsum()
SD.to_excel(r'C:/Users/Asus/OneDrive/Google Drive/python/VSCODE/SD.xlsx',index=False)
with pd.ExcelWriter(r'C:/Users/Asus/OneDrive/Google Drive/python/VSCODE/Re.xlsx') as writer: 
    # use to_excel function and specify the sheet_name and index
    # to store the dataframe in specified sheet
    data_Rec.to_excel(writer, sheet_name="data_Rec", index=False)
    data_Pay.to_excel(writer, sheet_name="data_Pay", index=False)
    # soPhatSinh.to_excel(writer, sheet_name="soPhatSinh", index=False)
    Agg.to_excel(writer, sheet_name="Agg", index=False)
    Agg_Month.to_excel(writer, sheet_name="Agg_Month",index=False)
    # TONG_THUCHI.to_excel(writer, sheet_name="TONG_THUCHI", index=False)
    PIVOT_TONG_THUCHI.to_excel(writer, sheet_name="PIVOT_TONG_THUCHI")
    pv.to_excel(writer, sheet_name="")
    SD.to_excel(writer, sheet_name="",index=False)
SD1 = SD.drop(['Amount'], axis = 1)
grouped_obj = SD1.groupby(["Bank"])
# for key, item in grouped_obj:
#     print(str(key))
#     print(str(item), "/n/n")


import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.ticker as ticker
def CT_CN(df):
    import pandas as pd
    if df['CT_CN'] == 'CT':
        val = 'CTY'
    elif df['CT_CN'] == 'CN':
        val = 'CTY'
    elif df['CT_CN'] == 'BA':
        val = 'CTY'
    else:
        val = 'CN'
    return val
#######################################################################################
df = pd.read_excel('', sheet_name='',header=None)
# df.to_csv('your_new_name.csv')
df = df.iloc[1: , :]
new_header = df.iloc[0] #grab the first row for the header
df = df[1:] #take the data less the header row
df.columns = new_header #set the header row as the df header
df['STATE'] = df['STATE'].fillna(value='TCB')
df['Payment Date'] = pd.to_datetime(df['Payment Date'],dayfirst=True).dt.strftime('%d/%m/%Y')
pd.options.display.float_format = '{:,}'.format
df['CT_CN']= df['Name'].str[:2]
df['CT_CN_PL'] = df.apply(CT_CN,axis=1)
df= df.drop(['CT_CN'], axis = 1)
df['Payment Date'] = df['Payment Date'].astype(str)
df['M_Y'] = df['Payment Date'].apply(lambda x: x[3:])
df = df[df["STATE"].str.contains("HUY")==False]
df.columns = df.columns.fillna('to_drop')
df.drop('to_drop', axis = 1, inplace = True)
df['proportion'] = df['Amount']/df['Amount'].sum()*100
df['Amount'] = pd.to_numeric(df['Amount'])
#################################
# Data thanh toán cho Cá Nhân
CN = df[(df['CT_CN_PL'] == 'CN')]
# CN['Amount']= pd.to_numeric(CN['Amount'])

# Data thanh toán cho Cty
CTY = df[(df['CT_CN_PL'] == 'CTY')]
import warnings

warnings.simplefilter(action='ignore', category=UserWarning)







#Filter top 10 Name by A
Top_name_CTY = CTY.groupby(['Name','Type'])['Amount'].sum().sort_values(ascending=False).head(10)
Top_name_CTY = pd.to_numeric(Top_name_CTY)
Top_name_CTY = Top_name_CTY.map('{:,.0f}'.format)
Top_name_CTY.reset_index().style.hide_index()





# plot total Amount by Name
plot_Name_CTY = CTY.groupby(['Name','Type'])['Amount'].sum().sort_values(ascending=False).head(10)
plot_Name_CTY = plot_Name_CTY.sort_values(ascending=True)
plot_Name_CTY
# Plot
fig, ax = plt.subplots(figsize=(15, 10))
plot_Name_CTY.plot.barh(ax=ax) # Try with barh
plt.rcParams.update({'font.size': 22})

# plt.rc('xtick', labelsize=20) 
# plt.rc('ytick', labelsize=20) 
plt.title("Top 10 TỔ CHỨC ĐƯỢC THANH TOÁN NHIỀU NHẤT          ")
# ax.xaxis.get_major_formatter().set_scientific(False)
# ax.plot([0, 1], [0, 2e7])
ax.text(1.3, 0.3, "DVT : Tỷ Đồng", transform = ax.transAxes, ha = "right", va = "baseline")

container = ax.containers[0]
ax.bar_label(container, labels=['{:,.0f}'.format(x/1000000000) for x in container.datavalues])

ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xlabel('Amount')
plt.xticks(rotation=0)
# plt.grid( linestyle='-', linewidth=0.5)
plt.show()
plt.rc('xtick', labelsize=20) 
plt.rc('ytick', labelsize=20) 


# ## TOP 15 TỔ CHỨC ĐƯỢC THANH TOÁN NHIỀU NHẤT THEO TỪNG THÁNG 
#         (DỰA DỮ LIỆU TRÊN THANH TOÁN KHÔNG BAO GỒM CHI LƯƠNG VÀ NỘP THUẾ)
#         TỪ 01/01/2022 - HIỆN TẠI


CTY_PROPO =(CTY.pivot_table(index='Name', columns='M_Y', values='proportion',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   # .sort_values('All', ascending=False, axis=1) # sort by column sum?
   .drop('All')    # drop row `All`
    
).head(15)

CTY_PROPO= CTY_PROPO.applymap('{0:.2f}%'.format)

# CTY_PROPO
CTY_VALUE =(CTY.pivot_table(index='Name', columns='M_Y', values='Amount',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   # .sort_values('All', ascending=False, axis=1) # sort by column sum?
   .drop('All')    # drop row `All`
    
).head(15)
CTY_VALUE = CTY_VALUE.applymap('{:,.0f}'.format)
CTY_VALUE = CTY_VALUE.reset_index()
CTY_PROPO = CTY_PROPO.reset_index()
CTY_PROPO = CTY_PROPO[['Name','All']]
result_cty = pd.merge(CTY_VALUE, CTY_PROPO, on=["Name"])
result_cty = result_cty[['Name','All_x', 'All_y', '01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
result_cty.rename(columns = {'All_x':'Total','All_y':'%'},inplace = True)
# result_cty.style.hide_index()
add_total = df[['Amount']].sum().rename('Total').fillna('').reset_index()
add_total = add_total.round(0)

# add_total = add_total.reset_index()
# add_total = add_total.applymap('{:,.0f}'.format)
add_total = add_total.append(result_cty)
add_total = add_total[['Name','Total', '%', '01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
add_total = add_total.fillna('')
# add_total.head(15)
# result_cty=result_cty.applymap('{:,.0f}'.format)
# result_cty
add_total.style.hide_index()
pd.options.display.float_format = '{:,.0f}'.format
# add_total.style.hide_index()
add_total.set_index('Name')


# ## TOP 15 CÁ NHÂN ĐƯỢC THANH TOÁN NHIỀU NHẤT
# (DỰA DỮ LIỆU TRÊN THANH TOÁN KHÔNG BAO GỒM CHI LƯƠNG VÀ NỘP THUẾ)

# DỮ LIỆU TỪ 01/01/2022 - ĐẾN HIỆN TẠI



CN_VALUE =(CN.pivot_table(index='Name', columns='Type', values='Amount',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(20)

CN_VALUE = CN_VALUE.applymap('{:,.0f}'.format)
CN_PROPO =(CN.pivot_table(index='Name', columns='Type', values='proportion',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(20)
CN_PROPO = CN_PROPO.applymap('{0:.2f}%'.format)

CN_VALUE = CN_VALUE.reset_index()
CN_PROPO = CN_PROPO.reset_index()
CN_PROPO = CN_PROPO[['Name','All']]

result_cn = pd.merge(CN_VALUE, CN_PROPO, on=["Name"])
result_cn= result_cn[['Name','All_x', 'All_y', '', '', '', '']]
result_cn.rename(columns = {'All_x':'Total','All_y':'%'},inplace = True)
add_totalcn = df[['Amount']].sum().rename('Total').fillna('').reset_index()
# add_totalcn = add_totalcn.round(0)

# # add_totalcn = add_totalcn.applymap('{:,.0f}'.format)

add_totalcn = add_totalcn.append(result_cn)
add_totalcn = add_totalcn[['Name','Total', '%', '', '', '', 'CONG DOAN']]
add_totalcn = add_totalcn.fillna('')


pd.options.display.float_format = '{:,.0f}'.format
# add_total.style.hide_index()
# add_total.set_index('Name')

# pd.options.display.float_format = '{:,0f}'.format
# add_totalcn['Total'] = add_totalcn['Total'].apply(lambda x : '{:,.0f}'.format(x))
# add_totalcn.set_index('Name')
# result_cn
add_totalcn.set_index('Name')

# # # result_cty.style.hide_index()

# CTY_PROPO||


# ## Top CÁC LOẠI CHI PHÍ ĐƯỢC THANH TOÁN NHIỀU NHẤT 
#     DỮ LIỆU BAO GỒM LƯƠNG VÀ THUẾ
#     TỪ 01/01/2022 - 30/11/2022

# ### TOTAL : TỔNG CÁC KHOẢN CHI BAO GỒM THUẾ VÀ LƯƠNG

# In[57]:



AggT['Amount'] = pd.to_numeric(AggT['Amount'])
AggT_Pay = AggT[AggT['Amount']<0]

AggT_Rec = AggT[AggT['Amount']>0]
# AggT_Pay
AggT_Pay = AggT_Pay[['Date','Amount','M_Y','Rec_CF']]
AggT_Pay['Amount'] = AggT_Pay['Amount']*-1
AggT_Pay['proportion'] = AggT_Pay['Amount']*100 / AggT_Pay['Amount'].sum()
# AggT_Pay
AggT_Pay_VALUE =(AggT_Pay.pivot_table(index='Rec_CF', columns='M_Y', values='Amount',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(15)

AggT_Pay_VALUE = AggT_Pay_VALUE.applymap('{:,.0f}'.format)

AggT_Pay_VALUE = AggT_Pay_VALUE.reset_index()
AggT_Pay_VALUE = AggT_Pay_VALUE[['Rec_CF','All','01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
# AggT_Pay_VALUE
AggT_Pay_PROPOR =(AggT_Pay.pivot_table(index='Rec_CF', columns='M_Y', values='proportion',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(21)

AggT_Pay_PROPOR= AggT_Pay_PROPOR.applymap('{0:.2f}%'.format)

AggT_Pay_PROPOR = AggT_Pay_PROPOR.reset_index()
AggT_Pay_PROPOR = AggT_Pay_PROPOR[['Rec_CF','All']]
# AggT_Pay_PROPOR
result_AggT_Pay = pd.merge(AggT_Pay_VALUE, AggT_Pay_PROPOR, on=["Rec_CF"])
result_AggT_Pay = result_AggT_Pay[['Rec_CF','All_x','All_y','01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
result_AggT_Pay.rename(columns = {'All_x':'Total','All_y':'%'},inplace = True)
add_totalpay = AggT_Pay[['Amount']].sum().rename('Total').fillna('').reset_index()
add_totalpay = add_totalpay.reset_index()
add_totalpay = add_totalpay.append(result_AggT_Pay)
add_totalpay = add_totalpay[['Rec_CF','Total','%','01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
add_totalpay = add_totalpay.fillna('')
add_totalpay.set_index('Rec_CF')


# plot total Amount by type
plot_type = data_Pay.groupby(["Rec_CF"])['Amount'].sum().sort_values(ascending=True).head(17)*-1
# Plot
fig, ax = plt.subplots(figsize=(25, 10))
plot_type.plot.bar(ax=ax) # Try with barh
ax.yaxis.get_major_formatter().set_scientific(False)
ax.plot([0, 1], [0, 2e7])
ax.text(-0.15, 1.05, "Billions", transform = ax.transAxes, ha = "left", va = "top")
container = ax.containers[0]
ax.bar_label(container, labels=['{:,.0f}'.format(x/1000000000) for x in container.datavalues])
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xticks(rotation=35)
plt.title("Top CÁC LOẠI CHI PHÍ ĐƯỢC THANH TOÁN NHIỀU NHẤT ")

plt.show()


# ## TỔNG TIỀN THANH TOÁN TỪ CÁC TÀI KHOẢN NGÂN HÀNG CỦA CTY
# ### TOTAL : TỔNG GIÁ TRỊ THANH TOÁN BẰNG EFORM QUA CHUYỂN KHOẢN NGÂN HÀNG KHÔNG BAO GỒM THUẾ VÀ LƯƠNG


# Top_Bank_count = Top_Bank_count.reset_index()
# Top_Bank_count['Code']= pd.to_numeric(Top_Bank_count['Code'])
# # # Top_Bank2 = Top_Bank2.groupby(['Bank'])['Code'].count().sort_values(ascending=False)
# Top_Bank_count = (Top_Bank_count.append(Top_Bank_count.sum(numeric_only=True), ignore_index=True))
# Top_Bank_count = Top_Bank_count.fillna(value='All')
# # Top_Bank_count




# Filter payment from Bank
Top_Bank = df.groupby(['STATE'])['Amount'].sum().sort_values(ascending=False)
Top_Bank = pd.to_numeric(Top_Bank)
# Top_Bank = Top_Bank.map('{:,.0f}'.format)
Top_Bank = Top_Bank.reset_index()
Top_Bank['Amount'] = pd.to_numeric(Top_Bank['Amount'])
Top_Bank= (Top_Bank.append(Top_Bank.sum(numeric_only=True), ignore_index=True))
Top_Bank = Top_Bank.fillna(value='All')
pd.options.display.float_format = '{:,.0f}'.format
Top_Bank


# plot total Amount by Bank
plot_Bank = df.groupby(['STATE'])['Amount'].sum().sort_values(ascending=False)
# Plot
fig, ax = plt.subplots(figsize=(7, 10))
plot_Bank.plot.bar(ax=ax) # Try with barh
plt.title("TỔNG TIỀN THANH TOÁN TỪ CÁC TÀI KHOẢN NGÂN HÀNG CỦA CTY")
ax.yaxis.get_major_formatter().set_scientific(False)
ax.plot([0, 1], [0, 2e7])
ax.text(-0.35, 0.97, "Billions", transform = ax.transAxes, ha = "left", va = "top")
container = ax.containers[0]
ax.bar_label(container, labels=['{:,.0f}'.format(x/1000000000) for x in container.datavalues])
ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xticks(rotation=90)
# plt.grid( linestyle='-', linewidth=0.5)
plt.show()


# ## TỔNG SỐ LỆNH THANH TOÁN TỪ CÁC TÀI KHOẢN NGÂN HÀNG CỦA CTY
# (KHÔNG BAO GỒM CHI LƯƠNG VÀ THUẾ)

# In[40]:


Top_Bank_count = df.groupby(['STATE'])['Code'].count().sort_values(ascending=False)
# # Top_Bank_count = pd.to_numeric(Top_Bank)
# # Top_Bank_count = Top_Bank_count.map('{:,.0f}'.format)
Top_Bank_count = Top_Bank_count.reset_index()
Top_Bank_count['Code']= pd.to_numeric(Top_Bank_count['Code'])
# # Top_Bank2 = Top_Bank2.groupby(['Bank'])['Code'].count().sort_values(ascending=False)
Top_Bank_count = (Top_Bank_count.append(Top_Bank_count.sum(numeric_only=True), ignore_index=True))
Top_Bank_count = Top_Bank_count.fillna(value='All')
pd.options.display.float_format = '{:,.0f}'.format
# Top_Bank_count = (Top_Bank_count.append(Top_Bank_count.sum(numeric_only=True), ignore_index=True))
Top_Bank_count



# plot total Amount by Bank
plot_Bank = df.groupby(['STATE'])['Code'].count().sort_values(ascending=False)
# Plot
fig, ax = plt.subplots(figsize=(8, 10))
plot_Bank.plot.bar(ax=ax) # Try with barh
plt.title("TỔNG SỐ LỆNH THANH TOÁN TỪ CÁC TÀI KHOẢN NGÂN HÀNG CỦA CTY")
ax.yaxis.get_major_formatter().set_scientific(False)
# ax.plot([0, 1], [0, 2e7])
ax.text(-0.35, 0.97, "Billions", transform = ax.transAxes, ha = "left", va = "top")
container = ax.containers[0]
ax.bar_label(container)
# # ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xticks(rotation=90)
# plt.grid( linestyle='-', linewidth=0.5)
plt.show()



# Filter quantity payment from Bank
Top_Bank2 = df.groupby(['Bank'])['Code'].count().sort_values(ascending=False).head(15)
Top_Bank2 = pd.to_numeric(Top_Bank2)

Top_Bank2 = Top_Bank2.reset_index()
Top_Bank2['Code']= pd.to_numeric(Top_Bank2['Code'])
# Top_Bank2 = Top_Bank2.groupby(['Bank'])['Code'].count().sort_values(ascending=False)
Top_Bank2 = (Top_Bank2.append(Top_Bank2.sum(numeric_only=True), ignore_index=True))
Top_Bank2 = Top_Bank2.fillna(value='All')
pd.options.display.float_format = '{:,.0f}'.format
# Top_Bank2['Code'] = Top_Bank2['Code'].apply(lambda x: str(x))
# Top_Bank2 = Top_Bank2.applymap('{:,.0f}'.format)
Top_Bank2



# plot quantity Payments by Bank
plot_Bank_quantity = df.groupby(['Bank'])['Code'].count().sort_values(ascending=False).head(10)
# Plot
fig, ax = plt.subplots(figsize=(15, 10))
plot_Bank_quantity.plot.bar(ax=ax) # Try with barh
plt.title("SỐ LƯỢNG LỆNH THANH TOÁN ĐẾN CÁC NGÂN HÀNG")
ax.yaxis.get_major_formatter().set_scientific(False)
container = ax.containers[0]
ax.bar_label(container)
# ax.plot([0, 1], [0, 2e7])
# ax.text(0, 1.05, "Billions", transform = ax.transAxes, ha = "left", va = "top")
# ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xticks(rotation=60)
# plt.grid( linestyle='-', linewidth=0.5)
plt.show()



AggT_Rec= AggT_Rec[['Date','Amount','M_Y','Rec_CF']]
# AggT_Rec['Amount'] = AggT_Pay['Amount']*-1
AggT_Rec['proportion'] = AggT_Rec['Amount']*100 / AggT_Rec['Amount'].sum()
# AggT_Rec
AggT_Rec_VALUE =(AggT_Rec.pivot_table(index='Rec_CF', columns='M_Y', values='Amount',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(14)

AggT_Rec_VALUE = AggT_Rec_VALUE.applymap('{:,.0f}'.format)

AggT_Rec_VALUE = AggT_Rec_VALUE.reset_index()
# AggT_Rec_VALUE
AggT_Rec_PROPOR =(AggT_Rec.pivot_table(index='Rec_CF', columns='M_Y', values='proportion',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   # .drop('All', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   .drop('All')    # drop row `All`
).head(14)

AggT_Rec_PROPOR = AggT_Rec_PROPOR.applymap('{0:.2f}%'.format)

AggT_Rec_PROPOR = AggT_Rec_PROPOR.reset_index()
AggT_Rec_PROPOR = AggT_Rec_PROPOR[['Rec_CF','All']]
# AggT_Rec_PROPOR
# AggT_Rec_VALUE = AggT_Rec_VALUE[['Rec_CF','All']]

result_AggT_Rec = pd.merge(AggT_Rec_VALUE, AggT_Rec_PROPOR, on=["Rec_CF"])
result_AggT_Rec = result_AggT_Rec[['Rec_CF','All_x','All_y','01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
result_AggT_Rec.rename(columns = {'All_x':'Total','All_y':'%'},inplace = True)
# result_AggT_Rec


add_totalrec = AggT_Pay[['Amount']].sum().rename('Total').fillna('').reset_index()
add_totalrec = add_totalrec.reset_index()
add_totalrec = add_totalrec.append(result_AggT_Rec)
add_totalrec = add_totalrec[['Rec_CF','Total','%','01/2022', '02/2022', '03/2022', '04/2022', '05/2022', '06/2022', '07/2022', '08/2022', '09/2022', '10/2022','11/2022']]
add_totalrec = add_totalrec.fillna('')
add_totalrec.set_index('Rec_CF')


Agg1 = pd.concat([data_Rec, data_Pay])
Agg2 = Agg1.groupby(["Rec_CF"])["Amount"].sum().sort_values( ascending=False).head(13)
Agg2 = Agg2.map('{:,.0f}'.format)
# Agg2


TOP_DT = Agg1.groupby(["Rec_CF"])["Amount"].sum().sort_values( ascending=False).head(12)
TOP_DT = TOP_DT.sort_values(ascending=True)
TOP_DT
# Plot
fig, ax = plt.subplots(figsize=(15, 10))
TOP_DT.plot.barh(ax=ax) # Try with barh
plt.title("TOP CÁC KHOẢN THU")
# plt.rcParams.update({'font.size': 22})
# plt.rc('xtick', labelsize=20) 
# plt.rc('ytick', labelsize=20) 
# ax.xaxis.get_major_formatter().set_scientific(False)
# ax.plot([0, 1], [0, 2e7])
ax.text(1.1, 0, "Billions", transform = ax.transAxes, ha = "right", va = "baseline")

container = ax.containers[0]
ax.bar_label(container, labels=['{:,.0f}'.format(x/1000000000) for x in container.datavalues])

ax.xaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.show()


# ## SỐ LƯỢNG LỆNH THANH TOÁN THEO GIÁ TRỊ

bins = [-np.infty, 100000000, 500000000, np.infty]
labels = ["<100M", "<500M", ">500M"]
df["Amount_Type"] = pd.cut(df["Amount"], bins=bins, labels=labels, right=False)




fig, ax = plt.subplots(figsize=(10, 10))
a.plot.bar(ax=ax) # Try with barh

plt.title("SỐ LƯỢNG LỆNH THANH TOÁN THEO GIÁ TRỊ")
ax.yaxis.get_major_formatter().set_scientific(False)
container = ax.containers[0]
ax.bar_label(container)
# ax.plot([0, 1], [0, 2e7])
ax.text(-0.3, 0.95, "Billions", transform = ax.transAxes, ha = "left", va = "top")
# ax.yaxis.set_major_formatter(ticker.FuncFormatter(lambda y, pos: '{:,.0f}'.format(y/1000000000))) 
plt.xticks(rotation=0)
# plt.grid( linestyle='-', linewidth=0.5)
plt.show()



a =(NTBP.pivot_table(index='Name', columns='Type', values='Amount',
               aggfunc='sum', fill_value=0, margins=True)   # pivot with margins 
   .sort_values('All', ascending=False)  # sort by row sum
   .drop('DEN BU', axis=1)                  # drop column `All`
   .sort_values('All', ascending=False, axis=1) # sort by column sum
   # .drop('All')    # drop row `All`
)
a.applymap('{:,.0f}'.format)

