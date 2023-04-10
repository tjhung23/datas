import glob
import os
import numpy as np
import re
import parse
import pdfplumber
import pandas as pd
import collections
from collections import namedtuple
#pd.set_option('display.max_columns', None) 
# path = r'C:/Users/Asus/OneDrive/Google Drive/python/pdf'
path = os.getcwd()
# files = os.listdir(path)
# path = r'D:/DOWNLOADS/PDF'X
extension = 'pdf'
# Line = namedtuple ('Line', 'Name Date Amount VAT SoHD')
Line = namedtuple ('Line', 'EF BP nguoiDN Date Name STK Bank Amountbf Amount')
# Line = namedtuple ('Line','Name STK Bank Amount')
os.chdir(path)
result = glob.glob('*.{}'.format(extension))

Comname_re = re.compile('Công\sty\sCổ\sPhần\sDịch\sVụ\sGiao\sHàng\sNhanh')
# line_re = re.compile('Tên/sđơn/svị/s[(]Company[)]/s[:]/s')
# line_re = re.compile('Tên đơn vị (Company) : ')
lines =[]
for filename in result:
        with pdfplumber.open(filename) as pdf:
            pages = pdf.pages
            for page in pdf.pages:
                text = page.extract_text()
                # print(text)
                for line in text.split('/n'):
                    # print(line)
                    comp = Comname_re.search(line)
                    match_EF = re.search('Mã\sđề\snghị[:]\s(.*)',line)
                    match_BP = re.search('Bộ\sphận\sđề\snghị[:]\s(.*)',line)
                    match_nguoiDN = re.search('Người\sđề\snghị[:]\s(.*)\sBộ',line)
                    match_Date = re.search('Tổng\sgiám\sđốc\s(.*)\n',line)
                    match_NhomCP = re.search('[(]\d{1}[)]\s[(][(]\d{1}[)][+][(]\d{1}[)][-][(]\d{1}[)][)]\s(.*)\n',line)
                    match_Name = re.search('(Người\sthụ\shưởng[:]|Người\shưởng\sthụ[:])\s(.*)\n(.*)\n(.*)\n(.*)',line)
                    match_STK = re.search('ài\skhoản[:]\s(.*)',line)
                    match_Bank = re.search('Ngân\shàng[:](.*)\n(.*)\n(.*)\n(.*)',line)
                    match_Amount = re.search('(Số\stạm\sứng\sthiếu[:]|Số\stiền\stạm\sứng[:])\s(.*)\sVND',line)
                    match_Amountbf = re.search('Số\stiền\s(.*)\sVND',line)
                    # match_Noidung = re.search('(Nội\sdung\sthanh\stoán[:]|Nội\sdung\stạm\sứng[:])(.*)\n(.*)',line)
                    # match_Amount_TU = re.search('Tổng\ssố\stiền\sđề\snghị\s(.*)\sVND',line)
                    if match_EF:
                        EF = re.findall('Mã\sđề\snghị[:]\s(.*)', match_EF.group())
                        EF = ''.join(str(x) for x in EF)
                        # print(EF)
                    if match_BP:
                        BP = re.findall('Bộ\sphận\sđề\snghị[:]\s(.*)', match_BP.group())
                        BP = ''.join(str(x) for x in BP).replace(" Thử Vai","").replace(" Tiền Thuê 2455 - Bưu","")
                        # print(BP)
                    if match_nguoiDN:
                        nguoiDN = re.findall('Người\sđề\snghị[:]\s(.*)\sBộ', match_nguoiDN.group())
                        nguoiDN = ''.join(str(x) for x in nguoiDN).replace(" Thử Vai","").replace(" Tiền Thuê 2455 - Bưu","")
                        # print(nguoiDN)
                    if match_Date:
                        Date = re.findall('Tổng\sgiám\sđốc\s(.*)\n', match_Date.group())
                        Date = re.findall('\d{2}/\d{2}/\d{4}', match_Date.group())
                        Date = ''.join(str(x) for x in Date)
                        # print(Date)
                    # if match_NhomCP:
                    #     NhomCP = re.findall('[(]\d{1}[)]\s[(][(]\d{1}[)][+][(]\d{1}[)][-][(]\d{1}[)][)]\s(.*)\n', match_NhomCP.group())
                    #     NhomCP = ''.join(str(x) for x in NhomCP).split(',')[0]
                        # print(NhomCP)
                    if match_Name:
                        Name = re.findall(r'(Người\sthụ\shưởng[:]|Người\shưởng\sthụ[:])\s(.*)\n(.*)\n(.*)\n(.*)', match_Name.group())
                        # Name = ''.join(str(x) for x in Name).split(",")[-1].replace("')","").replace("'","")
                        Name = ''.join(str(x) for x in Name).split(",")[1:]
                        # new_name = re.findall(r'((.*?)(Tài|Số) ', Name.group())
                        # Name = ''.join(str(x) for x in Name).replace('"',"")
                        Name = ''.join(str(x) for x in Name)
                        Name = re.match("(.*?)(Tài|Số)",Name).group().replace("'Tài","").replace("'Số","").replace("'","")
                        # print(Name)
                    if match_STK:
                        STK = re.findall('ài\skhoản[:]\s(.*)', match_STK.group())
                        STK = ''.join(str(x) for x in STK).replace(" Thanh","").replace(" Lan","").replace(" Không tiến hành","").replace(" 1712017 - Phạm Thị","").replace(" Hạnh","").replace(" Minh Trang","").replace(" Mai","").replace(" Hương","").replace(" Thu","").replace(" Thư","").replace(" Lệ","").replace( " Thu Hạnh","").replace( " 2032023 Thu ","").replace(" - Phùng Minh","").replace(" Tuấn","").replace(" Đã được duyệt bởi","").replace(" Xuân Hòa","").replace(" 3026068","").replace(" Trang","")
                        # print(STK)
                    if match_Bank:
                        Bank = re.findall('Ngân\shàng[:](.*)\n(.*)\n(.*)\n(.*)', match_Bank.group())
                        Bank = ''.join(str(x) for x in Bank).replace("(","").replace(")","").replace("',","").replace("'","")
                        # Bank = re.match("(.*?)Người đề",Bank).group().replace("Người đề","").replace("'","").lstrip()
                        # print(Bank)
                    if match_Amount:
                        Amount = re.findall('(Số\stạm\sứng\sthiếu[:]|Số\stiền\stạm\sứng[:])\s(.*)\sVND',match_Amount.group())
                        Amount = ''.join(str(x) for x in Amount).replace("'Số tạm ứng thiếu:', ","").replace("'Số tiền tạm ứng:', ","").replace("'","").replace("(","").replace(")","")
                        # print(Amount)
                    if match_Amountbf:
                        Amountbf = re.findall('Số\stiền\s(.*)\sVND',match_Amountbf.group())
                        Amountbf = ''.join(str(x) for x in Amountbf).split(':')[-1]   
                        lines.append(Line(EF, BP, nguoiDN, Date, Name, STK, Bank, Amountbf, Amount))
                    # if match_Noidung:
                    #     Noidung = re.findall('(Nội\sdung\sthanh\stoán[:]|Nội\sdung\stạm\sứng[:])(.*)\n(.*)', match_Noidung.group())
                    #     Noidung = ''.join(str(x) for x in Noidung).replace('"','').replace("'","").replace("[","").replace("]","").replace(",","").replace(":","").replace("Nội dung thanh toán","").replace("Nội dung tạm ứng","").replace("(","").replace(")","")
                    #     Noidung = re.sub('Số\stiền\s(.*)\sHạn\sthanh\stoán', '', Noidung)
                    #     Noidung = re.sub('Số\stiền\s(.*)\sVND', '', Noidung)
                        
                        # print(Noidung)
df = pd.DataFrame(lines)
df['EF'] = df['EF'].apply(lambda x: x.split(" ", 1)[0])
new = df['Bank'].str.split("-", n = 1, expand = True)
df['Bank'] = new[0]
df['Name'] = df['Name'].str.strip()
df['Bank'] = df['Bank'].str.strip().replace("Người đề","").replace("'","")
df['STK'] = df['STK'].str.strip()
df['STK'] = df['STK'].str.replace('-','')
df['STK'] = df['STK'].str.replace(' ','')
df['STK'] = df['STK'].str.replace('.','')
# df['Noidung'] = df['Noidung'].str.strip()
df['Bank'] = df['Bank'].replace(['ABBank','ACB', 'Agribank', 'Techcombank', 'SHB', 'Deutsche Bank AG, Vietnam', 'Sacombank', 'Vietcombank', 'Tien Phong Bank', 'Vietinbank', 'HDBank',  'Nam A Bank','Citibank','Maritime Bank','Dong A Bank','LienVietPostBank','Viet Capital Bank','VIBank','Ngân hàng Mizuho CN TP Hồ Chí Minh','KienLongBank','Eximbank','Viet A Bank','BaoViet Bank','SeABank'], ['ABBANK', 'ACB', 'Agribank','TECHCOMBANK','SHB','DEUTSCHE BANK','Sacombank','Vietcombank','TP Bank','Vietinbank','HD Bank','Nam A','Citi HCM','Maritimebank','Dong A','LPB','BAN VIET','VIB','MIZUHO HCM','Kien Long Bank','EXIM BANK','Viet A','BAO VIET BANK','Sea Bank'])
df.insert(9, 'STT', range(1, 1 + len(df)))
df.to_excel(r'NEF.xlsx',index=False)
print('done')

# print(df)
