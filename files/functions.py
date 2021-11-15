from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from os import path,getcwd
import pandas as pd
import re
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from string import Template
import requests


def check_enroll(df):
    data_col = list(df.columns)
    regex = re.compile(r'(((U|u)(niversity)\s(R|r)(oll)\s(Number|number|No.|No|no))|((E|e)(nrollment)\s(Number|number|No.|No|no)))')
    unirllno = list(filter(regex.match,data_col))
    unirllno = "".join(unirllno)
    for index,rows in df.iterrows():
        enroll = str(rows[unirllno])
        if re.match('^\d+$',enroll) == None:
            ind = -1
            for i in enroll:
                if not i.isdigit():
                    ind = enroll.index(i)
                    break
            enroll = enroll[0:ind]
        df.iloc[index,data_col.index(unirllno)] = int(enroll)
    df[unirllno]=df[unirllno].apply(lambda x: '{0:0>11}'.format(x))

def check_null(df):
    data_col = list(df.columns)
    regex = re.compile(r'((10th|Tenth)\s(%|perc|percentage))|((%)\s(M|m)(arks)(\s|)(-)(\s|)(10th|Tenth))')
    perc_10 = list(filter(regex.match,data_col))
    regex = re.compile(r'((12th|twelfth)\s(%|perc|percentage))|((%)\s(M|m)(arks)(\s|)(-)(\s|)(12th|twelfth))')
    perc_12 = list(filter(regex.match,data_col))
    regex = re.compile(r'((B|b)(\.?)(tech|Tech)\s(%|perc|percentage))|((%)\s(M|m)(arks)(\s|)(-)(\s|)(B|b)(\.?)(tech|Tech))|((A|a)(ggregate)\s(%|perc|percentage)(\s|)(-)(\s|)Graduation\s(\(without credits\)|))')
    perc_btech = list(filter(regex.match,data_col))
    regex = re.compile(r'((D|d)(iploma)\s(%|perc|percentage))|((%)\s(M|m)(arks)(\s|)(-)(\s|)(D|d)(iploma))')
    perc_diploma = list(filter(regex.match,data_col))
    regex = re.compile(r'(B|b)(ack|acklog|ackLog)(s|)')
    backlog = list(filter(regex.match,data_col))
    regex = re.compile(r'(C|c)(g|G)(p|P)(a|A)')
    cgpa = list(filter(regex.match,data_col))
    
    perc_10 = "".join(perc_10)
    perc_12 = "".join(perc_12)
    perc_btech = "".join(perc_btech)
    backlog = "".join(backlog)
    perc_diploma = "".join(perc_diploma)
    cgpa = "".join(cgpa)
    
    if(df[perc_10].isnull().any() == 1):
        df[perc_10].fillna(0,inplace = True)
    if(df[perc_12].isnull().any()==1):
        df[perc_12].fillna(0,inplace = True)
    if(df[perc_btech].isnull().any()==1):
        df[perc_btech].fillna(0,inplace = True)
    if(df[perc_diploma].isnull().any()==1):
        df[perc_diploma].fillna(0,inplace = True)
    if(df[backlog].isnull().any()==1):
        df[backlog].fillna(0,inplace = True)
    if(df[cgpa].isnull().any()==1):
        df[cgpa].fillna(0,inplace = True)
    
    for i,r in df.iterrows():
        if type(r[backlog])!=int:
            df.iloc[i,data_col.index(backlog)] = 0

    return [perc_10,perc_12,perc_btech,perc_diploma,backlog,cgpa]

def set_tenth(df,tenth):
    data_col = list(df.columns)
    for index,row in df.iterrows():
        if row[tenth] <= 10:
            df.iloc[index,data_col.index(tenth)] = 9.5 * row[tenth]

def counter(n):
    count = 0
    while n>0:
        count+=1
        n=n//10
    return count


def mobile_number_edit(df):
    data_col = list(df.columns)
    regex = re.compile(r'(((M|m)(obile)\s(Number|number|No.|No|no))|((P|p)(hone)|((P|p)(h)(\.?))\s(Number|number|No.|No|no)))')
    mobile = list(filter(regex.match,data_col))
    mobile = "".join(mobile)
    for index,row in df.iterrows():
        try:
            t = int(row[mobile])
            count = counter(t)
            if count != 10 :
                if count == 12 and t//(10**10) ==91:
                    k = t%(10**10)
                    df.iloc[index,data_col.index(mobile)] = k
                else:
                    df.iloc[index,data_col.index(mobile)] = 'Invalid Number'
                #print('Mobile Number at ',index,' to ', df.iloc[index][mobile])
        except:
            if re.match('^[0-9]+,[ 0-9]+',row[mobile]) != None:
                num1 = int(row[mobile].split(', ')[0])
                num2 = int(row[mobile].split(', ')[1])
                count = counter(num1)
                if count == 10:
                    df.iloc[index,data_col.index(mobile)] = num1
                elif count == 12 and num1//(10**10) ==91:
                    k = num1%(10**10)
                    df.iloc[index,data_col.index(mobile)] = k
                else:
                    df.iloc[index,data_col.index(mobile)] = 'Invalid Number'
                #print('Mobile Number at ',index,' to ', df.iloc[index][mobile])
                # count = counter(num2)
                # if count == 10:
                #     df.iloc[index,15] = num2
                # elif count == 12 and num2//(10**10) ==91:
                #     k = num2%(10**10)
                #     df.iloc[index,15] = k
                # else:
                #     df.iloc[index,15] = 'Invalid Number'
                #print('Residence No changed at ',index,' to ', df.iloc[index]['Residence No'])
            elif re.match('^[0-9 \u202c]+',row[mobile]) != None:
                mylist = row[mobile].split()
                num = ''
                for item in mylist:
                    num = num + item
                try:
                    num = int(num)
                    count = counter(num)
                    if count == 10:
                        df.iloc[index,data_col.index(mobile)] = num
                    elif count == 12 and num//(10**10) ==91:
                        k = num%(10**10)
                        df.iloc[index,data_col.index(mobile)] = k
                    else:
                        df.iloc[index,data_col.index(mobile)] = 'Invalid Number'
                    #print('Mobile Number changed at ',index,' to ', df.iloc[index][mobile])
                except:
                    num = list(num)
                    new = ''
                    for i in num:
                        if i.isdigit():
                            new = new + i
                    num = int(new)
                    count = counter(num)
                    if count == 10:
                        df.iloc[index,data_col.index(mobile)] = num
                    elif count == 12 and num//(10**10) ==91:
                        k = num%(10**10)
                        df.iloc[index,data_col.index(mobile)] = k
                    else:
                        df.iloc[index,data_col.index(mobile)] = 'Invalid Number'
                    #print('Mobile Number changed at ',index,' to ', df.iloc[index][mobile])
            else:
                print('ERROR IN PHONE NUMBER AT INDEX : ',index)




def process(m10,m12,bt,btech,dip,backs,company,location,stream):
    filetype = location.split('.')[-1]
    df = pd.DataFrame()
    if filetype.upper() != 'CSV' and filetype.upper() != 'XLS' and filetype.upper() != 'XLSX' and filetype.upper() != 'XLSM':
            return "FILE ERROR"
    else:
            location = Path(location)
            if filetype.upper() == 'CSV':
                    df = pd.read_csv(location)
            else:
                    df = pd.read_excel(location)
    check_enroll(df)
    collist = check_null(df) # collist = [10%,12%, btech%, diploma %, backlog]
    set_tenth(df,collist[0])
    mobile_number_edit(df)
    data_col = list(df.columns)
    regex = re.compile(r'(c|C)(lass)$')
    s_class = list(filter(regex.match,data_col))
    s_class = "".join(s_class)
    
    if m10 != '' :
        df = df[df[collist[0]] >= int(m10)]
    
    if m12 != '' and dip != '':
        df = df[(df[collist[1]] >= int(m12)) | (df[collist[3]] >= int(dip))]
    elif m12 != '':
        df = df[df[collist[1]] >= int(m12)]
    elif dip != '':
        df = df[df[collist[3]] >= int(dip)]

    if bt != '':
        if(btech=="perc"):
            df = df[df[collist[2]] >= int(bt)]
        elif(btech=="cgpa"):
            df = df[df[collist[5]] >= int(bt)]

    if backs != '':
        df = df[df[collist[4]] <= int(backs)]
    
    final = df[df[s_class].str.match(stream)]
    finalloc = path.join(getcwd(),"Database",company)
    try:
        final.to_csv(finalloc+'.csv',index = False)
        final.to_excel(finalloc+'.xlsx',index = False)
    except:
        return "FILE OPEN"
    return path.join(getcwd(),finalloc+".xlsx")
