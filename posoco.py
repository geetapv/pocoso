# -*- coding: utf-8 -*-
"""
Created on Fri Feb 17 21:47:07 2023

@author: GEETA
"""


import urllib.request
import os
from selenium import webdriver
import PyPDF2
import tabula      
from unicodedata import normalize
import re
import logging

data_dir=os.getcwd()

logging.basicConfig(filename=os.path.join(data_dir,'posoco.log'),level=logging.DEBUG,
                    format="%(asctime)s:%(levelname)s:%(message)s")


def find_pdf():
    while True:
        try:
            print("starting the process")
            logging.info("starting the process")
            driver = webdriver.Chrome(r"C:\Users\GEETA\Desktop\chromedriver.exe")
            url="https://posoco.in/"
            driver.get(url)
            driver.implicitly_wait(10)
            print("working")
            p1=driver.find_element_by_link_text('Transmission Pricing')
            driver.implicitly_wait(10)
            p1.click()          
            print("page found")
            p2=driver.find_element_by_link_text('Notification of Transmission Charges for the DICs')
            driver.implicitly_wait(10)
            p2.click()            
            print('inside other page')
            table=driver.find_element_by_tag_name("table")
            print(table)
            data = [item.get_attribute('href') for item in table.find_elements_by_tag_name("a")]
            u=data[0]   #taking the first link as it is the latest one
            logging.info("Got the latest pdf available")
            print("Got the latest pdf available")
            print(u)
            urllib.request.urlretrieve(u, os.path.join(data_dir,"notification_transmission_charges.pdf"))
            print('downloaded it in the path')
            logging.info('downloaded it in the path')
            driver.close()
            break
        except Exception as e:
            print(e)
            driver.close()

def finding_page_no(S):
    fi=PyPDF2.PdfReader(os.path.join(data_dir,"notification_transmission_charges.pdf"))
    pg_count =len(fi.pages)
    for i in range(0, pg_count):
        PgOb = fi.pages[i]
        data= PgOb.extract_text()
        newd=normalize('NFKD', data)
        newd=re.sub(' +', ' ', newd)
        if S in newd:
             print("String Found on Page: " + str(i))
             print(f'page number is {i+1}')
             logging.info("String Found on Page: " + str(i))
             logging.info(f'page number is {i+1}')
             return i+1
         



def formatting_df(p_no):   
    df_list=[] 
    df=tabula.read_pdf(os.path.join(data_dir,"notification_transmission_charges.pdf"),pages=p_no)
    df=df[0]
    df_list.append(df)
    last_no=df['S.No.'].values[-1]
    print(last_no)
    while True:
        try:
            p_no=p_no+1
            df=tabula.read_pdf(os.path.join(data_dir,"notification_transmission_charges.pdf"),pages=p_no)
            df=df[0]
            f_no=df['S.No.'].values[1]
            print(f_no)
            if f_no>last_no:
               print("new page")
               df_list.append(df)
               last_no=df['S.No.'].values[-1]
        except Exception as e:
              print(e)
              break
    print("fetched all the data of the table") 
    logging.info("fetched all the data of the table")    
    print("formatting the data")
    logging.info("formatting the data")
    df=df_list[0]      
    df=df.fillna('')
    x=df.iloc[0]
    x=x.tolist()
    x=x[:-1]
    x.insert(3,'')
    x=[i.replace('?','-') for i in x]
    df.iloc[0]=x
    cols=[i.replace('\r',' ') for i in df.columns]
    cols=[i.replace('?',"₹") for i in cols]
    df.columns=cols
    c=(df.columns[6:]).tolist()
    df.rename(columns={c[i+1]:c[i] for i in range(len(c)-1)},inplace=True)
    if len(df_list)>1:
        for d in df_list:
            d=d.iloc[1:,:]
            d=d.fillna("")
            d.columns=df.columns
            df=df.append(d)      #combining the data into one dataframe
    print("Saving data to excel file")
    logging.info('saving data to excel files')
    df.to_excel(os.path.join(data_dir,"Transmission_charges_DIC.xlsx"),index=False)
    

def main():
    S ="Transmission Charges for Designated ISTS Customers (DICs)"
    find_pdf()
    p=finding_page_no(S)
    formatting_df(p)
    
    
if __name__=="__main__":
    main()

