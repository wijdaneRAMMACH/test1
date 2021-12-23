#!/usr/bin/env python
# coding: utf-8

# In[149]:


import pandas as pd
import os

#pip install libconf
import libconf
import io

import psycopg2
#pip install psycopg2

import logging
import yaml
import logging.config

from datetime import datetime

#pip install pyftpdlib
from ftplib import FTP 

import win32com.client as win32


# In[150]:


TODAY = datetime.today().strftime('%Y/%m/%d')
logging.config.dictConfig(yaml.load(open('CONFIG/logging.yaml', 'r')))


# In[151]:


def read_conf():
    with io.open('CONFIG/config.cfg', encoding='utf-8') as f:
        cfg = libconf.load(f)
    return cfg
    
def init_connection(host_pg,user_pg,pwd_pg,port_pg):
    """
    This function initialize the postgre connection
    return : a pg connection
    """
    try:
        conn = psycopg2.connect(host=host_pg, port = port_pg, database="postgres", user= user_pg, password= pwd_pg)
        logging.info(" Connected succefully to postgre database"  )
        return conn
    except:
        logging.error(" There is an issue when connecting. Check user / password  ")
        
        
def read_table(conn):
    """
    This function read data from souscritootest 

    param conn: connexion to the source database 
    return : a panda dataframe containing the table  T_EDW_TGM
    """
    try:
        sql = "SELECT * FROM clients_crm ; "
        pd_data = pd.read_sql(sql, conn)
        pd_data['phone'] = pd_data['phone'].astype(str).str.replace('\.0', '').str.rjust(10,'0')
        logging.info(" data from pg table has been loaded successfully !")
        if(len(pd_data)==0): logging.warn("clients_crm table is empty ")
        return pd_data
        
    except:
        logging.error(" There was an issue with loading pg data  ")


def read_csv_ftp(path):
    try:
        df_file = pd.read_csv(path) 
        df_file = df_file.rename(columns={'incoming_number': 'phone'})
        df_file['phone'] = df_file['phone'].map(str)
        df_file.drop(df_file.columns[[-1]], axis=1, inplace=True)

        logging.info("raw_calls file has been successfully cleaned up!")
        return df_file
    except:
        logging.error("There is an issue with the raw_calls file")

def combine_data(df_table, df_file):
    try:
        result= df_table.join(df_file, lsuffix='_l', rsuffix='_r')
        #result= result.drop(result.columns[[0,1]], axis=1, inplace=True)
        logging.info("data has been successfully combined!")
        return result 
    except:
        logging.error("There is an issue with the data prepare !")
        
        
def send_mail( addr_to , data , file ):
    try:
        outlook=win32.Dispatch('outlook.application')
        mail=outlook.CreateItem(0)
        mail.To= addr_to
        mail.Subject='papernest [client Data]'
        mail.Body='Message body \n'+ data
        mail.HTMLBody='<p> Hello, <br><br>Please find below the list of clients: <br> <br> </p>' + data +                  '<p> Also attached is the list of customer calls.<br><br><br> </p>                       <p> Best regards. </p>'
 
        attachment= os.getcwd() +"\\" + file
        mail.Attachments.Add(attachment)
        mail.Send()
        logging.info("data has been successfully send!")
    except:
        logging.error("There is an issue with mail send !")
    
def papernest():
    """
    This function ....
    param dsc: ....
    param user: ....
    return password: ....
    """
    
    cfg = read_conf()
      
    # init connexion
    conn = init_connection(cfg.host_pg, cfg.user_pg, cfg.pwd_pg, cfg.port_pg)
    
    # pg table
    data_table = read_table(conn)
        
    #Read csv file from ftp
    url_ftp= "ftp://"+ cfg.user_ftp + ":"+ cfg.pwd_ftp + "@" + cfg.host_ftp +"/" +cfg.directory +"/"+ cfg.file_path
    data_file = read_csv_ftp(url_ftp)
    
    # new table    combine data first contact
    result= combine_data(data_table , data_file )
    result.to_csv( cfg.output_file , header = True , index = False)
    
    # send mail
    send_mail( cfg.addr_to , data_table.to_string() , cfg.output_file  )
    


# In[152]:


# run
logging.info('***** papernest APPLICATION ***** ' + TODAY)
papernest()


# ### testing  

# In[66]:


'Hello, <br> Please find below the list of clients: <br> <br> ' + data                 + ' <br><br><br> Also attached is the list of customer calls.<br><br><br> Best regards. \n'


# In[154]:


cfg = read_conf()
      
conn = init_connection(cfg.host_pg, cfg.user_pg, cfg.pwd_pg, cfg.port_pg)

data_table = read_table(conn)


# In[155]:


data_table.head()


# In[ ]:




