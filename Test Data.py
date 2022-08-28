#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from datetime import datetime


# In[2]:


#Read from Excel
TestData=pd.ExcelFile('Test Data.xlsx')
dfsales=pd.read_excel(TestData, 'Sales')
dfproduct=pd.read_excel(TestData, 'Product Info')
dfcomm=pd.read_excel(TestData, 'Agent Commission in %')


# In[3]:


#Unpivot Commission Sheet 
dfcommunpiv=pd.melt(dfcomm,id_vars=['Agent Code', 'Valid Date From'], var_name='Category', value_name='Perc')


# In[4]:


#Merging Sales with Product Info
dfresults=dfsales.merge(dfproduct, how='cross')


# In[5]:


dfresults=dfresults.rename(columns={"Product Code_x":"ProductCode_x", "Product Code_y":"ProductCode_y"})
dfresults=dfresults.rename(columns={"Transaction Date":"TransactionDate","From Date":"FromDate","To Date":"ToDate"})


# In[6]:


#Cleaning the Data
con1="(ProductCode_x==ProductCode_y)"
con2="(FromDate <= TransactionDate < ToDate)"
qry=con1 + " & " + con2
dfresults=dfresults.query(qry)


# In[7]:


#Drop Unwanted Columns
dfresults=dfresults.drop(['ProductCode_y', 'FromDate', 'ToDate', 'ProductCode_x'], axis=1)


# In[8]:


#Merge with Agent Commission
dfresults=dfresults.merge(dfcommunpiv, how='cross')


# In[9]:


dfresults=dfresults.rename(columns= {"Agent Code_x":"AgentCode_x","Agent Code_y":"AgentCode_y","Valid Date From":"ValidDateFrom","Product Category":"ProductCategory"})


# In[10]:


#Cleaning the Data
dfresults=dfresults.query("AgentCode_x == AgentCode_y")


# In[11]:


dfresults=dfresults.query("TransactionDate>=ValidDateFrom")


# In[12]:


dfresults=dfresults.query("ProductCategory==Category")


# In[13]:


dfresults=dfresults.drop_duplicates(subset=['Transaction ID'], keep='last')


# In[14]:


#Drop Unwanted Columns
dfresults=dfresults.drop(['AgentCode_y', 'ValidDateFrom', 'Category'], axis=1)


# In[15]:


dfresults=dfresults.rename(columns={"AgentCode_x":"Agent Code", "TransactionDate":"Transaction Date","ProductCategory":"Product Category"})


# In[16]:


#Finding the Commission
dfresults['Revenue']=dfresults['Perc']*dfresults['Price']/100


# In[17]:


#Commission per Agent
print("Commission per Agent\n")
print(dfresults.groupby('Agent Code')['Revenue'].agg(['sum']).rename(columns={'sum':'Commission per Agent'}))


# In[18]:


#Commission per Product Category
print("\n Commission per Product Category \n")
print(dfresults.groupby('Product Category')['Revenue'].agg(['sum']).rename(columns={'sum':'Commission per Product Category'}))


# In[19]:


#Number of Sales per Agent
print("\n Number of Sales per Agent \n")
print(dfresults.groupby('Agent Code')['Revenue'].agg(['count']).rename(columns={'count':'Number of Sales per Agent'}))


# In[20]:


#Number of Sales per Product
print("\n Number of Sales per Product \n")
print(dfresults.groupby('Product Category')['Revenue'].agg(['count']).rename(columns={'count':'Number of Sales per Product'}))


# In[21]:


#A matrix of the number of sales per agent per product category. (Agent at the row level, product category at the column level.)
print("\n Number of Sales per Agent per Product Category \n")
print(pd.pivot_table(dfresults, values='Revenue', margins_name='Total number sold', index='Agent Code', columns='Product Category', aggfunc='count', margins=True))


# In[22]:


#A matrix of the commission value sales per agent per product category. (Agent at the row level, product category at the column level.)
print("\n Commission Value Sales per Agent per Product Category \n")
print(pd.pivot_table(dfresults, values='Revenue', margins_name='Total Comm.', index='Agent Code', columns='Product Category', aggfunc=np.sum, margins=True))


# In[23]:


print("")
input("Press Enter to Exit")


# In[ ]:




