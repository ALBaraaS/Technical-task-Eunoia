#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pandas as pd
import numpy as np
from datetime import datetime


# In[2]:


TestData=pd.ExcelFile('Test Data.xlsx')
dfsales=pd.read_excel(TestData, 'Sales')
dfproduct=pd.read_excel(TestData, 'Product Info')
dfcomm=pd.read_excel(TestData, 'Agent Commission in %')


# In[3]:


#dfsales


# In[4]:


#dfproduct


# In[5]:


#dfcomm


# In[6]:


dfcommunpiv=pd.melt(dfcomm,id_vars=['Agent Code', 'Valid Date From'], var_name='Category', value_name='Perc')


# In[7]:


#dfcommunpiv.sort_values(by=['Agent Code'])


# In[8]:


#adc=dfcommunpiv.groupby(by='Agent Code')


# In[9]:


#adc


# In[10]:


#dfsales.loc[[0]]


# In[11]:


#dfsales['Price'].loc[[0]]


# In[12]:


#dfsales['Price'].loc[0]


# In[13]:


dfsales['Product Category']=""
dfsales['Commission']=""
dfsales['Revenue']=""


# In[14]:


x=dfsales.count()[0]
y=dfproduct.count()[0]
z=dfcommunpiv.count()[0]
for i in range(x):
    for j in range(y):
        if dfsales['Product Code'].loc[i]==dfproduct['Product Code'].loc[j]:
            if dfproduct['From Date'].loc[j]<=dfsales['Transaction Date'].loc[i] and dfproduct['To Date'].loc[j]>dfsales['Transaction Date'].loc[i]:
                dfsales['Product Category'].loc[i]=dfproduct['Product Category'].loc[j]
                break
    for k in range(z):
        if dfsales['Agent Code'].loc[i]==dfcommunpiv['Agent Code'].loc[k]:
            if dfsales['Product Category'].loc[i]==dfcommunpiv['Category'].loc[k]:
                if dfsales['Transaction Date'].loc[i]>=dfcommunpiv['Valid Date From'].loc[k]:
                    dfsales['Commission'].loc[i]=dfcommunpiv['Perc'].loc[k]/100


# In[15]:


dfsales['Revenue']=dfsales['Commission']*dfsales['Price']


# In[16]:


#dfsales


# In[17]:


dfsales.groupby('Agent Code')['Revenue'].agg(['sum', 'count'])


# In[18]:


dfsales.groupby('Product Category')['Revenue'].agg(['sum', 'count'])


# In[19]:


#dfsales['Pno']=dfsales['Product Code'].str[-3:].astype(int)
#dfproduct['Pno']=dfproduct['Product Code'].str[-3:].astype(int)
#dfsales['Ano']=dfsales['Agent Code'].str[-3:].astype(int)
#dfcommunpiv['Ano']=dfcommunpiv['Agent Code'].str[-3:].astype(int)


# In[20]:


#dfsales['Product Code']=dfsales['Product Code'].str[-3:]


# In[21]:


#dfsales['Product Code']=dfsales['Product Code'].astype(int)


# In[22]:


pd.pivot_table(dfsales, values='Revenue', index='Agent Code', columns='Product Category', aggfunc=np.sum, margins=True)


# In[23]:


pd.pivot_table(dfsales, values='Revenue', index='Agent Code', columns='Product Category', aggfunc='count', margins=True)


# In[24]:


#dfproduct.set_index(['Product Code', 'From Date', 'To Date'])
#dfproduct_dict = dfproduct.to_dict("index")


# In[25]:


#dfproduct_dict


# In[26]:


#dfproduct


# In[27]:


#x=dfsales.count()[0]
#y=dfproduct.count()[0]
#z=dfcommunpiv.count()[0]
#for i in range(x):
#    for Code, Cat, From, To in dfproduct_dict.items():
#        if dfsales['Product Code'].loc[i]==Code:
#            if dfsales['Transaction Date'].loc[i]>=From and dfsales['Transaction Date'].loc[i]<=To:
#                dfsales['Product Category'].loc[i]=Cat
#                break


# In[28]:


########                               From here a new solutio                          ################


# In[29]:


#trial1=dfsales.merge(dfproduct, how='cross')


# In[30]:


#trial1


# In[31]:


#trial1=trial1.rename(columns={"Product Code_x":"ProductCode_x", "Product Code_y":"ProductCode_y"})
#trial1=trial1.rename(columns={"Transaction Date":"TransactionDate","From Date":"FromDate","To Date":"ToDate"})


# In[32]:


#con1="(ProductCode_x==ProductCode_y)"


# In[33]:


#con2="(FromDate <= TransactionDate < ToDate)"


# In[34]:


#qry=con1 + " & " + con2


# In[35]:


#trial1=trial1.query(qry)


# In[36]:


#trial1


# In[37]:


#dfsales


# In[38]:


#trial11=trial1.drop(['ProductCode_y', 'FromDate', 'ToDate', 'ProductCode_x'], axis=1)


# In[39]:


#trial2=trial11.merge(dfcommunpiv, how='cross')


# In[40]:


#trial2


# In[41]:


#trial2=trial2.rename(columns= {"Agent Code_x":"AgentCode_x","Agent Code_y":"AgentCode_y","Valid Date From":"ValidDateFrom","Product Category":"ProductCategory"})


# In[42]:


#trial2=trial2.query("AgentCode_x == AgentCode_y")


# In[43]:


#trial2=trial2.query("TransactionDate>=ValidDateFrom")


# In[44]:


#trial2=trial2.query("ProductCategory==Category")


# In[45]:


#trial2.sort_values(by=['Transaction ID'])


# In[46]:


#trial2=trial2.drop_duplicates(subset=['Transaction ID'], keep='last')


# In[47]:


#trial2=trial2.drop(['AgentCode_y', 'ValidDateFrom', 'Category'], axis=1)


# In[ ]:




