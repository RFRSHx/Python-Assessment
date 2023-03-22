#!/usr/bin/env python
# coding: utf-8

# In[51]:


import pandas as pd
import datetime


# In[52]:


#Import the 3 excels
df1 = pd.read_excel('Option_Dataset_1.xlsx')
df2 = pd.read_excel('Option_Dataset_2.xlsx')
df3 = pd.read_excel('Option_Dataset_3.xlsx')


# In[53]:


#Aggregate data
df = pd.concat([df1,df2,df3])


# In[54]:


#Create Moneyness column
df["Moneyness"] = df["Spot Price"]/df["Strike Price"]


# In[55]:


#Update Implied Volatility
df["Implied Volatility"] = df["Implied Volatility"] / 100


# In[56]:


#Compute time in years

#Get current time
current_time = datetime.datetime.now()

#Convert Expiry Date data to datetime format
df['Expiry Date'] = pd.to_datetime(df['Expiry Date'], errors='coerce')

#Get the difference in seconds and transform it in years
df['Expiry Date(years)'] = (df['Expiry Date'] - current_time).dt.total_seconds()/31536000


# In[57]:


#Create Full Identifier
df["Full Identifier"] = df["Equity Ticker"] + "-" +df["Company"]


# In[58]:


#Sort by country
df.sort_values(by='Country', inplace=True)


# In[59]:


#Export in Option_Panel file
df.to_excel('Option_Panel.xlsx', index=False)

