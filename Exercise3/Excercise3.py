#!/usr/bin/env python
# coding: utf-8

# In[8]:


import pandas as pd
import numpy as np
from scipy.stats import norm


# In[7]:


#Read the excel file
df = pd.read_excel("Option_Panel.xlsx")


# In[9]:


# Define the Black-Scholes formula
def black_scholes(s, k, t, sigma, r, q, flag):
    
    
    d1 = (np.log(s/k) + (r - q + sigma**2/2)*t) / (sigma*np.sqrt(t))
    d2 = d1 - sigma*np.sqrt(t)
    if flag == "C":
        
        #Compute Call flag
        return s*np.exp(-q*t)*norm.cdf(d1) - k*np.exp(-r*t)*norm.cdf(d2)
    else:
        
        #Compute Put flag
        return k*np.exp(-r*t)*norm.cdf(-d2) - s*np.exp(-q*t)*norm.cdf(-d1)


# In[10]:


#Compute option price
df["Option Price"] = df.apply(lambda x: black_scholes(x["Spot Price"], x["Strike Price"], x["Expiry Date(years)"], x["Implied Volatility"],0.02,0,x["Flag"]), axis=1)


# In[11]:


#Compute with 10 contracts per Option Price
df["Option Price"] = df["Option Price"] * 10


# In[12]:


# Group the market value of the portfolio by country
table1 = df.groupby("Country")["Option Price"].sum()

# Group the market value of the portfolio by flag
table2 = df.groupby("Flag")["Option Price"].sum()


# In[13]:


#Export in new excel file
with pd.ExcelWriter("Result_Ex3.xlsx") as writer:
    table1.to_excel(writer, sheet_name="By Country")
    table2.to_excel(writer, sheet_name="By Flag")

