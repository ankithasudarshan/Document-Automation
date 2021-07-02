#!/usr/bin/env python
# coding: utf-8

# In[18]:


from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

template = "Formal-job-offer-letter-template-1.docx"


# In[19]:


document = MailMerge(template)
print(document.get_merge_fields())


# In[20]:


document.merge(
    GrossSalary='10L',
    NoOfHours='40',
    NoOfShares='100',
    Leaves='27',
    NoOfMonths='6',
    FullOrPartTime='Full',
    ManagersTitle='Mr Ram',
    WorkingDays='5',
    StartDate='10-03-2021',
    PayCheckDate='10-04-2021',
    TodaysDate='{:%d-%b-%Y}'.format(date.today()),
    PercentBonus='20%',
    JobTitle='Data Scientist',
    CompanyName='ABC Pvt Ltd',
    Sendersname='Tom Martinez')


# In[21]:


document.write('test-output.docx')


# In[23]:




import os
cwd = os.getcwd()
print(cwd)


# In[ ]:




