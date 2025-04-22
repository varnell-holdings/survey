# -*- coding: utf-8 -*-
"""
Created on Fri Apr  4 09:53:36 2025

@author: John2
"""
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)
# mail.To = 'tillett1957@gmail.com'
mail.To = 'matt@endoscopy.stvincents.com.au'

mail.Subject = 'Hello'
mail.HTMLBody = "testing survey!"

# mail.Send()
mail.Display()
