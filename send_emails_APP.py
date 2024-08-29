# -*- coding: utf-8 -*-
"""
Created on Thu Apr  4 09:45:27 2024

@author: ashehata
"""
### all imports 

import pandas   as    pd 
import numpy    as    np
from PyQt6.QtWidgets import QApplication , QMainWindow ,QLabel ,QPushButton ,QVBoxLayout , QWidget, QHBoxLayout ,QGridLayout ,QLineEdit ,QComboBox
from PyQt6.QtGui     import QIcon , QPixmap
from PyQt6.QtCore    import Qt
import os
import win32com.client
import time 


class Window(QMainWindow): 
    # our class is inheriting from the QMain window 
    def __init__(self):
        super().__init__()   # to initiate the super function 
        self.setMinimumSize(600,600)  # this is to set minimumSize 
        self.setWindowTitle(" Alten Buinees manager application")
        self.setWindowIcon(QIcon("Alten_Wall_Paper.png"))
    
        parentLayout=QGridLayout()  # to define the parent layout 
        self.label=QLabel()
        self.button = QPushButton("Submit your information")  # you can add buttons , labels and more 
        self.KeywordLabel =QLabel("Insert your key words in 'keyword1' ,'keyword2',etc......")
        self.excelfileLabel =QLabel("Insert the name of your excel file")
        self.ColumnCompetenceLabel =QLabel("insert the name of the column of the competence")
        self.keywordInput=QLineEdit()
        self.excelfileInput=QLineEdit()
        self.ColumnCompetenceInput=QLineEdit()
        self.label.setPixmap(QPixmap("Alten.jpg"))
        self.button.clicked.connect(self.clickHandler)
######################################################################### Labels and Inputs ###########################################3
        parentLayout.addWidget(self.button)    # then you can start adding widgets ,this widgets could be button ,labels and more 
        # Keyword Inputs

        parentLayout.addWidget(self.KeywordLabel,1,0)
        parentLayout.addWidget(self.keywordInput,1,1)

        # Excel file Inputs
        parentLayout.addWidget(self.excelfileLabel ,2,0)
        parentLayout.addWidget(self.excelfileInput,2,1)

        # ColumnCompetenceInput
        parentLayout.addWidget(self.ColumnCompetenceLabel,3,0)
        parentLayout.addWidget(self.ColumnCompetenceInput,3,1)
        
        parentLayout.addWidget(self.label)
        
        centreWidget=QWidget()
        centreWidget.setLayout(parentLayout)
        self.setCentralWidget(centreWidget)
        
    def clickHandler(self):
       print("Informtion Submitted")
       key_words       = self.keywordInput.text()
       word=''
       key_word_lists=[]
       for idx,letter in enumerate(key_words):
           if letter ==',' or idx == len(key_words)-1:
               key_word_lists.append(word)
               word=''
           word+=letter
       excel_file      = self.excelfileInput.text()
       competence_column = self.ColumnCompetenceInput.text()
       df=pd.read_excel(excel_file,sheet_name='All old and new')
       competences=df[competence_column]
       index=0
       indeces_list=[]
     # Competence is the competence of the manager and the area where he operates 
       a=1
       for competence in competences:
           if type(competence) ==str:
               # we need to change all the letters into lower cases 
               competence=competence.lower()
               for key_word in key_word_lists:
                   a+=1
                   key_word= key_word.lower()
                   if key_word in competence:
                       indeces_list.append(index)
                   
           index+=1 
       indeces_list=set(indeces_list)
       indecies=list(indeces_list)
       df2=pd.read_excel(excel_file)
       df2=df2.iloc[indecies]
       template_df=pd.read_excel(excel_file,sheet_name='Email Template')
       managers_dictionary =pd.DataFrame({'managers_names': df2['Manager'],'managers_email' : df2['E-mail'],'phone' :df2['Phone'],'BM SW' :df2['BM SW'] ,'Template' :template_df['Tidigare kund (Eng)'][2]})
       managers_dictionary.to_excel("output.xlsx")
       
       ###################################### email automate sending ####################################################################################################################
       ############### I can add another layer in the application with a message to ask if you want to send emails #########################

       emails=managers_dictionary['managers_email']
       ol = win32com.client.Dispatch('Outlook.Application').GetNamespace("MAPI")

       newmail = ol.OpenSharedItem(r"C:\Users\ashehata\Desktop\seka.msg")
       Mail=win32com.client.Dispatch('Outlook.Application').CreateItem(0x0)
       Mail.Subject = 'we have a new candidate'
       # the subject of the email
       
       for idx,email in enumerate(emails) :
           name=''
           i=0
           for letter in email:
               if i==0:
                   letter=letter.upper()
               if letter != '.':
                   name+=letter
               else:
                   break
               i+=1
           print(idx ,name)
           Mail.To=email
           Mail.HTMLbody ='Hi '+name +newmail.HTMLbody                             # it has a small bug here still 
           Mail.Display()
           time.sleep(4)
          
      

app =QApplication([])
window=Window()
window.show()
app.exec()
        
