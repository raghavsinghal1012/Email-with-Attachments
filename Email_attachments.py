import os
import smtplib
import imghdr
from email.message import EmailMessage
from PyPDF2 import PdfFileReader, PdfFileWriter
import pandas as pd
from tkinter import filedialog
from tkinter import *
from pdf2image import convert_from_path
from goto import with_goto
import win32com.client as client
import pathlib


excel_path=""
path_of_pdf_files=""
pdf=""
exdata=""
name=""
mail=""
number=""
EMAIL_ADDRESS = os.environ("my_email")
EMAIL_PASSWORD = os.environ("my_pass")




def columncheck(l):

    print("printing column names of your excel file")
    print()

    for i in l:
        print(i)
    print()
    a=input("Enter name of column consisting name of students:")
    b=input("Enter name of column consisting registeration number of students:")
    c=input("Enter name of column consisting emails of students:")
    if((a in l)and(b in l)and(c in l)):
        return [a,b,c]
    else:
        print("One of the column name is incorrect")
        print("please enter again")
        var=columncheck(l)
        return var

    
def excel_path1():
    global excel_path
    excel_path=filedialog.askopenfilename(filetypes=[("Excel files","*.xlsx"),("Excel files","*.csv"),("Excel files", "*.xls")])



def path_of_pdf_files1():
    global path_of_pdf_files
    path_of_pdf_files=filedialog.askdirectory()


@with_goto
def get_excel():
    label.start
    root1=Tk()
    root1.geometry('900x200')
    
    excel_button=Button(root1,text="SELECT EXCEL FILE",bg="yellow", fg="blue",font=("Arial Bold", 15),command=excel_path1)
    excel_button.grid(row = 2, column = 0)
    exit_button=Button(root1,text="NEXT",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root1.destroy)
    exit_button.grid(row = 2, column = 2)
    root1.mainloop()
    if (excel_path==""):
        goto.start
        
@with_goto
def start_func():
    get_excel()
    
    
    

     
    
    
def pdf_convert():
    global exdata
    global pdf
    global exdata
    global mail
    global name
    global number
    global path_of_pdf_files
    global excel_path
    global EMAIL_ADDRESS
    global EMAIL_PASSWORD
    #with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
     #   smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
    for i in range(exdata.shape[0]):
        val=exdata[name][i]
        val1=exdata[number][i]
        val2=exdata[mail][i]
        outlook=client.Dispatch('Outlook.Application')
        msg=outlook.CreateItem(0)
        msg.To=val2
        msg.Cc="raghav.singhal200110@gmail.com"
        msg.Subject='Certificate for Enrolment ID '+val1+' '+val
        msg.Body='hi '+val+'\n\n'+'please find your attachment\n\n'+'thank you'
        msg.Attachments.Add(path_of_pdf_files+'/'+val1+'_'+val+'.pdf')
        msg.Send()
        '''msg = EmailMessage()
        msg['Subject'] = 'Certificate for Enrolment ID '+val1+' '+val
        msg['From'] = EMAIL_ADDRESS
        msg['To'] = [val2,'raghav.singhal200110@gmail.com']                  
        msg.set_content('hi '+val+'\n\n'+'please find your attachment\n\n'+'thank you')
        with open(path_of_pdf_files+'/'+val1+'_'+val+'.pdf','rb') as f:
        file_data=f.read()
        msg.add_attachment(file_data,maintype='application',subtype='pdf',filename=val1+'_'+val)

        smtp.send_message(msg)'''


            
        '''with open(os.path.join(output_path,'{0}_{1}.pdf'.format(val1,val)), 'wb') as f:
            pdfw.write(f)
            f.close()'''


    
@with_goto
def convertfile():
    global exdata
    global pdf
    global name
    global number
    global path_of_pdf_files
    global mail
    global EMAIL_ADDRESS
    global EMAIL_PASSWORD
    exdata=pd.read_excel(excel_path)
    column_list=list(exdata.columns)
    column_list=columncheck(column_list)
    name=column_list[0]
    number=column_list[1]
    mail=column_list[2]
    label.start3
    root3=Tk()
    root3.geometry('900x200')
    output_button=Button(root3,text="SELECT FOLDER IN WHICH PDF ARE SAVED",bg="yellow", fg="blue",font=("Arial Bold", 15),command=path_of_pdf_files1)
    output_button.grid(row = 2, column = 0)
    exit_button=Button(root3,text="NEXT",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root3.destroy)
    exit_button.grid(row = 2, column = 2)
    #back_button=Button(root3,text="BACK",bg="yellow", fg="blue",font=("Arial Bold", 15),command=get_pdf)
    #back_button.grid(row = 2, column = 4)
    root3.mainloop()
    if(path_of_pdf_files==""):
        goto.start3

    
    root4=Tk()
    root4.geometry('900x200')
    PDF_button=Button(root4,text="CLICK TO SEND ALL EMAILS",bg="yellow", fg="blue",font=("Arial Bold", 15),command=pdf_convert)
    PDF_button.grid(row = 2, column = 0)
    #JPG_button=Button(root4,text="CLICK TO CONVERT IN JPG",bg="yellow", fg="blue",font=("Arial Bold", 15),command=jpg_convert)
    #JPG_button.grid(row = 2, column = 2)
    exit_button=Button(root4,text="Exit",bg="yellow", fg="blue",font=("Arial Bold", 15),command=root4.destroy)
    exit_button.grid(row = 4, column = 0)
    root4.mainloop()
    





    



if __name__=="__main__":
    start_func()
    convertfile()
