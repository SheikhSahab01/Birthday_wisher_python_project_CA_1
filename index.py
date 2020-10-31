import pandas as pd 
import datetime
import smtplib
from selenium import webdriver
from getpass import getpass
import time
from tkinter import *

root = Tk()
root.geometry("400x600")
root.title("Birthday wisher")

Label(root,text="Birthday Wisher",font="Montserrat 16 bold").pack(side=TOP,pady=20)
ent1=StringVar()
ent2=StringVar()
ent3=StringVar()
ent4=StringVar()

Label(root,text="Enter your Email:",font="Montserrat 12 bold").pack(pady=(100,0))
entry1=Entry(root,textvariable=ent1)
entry1.pack(pady=5,ipady=7,ipadx=15)

Label(root,text="Password:",font="Montserrat 12 bold").pack()
entry2=Entry(root,textvariable=ent2, show="*")
entry2.pack(pady=5,ipady=7,ipadx=15)

Label(root,text="Enter your instagram user name :",font="Montserrat 12 bold").pack()
entry3=Entry(root,textvariable=ent3)
entry3.pack(pady=5,ipady=7,ipadx=15)

Label(root,text="Password:",font="Montserrat 12 bold").pack()
entry4=Entry(root,textvariable=ent4, show="*")
entry4.pack(pady=5,ipady=7,ipadx=15)

def fun1():
    inuser = entry3.get()
    ipassword = entry4.get()
    GMAIL_ID = entry1.get()
    GMAIL_PSWD = entry2.get()

    print(GMAIL_ID)
    print(inuser)

    ent1.set('')
    ent2.set('')
    ent3.set('')
    ent4.set('')

    
    def sendEmail(to,sub, msg):
        print(f"Email to {to} sent with subject: {sub} and message {msg}")
        s = smtplib.SMTP('smtp.gmail.com' , 587)
        s.starttls()
        s.login(GMAIL_ID, GMAIL_PSWD)
        s.sendmail(GMAIL_ID, to, f"subject:{sub}\n\n{msg}")
        s.quit()

    if __name__ == "__main__":
        df = pd.read_excel("data.xlsx")
        # print(df)
        today = datetime.datetime.now().strftime("%d-%m")
        yearnow = datetime.datetime.now().strftime("%Y")
        # print(today)
        # print(type(today))
        writeInd = []
        for index, item in df.iterrows():
            # print(index, item['Birthday'])
            bday = item['Birthday'].strftime("%d-%m")
            # print(bday)
            if (today == bday) and yearnow not in str(item['year']):
                sendEmail(item['Email'], "happy birthday "+ item['Name'] ,item['Dialouge'])
                writeInd.append(index)

                print(item['Name'] + " is having birthday today")
                
                            
                driver = webdriver.Chrome()
                driver.get("https://www.instagram.com/accounts/login")

                driver.maximize_window()
                msg_box1 = driver.find_element_by_xpath("//*[@id='loginForm']/div/div[1]/div/label/input")
                msg_box1.send_keys(inuser)

                msg_box2 = driver.find_element_by_xpath("//*[@id='loginForm']/div/div[2]/div/label/input")
                msg_box2.send_keys(ipassword)

                login = driver.find_element_by_css_selector('button[type=submit]')
                login.click()

                time.sleep(5)

                popup1 = driver.find_element_by_xpath("//*[@id='react-root']/section/main/div/div/div/div/button")
                popup1.click()

                time.sleep(5)

                popup2 = driver.find_element_by_xpath("/html/body/div[4]/div/div/div/div[3]/button[2]")
                popup2.click()

                time.sleep(2)

                search = driver.find_element_by_xpath("//*[@id='react-root']/section/nav/div[2]/div/div/div[2]/div/div").click()
                time.sleep(3)

                search1 = driver.find_element_by_xpath("//*[@id='react-root']/section/nav/div[2]/div/div/div[2]/input")
                search1.send_keys(item['iuser'])

                time.sleep(3)

                senderselector = driver.find_element_by_xpath("//*[@id='react-root']/section/nav/div[2]/div/div/div[2]/div[3]/div[2]/div/a[1]")
                senderselector.click()

                time.sleep(5)

                msgbtn = driver.find_element_by_xpath("//*[@id='react-root']/section/main/div/header/section/div[1]/div[1]/div/div[1]/div/button")
                msgbtn.click()
            
                time.sleep(5)

                msginpt = driver.find_element_by_xpath("//*[@id='react-root']/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[2]/textarea")
                msginpt.send_keys(item['Dialouge'])

                time.sleep(3)

                sendmsg = driver.find_element_by_xpath("//*[@id='react-root']/section/div/div[2]/div/div/div[2]/div[2]/div/div[2]/div/div/div[3]/button")
                sendmsg.click()
                time.sleep(2)

                profile = driver.find_element_by_xpath("//*[@id='react-root']/section/div/div[1]/div/div[3]/div/div[5]/span/img")
                profile.click()
                time.sleep(2)

                profile1 = driver.find_element_by_xpath("//*[@id='react-root']/section/div/div[1]/div/div[3]/div/div[5]/div[2]/div/div[2]/div[2]/a[1]/div")
                profile1.click()
                time.sleep(2)

                setting = driver.find_element_by_xpath("//*[@id='react-root']/section/main/div/header/section/div[1]/div[2]")
                setting.click()
                time.sleep(2)

                logout = driver.find_element_by_xpath("/html/body/div[5]/div/div/div/div/button[9]")
                logout.click()

                driver.close()
        
        # print(writeInd)
        for i in writeInd:
            yr = df.loc[i, 'year']
            # print(yr )
            df.loc[i, 'year'] = str(yr) + ',' + str(yearnow)
            # print(df.loc[i, 'year'] )

        # print(df)
        df.to_excel('data.xlsx', index = False)

  
Button(root,text="Submit",command=fun1).pack(pady=10)
root.mainloop()


