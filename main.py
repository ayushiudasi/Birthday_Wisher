import pandas as pd
import datetime
import smtplib
# write your authentication details
GMAIL_ID=''
GMAIL_PSWD=''

#function to send Email
def sendmail(to ,sub,msg):
   s = smtplib.SMTP('smtp.gmail.com',587)
   s.starttls()
   s.login(GMAIL_ID,GMAIL_PSWD)
   s.sendmail(GMAIL_ID,to,f"Subject: {sub}\n\n{msg}")
   s.quit()
if __name__ == '__main__':
    df = pd.read_excel("data.xlsx")
    today = datetime.datetime.now().strftime("%d-%m") # using today's date in format (dd-mm)
    yearnow= datetime.datetime.now().strftime("%Y") #using present year
    writeInd=[] #list to store index of people who received Email
    for index,item in df.iterrows():
        bday=item['Birthday'].strftime("%d-%m")

        if (today==bday) and yearnow not in str(item['Year']):
            sendmail(item['Email'],"Happy Birthday",item['Dialogue'])
            writeInd.append(index)

    if (len(writeInd)!=0):
        for i in writeInd:
            yr=df.loc[i,'Year']
            df.loc[i,'Year']=str(yr)+', '+str(yearnow)# updating list so that person do not get Email more than once

    df.to_excel('data.xlsx',index=False)

