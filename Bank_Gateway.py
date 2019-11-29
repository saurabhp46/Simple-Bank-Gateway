import openpyxl
import random
import smtplib
from tabulate import tabulate
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from openpyxl import load_workbook
wb = load_workbook('Account_details.xlsx')
sheet = wb['UserDetails']
ui=[]
ued={}
c1=c2=c3=c4=wc1=wc2=otp=0
ch=''
fromaddr='___________YOUR EMAIL ID FROM WHICH EMAIL WILL TO BE SEND___________'
for i in range(2,8):
    ui.append(str(sheet['B'+str(i)].value))
    ui.append(sheet['D'+str(i)].value)
    ued[sheet['E'+str(i)].value]=""
while c1<3:
    uid = input("Enter Full Name or User ID: ")
    if uid in ui:
        while c2<3:
            ban=input("Enter Bank Account Number")
            for j in range(2,8):
                if uid==str(sheet['B'+str(j)].value) or uid == sheet['D'+str(j)].value:
                    if sheet['G'+str(j)].value == "Active":
                        if ban == sheet['C'+str(j)].value:
                            print("User Validated")
                            wc1=1
                            #stage2
                            otp=random.randrange(1111,10000)
                            for k in range(2,8):
                                if uid == str(sheet['B'+str(k)].value) or uid == sheet['D'+str(k)].value:
                                    toaddr=sheet['E'+str(k)].value
                                    ued[toaddr]=str(otp)
                                
                            #Email Process
                            msg = MIMEMultipart()
                            msg['From'] = fromaddr
                            msg['To'] = toaddr
                            msg['Subject'] = "noreply"
                            body = "OTP for your bank account access is "+str(otp)
                            msg.attach(MIMEText(body,'plain'))
                            server= smtplib.SMTP('smtp.gmail.com',port=587)
                            server.starttls()
                            server.login(fromaddr,"___________YOUR EMAIL PASSWARD OF PREVIOUSLY PROVIDED EMAIL ID__________")
                            text = msg.as_string()
                            server.sendmail(fromaddr,toaddr,text)
                            server.quit()#Email Process End
                        
                            print("OTP is generated and send to registered email:",toaddr)
                            while c3<3:
                                ipotp=input("Enter recieved OTP: ")
                                if ipotp == ued[toaddr]:
                                    print("Valid OTP")
                                    wc2=1
                                    #stage3
                                    for j in range(2,8):
                                        if ban == sheet['C'+str(j)].value:
                                            bal=sheet['F'+str(j)].value
                                            cell='F'+str(j)
                                    cbal=bal
                                    print("Your Currunt Balance: ",bal)
                                    while c4<3:
                                        ch=input("Enter Your Choice ‘C’ for credit money and ‘D’ for debit money: ").upper()
                                        if ch=='C':
                                            camt=int(input("Enter amount to be credited: "))
                                            bal=bal+camt
                                            print("Successful Credit of Rs.",camt,"completed")
                                            sheet[cell] = bal
                                            wb.save('Account_details.xlsx')
                                            print("Your Updated Balance is: ",bal)
                                            print("\nMINI STATEMENT\n*************************************************************************************************************")
                                            print (tabulate([[uid,ban,cbal,camt,bal]], headers=["Account Holder ID/Name","A/C No.","Privious Balance","Credited Amount","Updated Balance"], tablefmt='orgtbl'))
                                            print("*************************************************************************************************************")
                                            break
                                        elif ch=='D':
                                            damt=int(input("Enter amount to be debited: "))
                                            if damt<=bal:
                                                bal=bal-damt
                                                print("Successful Debit of Rs.",damt,"completed")
                                                sheet[cell] = bal
                                                wb.save('Account_details.xlsx')
                                                print("Your Updated Balance is: ",bal)
                                                print("\nMINI STATEMENT\n*************************************************************************************************************")
                                                print (tabulate([[uid,ban,cbal,damt,bal]], headers=["Account Holder ID/Name","A/C No.","Privious Balance","Debited Amount","Updated Balance"], tablefmt='orgtbl'))
                                                print("*************************************************************************************************************")
                                            else:
                                                print("Insufficient Balance")
                                            break
                                        else:
                                            print("Invalid Choice... Choose ‘C’ for credit money and ‘D’ for debit money: ")
                                            c4+=1
                                    if c4==3:
                                        for j in range(2,8):
                                            if uid == str(sheet['B'+str(j)].value) or uid == sheet['D'+str(j)].value:
                                                sheet['G'+str(j)] = "Blocked"
                                                wb.save('Account_details.xlsx')
                                                print("You enterd invalid Choice three times Your Account is Blocked")
                                    break
                                else:
                                    print("Invalid OTP")
                                    c3+=1
                                    wc2=0
                                
                                if wc2==1:
                                    break
                            if c3==3:
                                for j in range(2,8):
                                    if uid == str(sheet['B'+str(j)].value) or uid == sheet['D'+str(j)].value:
                                        sheet['G'+str(j)] = "Blocked"
                                        wb.save('Account_details.xlsx')
                                        print("You enterd invalid OTP three times Your Account is Blocked")
                            break
                        else:
                            print("Invalid User")
                            c2+=1
                            wc1=0
                            break
                        wc1=1
                    else:
                        print("Sorry Your Account is Blocked")
                        wc1=1
                        break
                        
            if wc1==1:
                break
        
        if c2==3:
            for j in range(2,8):
                if uid == str(sheet['B'+str(j)].value) or uid == sheet['D'+str(j)].value:
                    sheet['G'+str(j)] = "Blocked"
                    wb.save('Account_details.xlsx')
                    print("You enterd invalid A/C no. three times Your Account is Blocked")
        
        break
    else:
        print("Invalid User")
        c1+=1
if c1==3:
    print("User Validation Program Exited\nTry again")
