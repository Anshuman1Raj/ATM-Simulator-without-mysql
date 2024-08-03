import smtplib
import random
from win32com.client import Dispatch
import json

a = random.randint(100000, 999999)

myemail = "sonal44443@gmail.com"
password = "ffjl kwne dckf eymo"

subject = "Your Verification Code"
body = f"Your verification code is {a}."
msg = f"Subject: {subject}\n\n{body}"

try:
    connection = smtplib.SMTP("smtp.gmail.com", port=587)
    connection.starttls()
    connection.login(user=myemail, password=password)

    connection.sendmail(from_addr=myemail, to_addrs="ansh49094@gmail.com", msg=msg)

    print("Email sent successfully!")

    speak = Dispatch("SAPI.SpVoice")
    speak.Voice = speak.GetVoices().Item(1)

    with open('rohit.json', 'r') as file:
        rohit_data = file.read()

    av= json.loads(rohit_data)

    def taking_input():
        speak.Speak("Enter your account number : ")
        Account_Number=int(input("Enter your account number : "))
        if Account_Number==av["accountnumber"]:
            speak.Speak("Enter your atm pin : ")
            Atm_Pin=int(input("Enter your atm pin : "))
            if Atm_Pin==av["atmpin"]:
                speak.Speak("Enter your account type : ")
                Account_Type=input("Enter your account type : ")
                if Account_Type==av["account"]:
                    speak.Speak("Enter your mobile number : ")
                    Mobile_Number=int(input("Enter your Mobile Number : "))
                    if Mobile_Number==av["phoneNumbers"]:
                        print(av["account_balance"])
                        speak.Speak("Enter the ammount you want to withdrawl")
                        Amount=int(input("Enter the ammount you want to withdrawl :"))    
                        if  av["account_balance"]>0:
                            speak.Speak("Enter your otp ")
                            otp=int(input("Enter your otp : "))
                            if otp==a:
                                if (av["account_balance"] - Amount)>0:
                                    speak.Speak(f"You have Credited by {Amount} and your updated Balance is {av['account_balance']-Amount}")
                                    print(f"You have Credited by {Amount} and your updated Balance is {av['account_balance']-Amount}")
                                    av["account_balance"] = av['account_balance']-Amount
                                    with open('rohit.json', 'w') as file:
                                        json.dump(av, file)
                                else:
                                    speak.Speak("There is insufficient amount in Your Account")
                                    print("There is insufficient amount in Your Account")
                            else:
                                speak.Speak("You have Entered Wrong otp ")
                                print("You have Entered Wrong otp")        
                        else:               
                            speak.Speak("There is insufficient amount in Your Account")
                            print("There is insufficient amount in Your Account")
                    else:
                        speak.Speak("You have Enter wrong mobile number ")
                        print("You have Enter wrong mobile number ")
                else:
                    speak.Speak("You have Enter wrong Account type ")
                    print("You have Enter wrong Account type ")
            else:
                speak.Speak("You have Enter wrong Atm Pin ")
                print("You have Enter wrong Atm Pin ")
        else:
            speak.Speak("You have Enter wrong Account number ")
            print("You have Enter wrong Account number ")
    
    taking_input()

except Exception as e:
    print(f"An error occurred: {e}")

finally:
    connection.quit()    