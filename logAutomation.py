import win32com.client as win32
from tkinter import Tk
from tkinter.filedialog import askopenfilename
from colorama import Fore, Style

done = False
emails = ""

def send_mail(sendTo, cc, tabs):
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = sendTo
    mail.CC = cc
    mail.Body = ('New work has been done on ' + tabs + " tab(s).\n" 
        "\n"
        "MSS- extensions for record keeping\n"
        "\n"
        "Thank you.")

    Tk().withdraw() 
    filename = askopenfilename() 

    attachment = filename
    mail.Attachments.Add(attachment)

    pathSplit = filename.split('/')
    fileSplit = pathSplit[5].split(".21.xlsx")
    mail.Subject = 'Updated ' + fileSplit[0] + " log"

    mail.Display(True)

def names_to_email(name):
    if(name == '1'):
        return "MortgageSystemSupport@53.com"
    elif(name == '2'):
        return "Ken.Cooper@53.com"
    elif(name == '3'):
        return "Christian.Brune@53.com"
    elif(name == '4'):
        return "Joseph.Sansone@53.com"
    elif(name == '5'):
        return "JAMIE.POLAN@53.com"

def num_to_tab(num):
    if(num == '1'):
        return "Sales Refi"
    elif(num == '2'):
        return "Sales Purchase"
    elif(num == '3'):
        return "UW APPROVAL REFI"
    elif(num == '4'):
        return "UW APPROVAL PURCHASE"
    elif(num == '5'):
        return "In Closing"
    elif(num == '6'):
        return "all"

def tabsChanged(sendTo, carbons, done):
    tabs = ""
    numTabs = 0

    while not done:
        tab = input("What tabs were changed?\n"
            "1. Sales Refi\n"
            "2. Sales Purchase\n"
            "3. UW APPROVAL REFI\n"
            "4. UW APPROVAL PURCHASE\n"
            "5. In Closing\n"
            "6. All Tabs\n"
            "7. Done\n")

        if(tab == "7" or tab == "6"):
            if(tab == "6"):
                tabs += num_to_tab(tab)

            tabsSplit = ""

            if(numTabs == 2):
                tabs = tabs.replace(",", " and")
            elif(numTabs > 2):
                tabsSplit = tabs.split(",")
                lastTab = tabsSplit[len(tabsSplit) - 1]
                tabs = tabs.replace("," + lastTab, " and" + lastTab)
            send_mail(sendTo, carbons, tabs)
            done = True
        elif(tabs == ""):
            tabs += num_to_tab(tab)
            numTabs += 1
        else:
            tabs += ", " + num_to_tab(tab)
            numTabs += 1
        print("Tabs Changed: " + tabs + "\n")

def carbonCopy(sendTo, done):
    carbons = ""

    while not done:
        name = input("Who would you like to CC?\n"
            "1. MSS\n"
            "2. Ken Cooper\n"
            "3. Christian Brune\n"
            "4. Joe Sansone\n"
            "5. Jamie Polan\n"
            "6. Done\n")

        if(name == "6"):
            tabsChanged(sendTo, carbons, done)
            done = True
        elif(carbons == ""):
            carbons += names_to_email(name)
        else:
            carbons += "; " + names_to_email(name)
        print("CC: " + carbons + "\n")

def string_validation(list, email):
    if(email not in list):
        return True
    else:
        return False        

while not done:
    name = input("Who would you like to send this to?\n"
        "1. MSS\n"
        "2. Ken Cooper\n"
        "3. Christian Brune\n"
        "4. Joe Sansone\n"
        "5. Jamie Polan\n"
        "6. Done\n"
        "7. Default\n")

    if(name == "7"):
            emails = "MortgageSystemSupport@53.com; Ken.Cooper@53.com"
            cc = "Christian.Brune@53.com; Joseph.Sansone@53.com"
            tabsChanged(emails, cc, done)
    elif(name == "6"):
        carbonCopy(emails, done)
        done = True

    if(string_validation(emails, names_to_email(name)) and name != "7"): 
        if(emails == ""):
            emails += names_to_email(name)
        else:
            emails += "; " + names_to_email(name)
    else:
        print(Fore.RED + 'This email has been added already!' + Style.RESET_ALL)
    print("To: " + emails + "\n")