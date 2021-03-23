from openpyxl import load_workbook
import xlrd as xl
import smtplib


def Send(server, rev_id, send_id, no):     # Function to send Mail
    message = "Your registered number is " + str(no)
    try:
        server.sendmail(send_id, rev_id, message)
        return True
    except:
        return False


# Load the Active sheet of the Excel File

wb = load_workbook(filename="Data.xlsx")
sheet = wb.active
loc = ("Data.xlsx")          # Calculate the number of rows
wb = xl.open_workbook(loc)
s1 = wb.sheet_by_index(0)
s1.cell_value(0, 0)
rows = s1.nrows

server = smtplib.SMTP('smtp.gmail.com', 587)  # Setting server for login
server.ehlo()
server.starttls()

sen_email = input("Email Id : ")        # Credentials for Login
sen_pass = input("Password :")

try:
    server.login(sen_email, sen_pass)   # Attempting to login
except:
    print("Wrong Credentials or Less Secure Web Access not enabled")

for i in range(2, rows+1):            # Iterating through every Row to send mail to every mail id
    mail = sheet["A"+str(i)].value
    number = sheet["B"+str(i)].value

    resp = Send(server, mail, sen_email, number)  # Response of Send

    if resp:                          # Sets Yes value to the cell next to the mail ID if Send is successful
        sheet["C"+str(i)] = "Yes"

    else:                              # Sets No value to the cell next to the mail ID if Send is unsuccessful

        sheet["C"+str(i)] = "No"

server.quit()
