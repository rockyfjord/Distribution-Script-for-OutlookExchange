from xlrd import open_workbook
import win32com.client as win32
import sys
import re


max_recipients = 300  # This will limit the number of people a single message is sent to.
                      # A value of 2 for 100 recipients means 50 e-mails will be sent.

def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i+n]


def read_excel_list():
    """Open Excel file and read e-mails from first column."""
    wb = open_workbook(email_wb)
    ws = wb.sheet_by_index(0)
    raw_list = ws.col_values(0, 0)
    """Remove duplicates and divide up by max recipients."""
    emails = list(set([email.strip(' ') for email in raw_list if re.search(r"@*?\.com", email)]))
    email_blocks = chunks(emails, max_recipients)
    return [';'.join(email) for email in email_blocks]  # Convert the list into list of strings

if not len(sys.argv[1:]) == 2:
    print("This script requires 1 excel file and 1 e-mail template")
    raise Exception("InputError", *sys.argv[1:])
template = [file for file in sys.argv[1:] if file.endswith("msg")][0]
email_wb = [file for file in sys.argv[1:] if file.endswith("xls") or file.endswith("xlsx")][0]
email_strings = read_excel_list()
print(email_strings)

outlook = win32.Dispatch('outlook.application')
for recipients in email_strings:
    item = outlook.CreateItemFromTemplate(template)
    item.To = recipients
    item.Send()