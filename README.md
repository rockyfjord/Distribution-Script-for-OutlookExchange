# Distribution-Script-for-OutlookExchange
Script will divide up e-mails to avoid message size limits on an exchange server.

Prerequisites:
  1. Python 3      - https://www.python.org/downloads/release/python-351/
  2. xlrd module   -  https://pypi.python.org/pypi/xlrd
  3. pywin32       - https://sourceforge.net/projects/pywin32/files/pywin32/Build%20220/

This script requires two prepared files: 
  1. An Excel file that contains a list of e-mails in the first column.
  2. A saved Outlook Template (*.msg). Create an e-mail as you normally would, but File->SaveAs instead of sending.
  
Drag an drop both file onto the script and let 'er fly!
This script is useful for ad-hoc reports where a standard distribution list does not apply.
