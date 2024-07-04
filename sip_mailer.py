from datetime import datetime,timedelta
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import formataddr
import glob
import os,sys
from pathlib import Path
import smtplib
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import pytesseract
import time
import pandas as pd

# Get the home directory
home_dir = Path.home()

# Construct the path to the Downloads folder
downloads_dir = home_dir / "Downloads"

# Set up your Tesseract OCR path if it's not in your PATH environment variable
pytesseract.pytesseract.tesseract_cmd = r'C:/Program Files/Tesseract-OCR/tesseract.exe'

# # Create a new instance of the Chrome driver
driver = webdriver.Chrome()

# Function to send email
def send_email(to_address, subject, body):
    from_address = os.environ.get('EMAIL_ID')
    app_password = os.environ.get('PASSWORD')

    # Set up the server
    server = smtplib.SMTP(host='smtp.gmail.com', port=587)
    server.starttls()
    server.login(from_address, app_password)

    # Create the email
    msg = MIMEMultipart()
    msg['From'] = formataddr(('Shivgan Associates', from_address))
    msg['To'] = to_address
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'html'))

    # Send the email
    server.send_message(msg)
    server.quit()
    
# Function to get xls file paths
def get_latest_xls_files(num_files=2):
    # Construct the path to the Downloads folder
    downloads_path = os.path.expanduser(downloads_dir)

    # Search for .xls files in the Downloads folder
    search_pattern = os.path.join(downloads_path, '*.xls')
    xls_files = glob.glob(search_pattern)

    # Sort files by modification time (latest first)
    xls_files.sort(key=os.path.getmtime, reverse=True)

    # Take the first num_files files
    latest_xls_files = xls_files[:num_files]

    return latest_xls_files

accounts = [
    {
        "name": "SID",
        "id": os.environ.get('SID_ID'),
        "password": os.environ.get('SID_PASSWORD')
    },
    # {
    #     "name": "RAJAN",
    #     "id": os.environ.get('RAJAN_ID'),
    #     "password": os.environ.get('RAJAN_PASSWORD')
    # },
    # {
    #     "name": "RESHMA",
    #     "id": os.environ.get('RESHMA_ID'),
    #     "password": os.environ.get('RESHMA_PASSWORD')
    # }
]

# Navigate to the login page
driver.get("https://www.njindiaonline.in/pdesk/login.fin?cmdAction=login")

for acc in accounts:
    # Locate the username and password fields and enter the login details
    username = driver.find_element(By.NAME, 'partnerId1')
    password = driver.find_element(By.NAME, 'password1')
    username.send_keys(str(os.environ.get('SID_ID')))
    password.send_keys(str(os.environ.get('SID_PASSWORD')))

    print(os.environ.get('SID_ID'))
    sys.stdout.flush()
    print(os.getenv('SID_ID'))
    sys.stdout.flush()
    
    # Capture the CAPTCHA image
    captcha_image = driver.find_element(By.ID, 'imgCaptcha')  # Update the XPath

    # Save the CAPTCHA image
    captcha_image.screenshot('captcha.png')

    # Use OCR to read the CAPTCHA
    captcha_text = pytesseract.image_to_string(Image.open('captcha.png')).strip()
    print(captcha_text)
    sys.stdout.flush()
    captcha_text = captcha_text.replace(" ", "")

    # Enter the CAPTCHA text
    captcha_field = driver.find_element(By.NAME, 'capcode')
    captcha_field.send_keys(captcha_text)

    # Submit the form
    driver.find_element(By.NAME, 'action').click()
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.NAME, 'action'))).click()
    time.sleep(5)
    WebDriverWait(driver, 5).until(EC.presence_of_element_located((By.XPATH, '//a[text()="Stock Exchange"]'))).click()
    # driver.find_element(By.XPATH, '//a[text()="Stock Exchange"]').click()
    driver.find_element(By.XPATH, '//b[text()="SIP Status Report"]').click()
    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.NAME, 'apply'))).click()
    # driver.find_element(By.NAME, 'apply').click()
    time.sleep(3)
    driver.find_element(By.ID, 'export_xls').click()

    driver.execute_script("window.history.go(-1)")
    time.sleep(2)
    driver.find_element(By.XPATH, "//a[@onclick='javascript:getAccountDetail();']").click()

    # Get the current window handle
    original_window = driver.current_window_handle

    # Wait for the new window or tab
    WebDriverWait(driver, 3).until(EC.number_of_windows_to_be(2))

    # Loop through until we find a new window handle
    for window_handle in driver.window_handles:
        if window_handle != original_window:
            driver.switch_to.window(window_handle)
            break

    WebDriverWait(driver, 3).until(EC.presence_of_element_located((By.ID, 'export_xls'))).click()
    time.sleep(5)
    
    # Load the sheets from the provided files
    latest_xls_files = get_latest_xls_files(num_files=2)
    file1_path = latest_xls_files[0].replace("\\","/")
    file2_path = latest_xls_files[1].replace("\\","/")

    print(file1_path,file2_path)

    # Load the sheets into pandas DataFrames
    sheet1 = pd.read_excel(file1_path, engine='xlrd')
    sheet2 = pd.read_excel(file2_path, engine='xlrd')
    sheet1.columns = sheet1.iloc[0]
    sheet1 = sheet1.iloc[1:].reset_index(drop=True)
    sheet1.drop(sheet1.tail(2).index,inplace=True)
    sheet2.columns = sheet2.iloc[0]
    sheet2 = sheet2.iloc[1:].reset_index(drop=True)
    sheet2.drop(sheet2.tail(1).index,inplace=True)

    # Rename the columns to have the same name
    if "ActiveTrading" in file1_path:
        sheet1.rename(columns={'Client Code (UCC)': 'UCC'}, inplace=True)
    else:
        sheet2.rename(columns={'Client Code (UCC)': 'UCC'}, inplace=True)


    # Merge the sheets based on the UCC column
    merged_sheet = pd.merge(sheet1, sheet2, on='UCC', how='inner')
    merged_sheet['SIP Submission Date'] = pd.to_datetime(merged_sheet['SIP Submission Date'], dayfirst=True)

    end_date = datetime.now()
    start_date = end_date - timedelta(days=7)


    merged_sheet = merged_sheet[(merged_sheet['SIP Submission Date'] >= start_date) & (merged_sheet['SIP Submission Date'] <= end_date)]

    # Save the merged sheet to a new file
    output_path = downloads_dir /'merged_sheet.xlsx'
    merged_sheet.to_excel(output_path, index=False)
    # merged_sheet = pd.read_excel(output_path)

    # Count the occurrences of each Investor
    investor_counts = merged_sheet['Investor'].value_counts()

    # Investors appearing only once
    single_time_investors = investor_counts[investor_counts == 1].index.tolist()

    # Investors appearing multiple times
    multiple_times_investors = investor_counts[investor_counts > 1].index.tolist()

    # DataFrames for single time and multiple times investors
    single_sip = merged_sheet[merged_sheet['Investor'].isin(single_time_investors)].reset_index()
    multiple_sip = merged_sheet[merged_sheet['Investor'].isin(multiple_times_investors)].reset_index()

    # Create sub-dataframes for multiple-time investors
    multiple_investor_dfs = {investor: multiple_sip[multiple_sip['Investor'] == investor] for investor in multiple_times_investors}


    # Iterate through the single SIP DataFrame and send emails
    for index, row in single_sip.iterrows():
        to_address = row['E-Mail ID']
        name = row['Investor']
        scheme = row['Scheme']
        amount = row['Installment Amt']
        sip_date = row['SIP Start Date']
        
        # Define the email subject and body
        subject = "Mutual Fund SIP Transaction Confirmation"
        body = f"""
        <div dir="ltr"><span style="color:rgb(34,34,34)">Dear {name}</span><span style="color:rgb(34,34,34)">,</span><br style="color:rgb(34,34,34)"><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">I trust this email finds you well. We appreciate your continued trust and partnership with Shivgan Associates. This communication is to confirm the recent mutual fund transaction that aligns with your long term investment goals.</span><br style="color:rgb(34,34,34)">
        <div><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">SIP Scheme: {scheme}</span><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">Start Date:{sip_date}</span><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">Amount: ₹{amount}</span>
        <div style="color:rgb(34,34,34)"><div><br>We want to emphasize that your recent transaction reflects your commitment to a long-term investment strategy. The chosen mutual fund aligns with your goal of investing, demonstrating a thoughtful and strategic approach to wealth building.<br><br><b>Key Points:</b><br><br><b>Long-Term Perspective</b>: Investing for long term allows your portfolio to benefit from the power of compounding and ride out market fluctuations.<br><br><b>Diversification</b>: The selected mutual fund is part of a diversified portfolio, which helps spread risk and enhance potential returns over the long run.<br><br><b>Professional Management</b>: The fund is managed by experienced professionals who continuously analyze market trends and make strategic decisions to optimize returns.<br><br><b>Regular Monitoring</b>: Our team will continue to monitor the performance of your investment and provide periodic NFO updates to ensure that you don't miss any good investing opportunities.<br><br>If you have any questions or require further clarification regarding this transaction or your overall investment strategy, please do not hesitate to reach out to us. We are here to assist you in achieving your financial goals and ensuring a smooth and transparent investment experience.<br><br>Thank you for choosing Shivgan Associates as your trusted financial partner. We look forward to continuing our journey together towards financial success.</div><div><br>Regards,<br clear="all"><div><div dir="ltr"><div dir="ltr"><table cellpadding="0" cellspacing="0" style="font-size:16px;border-collapse:collapse;font-family:Arial;line-height:1.15;color:rgb(0,0,0)"><tbody><tr><td style="vertical-align:top;padding:0.01px 14px 0.01px 1px;width:65px;text-align:center"><img src="https://ci3.googleusercontent.com/meips/ADKq_NbmscugMmVQ9cbKiW0YxLtCdkFa2elmHB6ErWk82zNjD_NB90zb_8uP40JxmEcMbW8AysrpZ7SCcZuV5wr9Ycs3hN17X6hHrMWYggdaprEeJK25tPrQ-znwXmJs7Tk3RcmYM4vikLo=s0-d-e1-ft#https://d36urhup7zbd7q.cloudfront.net/a/badf5ea0-e04c-4a02-8a25-496ee644073a.jpeg" height="65" width="65" alt="photo" style="width:65px;vertical-align:middle;border-radius:0px;height:65px"></td><td valign="top" style="padding:0.01px 0.01px 0.01px 14px;vertical-align:top;border-left:1px solid rgb(189,189,189)"><table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tbody><tr><td style="padding:0.01px"><p style="margin:0.1px;line-height:19.2px"><span style="font-weight:bold;color:rgb(100,100,100);letter-spacing:0px">Rajan G. Shivgan</span><br><span style="font-size:13px;font-weight:bold;color:rgb(100,100,100)">Family's Financial Expert,&nbsp;</span><span style="font-size:13px;font-weight:bold;color:rgb(100,100,100)">Shivgan&nbsp;</span><font color="#646464"><span style="font-size:13px"><b>Associates</b></span></font></p></td></tr><tr><td><table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tbody><tr><td nowrap="" width="71" style="padding-top:14px;width:71px"><p style="margin:1px;line-height:10.89px;font-size:11px">9619801859 / 8928232224 |&nbsp;<a href="mailto:shivganassociates14@gmail.com" style="color:rgb(17,85,204)" target="_blank">shivganassociates14@gmail.com</a>&nbsp;</p><div><br></div></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div></div></div><br></div></div></div></div>
        """
        
        # Send the email
        # send_email(to_address, subject, body)
        print("Email sent to ",name)

    # Iterate through the multiple SIP DataFrame and send emails
    for investor, investments in multiple_investor_dfs.items():
        # Define the email subject and body
        subject = "Mutual Fund SIP Transaction Confirmation"
        body = '<div dir="ltr"><span style="color:rgb(34,34,34)">Dear {}</span><span style="color:rgb(34,34,34)">,</span><br style="color:rgb(34,34,34)"><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">I trust this email finds you well. We appreciate your continued trust and partnership with Shivgan Associates. This communication is to confirm the recent mutual fund transaction that aligns with your long term investment goals.</span><br style="color:rgb(34,34,34)">'.format(investor)
        for index,inv in investments.iterrows():
            body+= '<div><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">SIP 1: {}</span><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">Start Date: {}</span><br style="color:rgb(34,34,34)"><span style="color:rgb(34,34,34)">Amount: ₹{}</span>'.format(inv['Scheme'],inv['SIP Start Date'],inv['Installment Amt'])

        body+='''<div style="color:rgb(34,34,34)"><div><br>We want to emphasize that your recent transaction reflects your commitment to a long-term investment strategy. The chosen mutual fund aligns with your goal of investing, demonstrating a thoughtful and strategic approach to wealth building.<br><br><b>Key Points:</b><br><br><b>Long-Term Perspective</b>: Investing for long term allows your portfolio to benefit from the power of compounding and ride out market fluctuations.<br><br><b>Diversification</b>: The selected mutual fund is part of a diversified portfolio, which helps spread risk and enhance potential returns over the long run.<br><br><b>Professional Management</b>: The fund is managed by experienced professionals who continuously analyze market trends and make strategic decisions to optimize returns.<br><br><b>Regular Monitoring</b>: Our team will continue to monitor the performance of your investment and provide periodic NFO updates to ensure that you don't miss any good investing opportunities.<br><br>If you have any questions or require further clarification regarding this transaction or your overall investment strategy, please do not hesitate to reach out to us. We are here to assist you in achieving your financial goals and ensuring a smooth and transparent investment experience.<br><br>Thank you for choosing Shivgan Associates as your trusted financial partner. We look forward to continuing our journey together towards financial success.</div><div><br>Regards,<br clear="all"><div><div dir="ltr"><div dir="ltr"><table cellpadding="0" cellspacing="0" style="font-size:16px;border-collapse:collapse;font-family:Arial;line-height:1.15;color:rgb(0,0,0)"><tbody><tr><td style="vertical-align:top;padding:0.01px 14px 0.01px 1px;width:65px;text-align:center"><img src="https://ci3.googleusercontent.com/meips/ADKq_NbmscugMmVQ9cbKiW0YxLtCdkFa2elmHB6ErWk82zNjD_NB90zb_8uP40JxmEcMbW8AysrpZ7SCcZuV5wr9Ycs3hN17X6hHrMWYggdaprEeJK25tPrQ-znwXmJs7Tk3RcmYM4vikLo=s0-d-e1-ft#https://d36urhup7zbd7q.cloudfront.net/a/badf5ea0-e04c-4a02-8a25-496ee644073a.jpeg" height="65" width="65" alt="photo" style="width:65px;vertical-align:middle;border-radius:0px;height:65px"></td><td valign="top" style="padding:0.01px 0.01px 0.01px 14px;vertical-align:top;border-left:1px solid rgb(189,189,189)"><table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tbody><tr><td style="padding:0.01px"><p style="margin:0.1px;line-height:19.2px"><span style="font-weight:bold;color:rgb(100,100,100);letter-spacing:0px">Rajan G. Shivgan</span><br><span style="font-size:13px;font-weight:bold;color:rgb(100,100,100)">Family's Financial Expert,&nbsp;</span><span style="font-size:13px;font-weight:bold;color:rgb(100,100,100)">Shivgan&nbsp;</span><font color="#646464"><span style="font-size:13px"><b>Associates</b></span></font></p></td></tr><tr><td><table cellpadding="0" cellspacing="0" style="border-collapse:collapse"><tbody><tr><td nowrap="" width="71" style="padding-top:14px;width:71px"><p style="margin:1px;line-height:10.89px;font-size:11px">9619801859 / 8928232224 |&nbsp;<a href="mailto:shivganassociates14@gmail.com" style="color:rgb(17,85,204)" target="_blank">shivganassociates14@gmail.com</a>&nbsp;</p><div><br></div></td></tr></tbody></table></td></tr></tbody></table></td></tr></tbody></table></div></div></div><br></div></div></div></div>'''

        
        # Send the email
        # send_email(to_address, subject, body)
        print("Email sent to ",investor)
    
    driver.close()
    driver.switch_to.window(driver.window_handles[0])
    driver.get(os.environ.get('PARTNER_DESK'))


driver.quit()

