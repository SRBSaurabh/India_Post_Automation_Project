import os
import time
import shutil
import sqlite3
from urllib.request import urlopen
from datetime import date, datetime
from timeit import default_timer as timer

try:
    import cv2
    import pyttsx3
    import pytesseract
    import openpyxl
    import pandas as pd
    import colorama
    from PIL import Image
    from colorama import Fore, Back, Style
    from selenium import webdriver
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.options import Options
    from tqdm import tqdm
    from playsound import playsound
    from bs4 import BeautifulSoup
except Exception as ex:
    print('\n       ', ex, '\n')
    os.system('pip install -r requirements.txt')
    print("\n\n"
          "=======================================================================================================\n\n"
          "                          @ Necessary 'FIRST Time' - Additional Requirements Installed ..! !                 \n\n"
          "                     (^_^)    Please Restart the Program Once Again ..! !                \n\n"
          "=======================================================================================================\n\n")

    input('Press ENTER to EXIT . . .')
    quit()

print('@@@@@@@@@____________Fully Developed By "Saurabh Chaudhari" ______ :)')
# Chrome Download Path
sss = timer()
today = date.today().strftime("%d-%m-%Y")
opn = (date.today() - pd.DateOffset(months=12 * 5)).strftime("%Y-%m-%d")
MM = datetime.now().strftime("%m")
YYYY = datetime.now().strftime("%Y")
colorama.init(autoreset=True)
Current_Year = YYYY
Current_Month = datetime.now().strftime("%b")
Current_month_num = MM
a, b = [], []

# Previous RD-Interest Rates with FORMAT Of
# Opening | Closing | Interest | 60-Months | 120-Months
Int_Rates = """01-12-2011   31-03-2012  8.00    738.62  1836.17
01-04-2012  31-03-2013  8.40    746.53  1877.73
01-04-2013  31-03-2014  8.30    744.53  1867.24
01-04-2014  17-11-2014  8.40    746.53  1877.73
18-11-2014  20-01-2015  8.40    746.53  1877.73
21-01-2015  31-03-2015  8.40    746.53  1877.73
01-04-2015  31-03-2016  8.40    746.53  1877.73
01-04-2016  30-09-2016  7.40    726.97  1775.88
01-10-2016  31-03-2017  7.30    725.05  1766.07
01-04-2017  30-06-2017  7.20    723.14  1756.32
01-07-2017  31-12-2017  7.10    721.23  1746.64
01-01-2018  30-09-2018  6.90    717.43  1727.46
01-10-2018  31-12-2018  7.30    725.05  1766.07
01-01-2019  31-03-2019  7.30    725.05  1766.07
01-04-2019  30-06-2019  7.30    725.05  1766.07
01-07-2019  30-09-2019  7.20    723.14  1756.32
01-10-2019  31-12-2019  7.20    723.14  1756.32
01-01-2020  31-03-2020  7.20    723.14  1756.32
01-04-2020  {}  5.80    696.967   1626.48""".format(today)

cal = {"Jan": "01",
       "Feb": "02",
       "Mar": "03",
       "Apr": "04",
       "May": "05",
       "Jun": "06",
       "Jul": "07",
       "Aug": "08",
       "Sep": "09",
       "Oct": "10",
       "Nov": "11",
       "Dec": "12"}

dtStamp = datetime.now().strftime("%d-%m-%Y  %H:%M:%S")


def ColoredPrint(msg, color):
    colorama.init(autoreset=True)

    return (f"{color}{Style.BRIGHT}\n\n\n"
            "---------*********************************************************************************-----------\n\n"
            f"{Fore.LIGHTMAGENTA_EX}{Style.BRIGHT}                       {msg}                  \n\n"
            f"{color}{Style.BRIGHT}"
            f"---------********************************************************************************-----------\n\n")


def recCaptcha(img_path):
    result = None
    # Read image with opencv
    img = cv2.imread(img_path)
    # Convert to gray
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    cv2.imwrite("clear.png", img)
    result = pytesseract.image_to_string(Image.open("clear.png"))
    os.remove("clear.png")
    return result[:-2]


# Getting Id & Password...
wb = openpyxl.load_workbook(r"Portal.xlsx", read_only=True, data_only=True)
sheet = wb["Summary"]
password = sheet['J1'].value
identity = sheet['Z1'].value
wb.close()

# Browser Setup & WebPage Loading...
option = Options()
preferences = {"download.default_directory": os.getcwd(), "safebrowsing.enabled": "false"}  # Chrome download Path
option.add_experimental_option("prefs", preferences)
option.add_argument("--log-level=3")
option.add_argument("--headless")  # =========to make headless
option.add_argument("--window-size=1280,720")

url = 'https://dopagent.indiapost.gov.in/corp/AuthenticationController?FORMSGROUP_ID__=AuthenticationFG' \
          '&__START_TRAN_FLAG__=Y&__FG_BUTTONS__=LOAD&ACTION.LOAD=Y&AuthenticationFG.LOGIN_FLAG=3&BANK_ID=DOP' \
          '&AGENT_FLAG=Y '
try:
    browser = webdriver.Chrome(ChromeDriverManager().install(), options=option)
    browser.get(url)
    print(ColoredPrint('Opening India Post in Background.....', Fore.LIGHTBLUE_EX))
except Exception as eo:
    print(ColoredPrint('Internet Problem\n\n --> >' + str(eo), Fore.LIGHTRED_EX))
    time.sleep(5)
    quit()

try:
    browser.find_element_by_id('AuthenticationFG.USER_PRINCIPAL').send_keys(identity)
    browser.find_element_by_id('AuthenticationFG.ACCESS_CODE').send_keys(password)

    try:
        browser.find_element_by_id('IMAGECAPTCHA').screenshot('captcha.png')
        cap = recCaptcha('captcha.png')

        while len(cap) > 6 or sum([i not in 'ABCDEFGHKLMNPRTUVWXYabdfhkmnpqrtuwxy2345678' for i in cap]) > 0:
            os.remove("captcha.png")
            browser.find_element_by_id('TEXTIMAGE').click()
            browser.find_element_by_id('IMAGECAPTCHA').screenshot('captcha.png')
            cap = recCaptcha('captcha.png')

        time.sleep(1.2)
        browser.find_element_by_id('IMAGECAPTCHA').screenshot('captchaaa.png')
        cap = recCaptcha('captchaaa.png')
        print(cap)
        browser.find_element_by_id('AuthenticationFG.VERIFICATION_CODE').clear()
        time.sleep(0.5)
        browser.find_element_by_id('AuthenticationFG.VERIFICATION_CODE').send_keys(cap)
        time.sleep(2)
        browser.find_element_by_id('VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED').click()
        os.remove("captchaaa.png")
        os.remove("captcha.png")
        browser.implicitly_wait(8)
        # Assign aa, bb just for the sake of password timeline
        aa = browser.find_element_by_id('signOnpwd').get_attribute('innerText')
        bb = browser.find_element_by_id('HREF_DashboardFG.LOGIN_EXPIRY_DAYS').get_attribute('innerText')
        print(
            f"{'Login Successful..!'}\n\n   {Fore.LIGHTRED_EX}{Style.BRIGHT}{aa}------------------ > > > {Fore.LIGHTGREEN_EX}{Style.BRIGHT}{bb}\n\n")
    except:
        try:
            browser.close()
            try:
                browser = webdriver.Chrome(options=option)
                browser.get(url)
            except Exception as e:
                print(ColoredPrint('Internet Problem\n\n --> >' + str(e), Fore.LIGHTRED_EX))
                time.sleep(5)
                quit()

            print('Varify "CAPTCHA" Manually...!!!')
            browser.find_element_by_id('AuthenticationFG.USER_PRINCIPAL').send_keys(identity)
            browser.find_element_by_id('AuthenticationFG.ACCESS_CODE').send_keys(password)
            browser.find_element_by_id('IMAGECAPTCHA').screenshot('captcha.png')
            cap = recCaptcha('captcha.png')

            print(ColoredPrint(cap, Fore.LIGHTRED_EX))
            Image.open('captcha.png').show()
            captcha = input('Enter the "CAPTCH" here --> ')
            if len(captcha) < 1:
                pass
            else:
                cap = captcha
            print(ColoredPrint(cap, Fore.LIGHTGREEN_EX))
            time.sleep(0.25)
            browser.find_element_by_id('AuthenticationFG.VERIFICATION_CODE').send_keys(cap)
            time.sleep(2)
            browser.find_element_by_id('VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED').click()
            os.remove("captcha.png")
            browser.implicitly_wait(12)
            # Assign aa, bb just for the sake of password timeline
            aa = browser.find_element_by_id('signOnpwd').get_attribute('innerText')
            bb = browser.find_element_by_id('HREF_DashboardFG.LOGIN_EXPIRY_DAYS').get_attribute('innerText')
            print(
                f"{'Login Successful..!'}\n\n   {Fore.LIGHTRED_EX}{Style.BRIGHT}{aa}------------------ > > > {Fore.LIGHTGREEN_EX}{Style.BRIGHT}{bb}\n\n")
        except Exception as e:
            print(ColoredPrint('"Failed" Due to InCorrect CAPTCHA ... !!!', Fore.LIGHTRED_EX))
            quit(e)

    browser.find_element_by_id('Accounts').click()
    browser.find_element_by_id('Agent Enquire & Update Screen').click()
    print(ColoredPrint("We Are In The 'INDIA POST'    (^_^)", Fore.GREEN))
    time.sleep(.7)
    browser.find_element_by_id('NEXT_ACCOUNTS').click()
    browser.find_element_by_id('printpreview').click()
    new_tab = browser.window_handles[1]
    browser.switch_to.window(new_tab)
    link = browser.current_url
except Exception as e:
    # print(ColoredPrint(f'Internet / Website is BUSY.....!!\n{e}', Fore.LIGHTRED_EX))
    quit(ColoredPrint(f'Internet / Website is BUSY.....!!\n{e}', Fore.LIGHTRED_EX))

st = timer()

# # to save password remainder
conn = sqlite3.connect('Portal_Data.sqlite')
cur = conn.cursor()
cur.execute('DROP TABLE IF EXISTS Password')
cur.execute('CREATE TABLE Password (Msg TEXT, Days TEXT, Date TEXT)')
cur.execute('INSERT INTO Password(Msg, Days, Date) VALUES (?, ?, ?)', (aa, bb, dtStamp))
conn.commit()
cur.close()

# # to Update Portal
raw = urlopen(link).read()
soup = BeautifulSoup(raw, 'html.parser')
extract = (soup.get_text().split('\n'))

try:
    j1 = extract.index('Deposit Accounts List') + 6
    j2 = extract.index('Additional Information') - 12

    for i in tqdm(extract[j1:j2:8], unit=' Items', desc=f'{Fore.LIGHTRED_EX}{Style.BRIGHT}Dataset Downloading',
                  bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.LIGHTGREEN_EX, Fore.RESET)):
        a.append(i)
    for i in extract[j1 + 1:j2:8]:
        if len(i) > 0:
            lst_nm = [i.split()[-1]]
            lst_nm.extend(i.split()[0:-1])
            b.append(" ".join(lst_nm))
        else:
            continue
    c = [i.split('.')[0].replace(',', '') for i in extract[j1 + 2:j2:8]]
    d = [_ for _ in extract[j1 + 3:j2:8]]
    e = [i for i in extract[j1 + 4:j2:8]]
    f = [i[:2] for i in extract[j1 + 4:j2:8]]
    g = [i[3:6] for i in extract[j1 + 4:j2:8]]
    h = [i[7:11] for i in extract[j1 + 4:j2:8]]

except ValueError as e0:
    # print(ColoredPrint(f'Your Session is Expired \n\n{e0}', Fore.LIGHTRED_EX))
    quit(ColoredPrint(f'Your Session is Expired \n\n{e0}', Fore.LIGHTRED_EX))


def show_updates(old_or_new_list, color):
    """To Show Which Accounts Are Removed & Added"""
    con = sqlite3.connect('Portal_Data.sqlite')
    cursor = con.cursor()
    for _ in old_or_new_list:
        show = list(cursor.execute('SELECT Account_No, Account_Name, Denomination, Opening_Date, Closing_Date, '
                                   'Month_Paid_Upto FROM Portal WHERE Account_No = ?', (_,)))
        print(f"{color}{Style.BRIGHT}{list(show[0])}")
    con.close()


# Catching New & Old Accounts From Website & DataSet respectively
updated_accounts = [i for i in a if i]
conn = sqlite3.connect('Portal_Data.sqlite')
cur = conn.cursor()
old_accounts = list(cur.execute('SELECT Account_No FROM Portal'))
old_accounts = [list(i)[0] for i in old_accounts if i]
conn.close()

# Showing Removed Accounts
removed = set(old_accounts).difference(updated_accounts)
# print('Accounts Removed From Old Portal List ===>', list(removed), len(set(removed)))
print('\n\n', ColoredPrint(f"No. Of 'OLD' Accounts 'Removed' ---> {Fore.LIGHTGREEN_EX}{Style.BRIGHT}{len(removed)}",
                           Fore.LIGHTRED_EX))
show_updates(removed, Fore.LIGHTRED_EX)

conn = sqlite3.connect('Portal_Data.sqlite')
cur = conn.cursor()
cur.execute('DROP TABLE IF EXISTS Portal')
cur.execute('CREATE TABLE Portal (Account_No TEXT, Account_Name TEXT, Denomination INTEGER, Opening_Date TEXT, '
            'Closing_Date TEXT, Month_Paid_Upto INTEGER, '
            'Next_Installment_Due_Date TEXT, Regular_Installment TEXT, Pending_Installment TEXT, '
            'Advance_Installment TEXT, Total_Return INTEGER, Int REAL, '
            'Day INTEGER, Month TEXT, Month_Num INTEGER, Year INTEGER)')
for k0 in list(zip(a, b, c, d, e, f, g, h)):
    cur.execute('INSERT INTO Portal (Account_No, Account_Name, Denomination, Month_Paid_Upto, '
                'Next_Installment_Due_Date, Day, Month, Year) VALUES (?,?,?,?,?,?,?,?)', k0)
cur.execute("UPDATE Portal SET Regular_Installment = '2nd Half' WHERE Day > 15 AND (Month, Year) = (?,?)",
            (Current_Month, Current_Year))
cur.execute("UPDATE Portal SET Regular_Installment = '1st Half' WHERE Day < 15 AND (Month, Year) = (?,?)",
            (Current_Month, Current_Year))
cur.execute("UPDATE Portal SET Regular_Installment = '1st Half' WHERE Day = 15 AND (Month, Year) = (?,?)",
            (Current_Month, Current_Year))
for i in cal.keys():
    val = cal.get(i)
    cur.execute("UPDATE Portal SET Month_Num = ? WHERE Month = ?", (val, i))

cur.execute("UPDATE Portal SET Advance_Installment = (Month_Num - ?) WHERE (Year = ? AND Month_Num > ?)",
            (Current_month_num, Current_Year, Current_month_num))
cur.execute("UPDATE Portal SET Advance_Installment = (Month_Num + 12 - ?) WHERE  Year > ?",
            (Current_month_num, Current_Year))
cur.execute("UPDATE Portal SET Pending_Installment = (? - Month_Num) WHERE (Month_Paid_Upto <> 60 "
            "AND Year = ? AND Month_Num < ?)", (Current_month_num, Current_Year, Current_month_num))
cur.execute("UPDATE Portal SET Pending_Installment = (12 - Month_Num + ?) WHERE  Year < ?",
            (Current_month_num, Current_Year))

a1 = cur.execute('SELECT Year, Month_Num, Day, Month_Paid_Upto, Account_No FROM Portal')
odate1, cdate1, ac_no = [], [], []
for date in a1:
    try:
        date02 = f"{date[0]}-{int(date[1]):02}-{date[2]:02}"
    except TypeError:
        continue

    current_date = datetime.fromisoformat(date02)
    upto = int(date[3])
    if upto <= 60:
        cdate = current_date + pd.DateOffset(months=(60 - upto))
        odate = cdate - pd.DateOffset(months=12 * 5)
        cdate = cdate.strftime('%Y-%m-%d')
        odate = odate.strftime('%Y-%m-%d')
        odate1.append(odate)
        cdate1.append(cdate)
        ac_no.append(date[4])
    else:
        cdate = current_date + pd.DateOffset(months=(120 - upto))
        odate = cdate - pd.DateOffset(months=12 * 10)
        cdate = cdate.strftime('%Y-%m-%d')
        odate = odate.strftime('%Y-%m-%d')
        odate1.append(odate)
        cdate1.append(cdate)
        ac_no.append(date[4])

for zz in list(zip(odate1, cdate1, ac_no)):
    cur.execute("UPDATE Portal SET Opening_Date = ?, Closing_Date = ? WHERE Account_No = ?", zz)

# def maturity_calc(monthly_deposit, interest_percent, months_paid):
#     """  Getting Maturity Amount... """
#     int_factor = interest_percent / 1200
#     deposited, total_normal_int = 0.00, 0.00
#     for month in range(1, months_paid + 1):
#         if month % 3 == 0:
#             deposited += monthly_deposit
#             total_normal_int += deposited * int_factor
#             deposited += total_normal_int
#             total_normal_int = 0.00
#         else:
#             deposited += monthly_deposit
#             total_normal_int += deposited * int_factor
#     return round(deposited, 4) - 0.0200

cur.execute("UPDATE Portal SET Total_Return = NULL, Int = NULL")
for itm in Int_Rates.split('\n')[::-1]:
    From = datetime.strptime(itm.split()[0], "%d-%m-%Y").strftime("%Y-%m-%d")
    To = datetime.strptime(itm.split()[1], "%d-%m-%Y").strftime("%Y-%m-%d")
    Interest = itm.split()[2]
    Int_Factor = itm.split()[3]
    Ext_Factor = itm.split()[4]
    cur.execute("UPDATE Portal SET Total_Return = ROUND(Denomination/10 * ?), Int = ? WHERE Opening_Date is NULL AND "
                "date(?) >= date(?) AND date(?) <= date(?)", (Int_Factor, Interest, opn, From, opn, To))
    cur.execute("UPDATE Portal SET Total_Return = ROUND(Denomination/10 * ?), Int = ? WHERE Month_Paid_Upto <= 60 AND "
                "date(Opening_Date) >= date(?) AND date(Opening_Date) <= date(?)", (Int_Factor, Interest, From, To))
    cur.execute("UPDATE Portal SET Total_Return = ROUND(Denomination/10 * ?), Int = ? WHERE Month_Paid_Upto > 60 AND "
                "date(Opening_Date) >= date(?) AND date(Opening_Date) <= date(?)", (Ext_Factor, Interest, From, To))
conn.commit()
cur.close()

# Showing Newly Added Accounts
new = set(updated_accounts).difference(old_accounts)
# print('Newly Added Accounts ===> ', list(new), len(set(new)))
print(
    ColoredPrint(f"No. Of 'NEW' Accounts 'Added' ---> {Fore.LIGHTGREEN_EX}{Style.BRIGHT}{len(new)}", Fore.LIGHTGREEN_EX),
    '\n\n')
show_updates(new, Fore.LIGHTGREEN_EX)

print(ColoredPrint("100% 'Portal Updated' Successfully..!!  :)", Fore.LIGHTGREEN_EX))
ed = timer()
ddd = timer()

# # Too Delete Installments PAID History of LAST Month
conn = sqlite3.connect('Portal_Data.sqlite')
cur = conn.cursor()
try:
    eraser = cur.execute(
        "SELECT substr(Date,4,2) AS MM, substr(Date,7,4) AS YYYY FROM History WHERE YYYY= ? AND MM = ? LIMIT 1",
        (YYYY, MM))
    gg = [g for g in eraser]
    if len(gg) < 1:
        print(ColoredPrint(f'Current Month: {MM} & Year: {YYYY} NOT Found in Current DataBase....!',
                           Fore.LIGHTMAGENTA_EX))
        print(3 * ColoredPrint(f'{Fore.LIGHTRED_EX}Your Installments PAID History Of Previous Month is "Deleted" '
                               f'Successfully..! ! !', Fore.LIGHTRED_EX))
        cur.execute('DROP TABLE IF EXISTS History')
        cur.execute('CREATE TABLE History(LOT TEXT, Accounts TEXT, Date TEXT)')
        # # ------------------- >>>>>>>>>> To Remove RDs Directory/
        shutil.rmtree('RDs')
        time.sleep(0.30)
        os.system('mkdir RDs')
        print(ColoredPrint(' Successfully Completed Cleaning Old "RDs" folder....! !', Fore.LIGHTMAGENTA_EX))
    else:
        print(f'\t\t{Fore.LIGHTBLUE_EX}{Style.BRIGHT} Current Month: {MM} & Year: {YYYY} are Already Present in '
              f'Current DataBase....!\n\n\n')
        pass
except Exception as e:
    print(2 * ColoredPrint(f'\n\n{Fore.LIGHTRED_EX}{Style.BRIGHT} Your Installments PAID History Deletion "FAILED" ! ! '
                           f'!\n\n{e}', Fore.LIGHTRED_EX))
conn.commit()
cur.close()

print(f"{Fore.LIGHTMAGENTA_EX} DataBase Updated in Just:     '{ed - st}' Seconds")
print(ColoredPrint(f"Fully  Completed in :'{ddd - sss}' Seconds", Fore.LIGHTGREEN_EX))

# Completion Sound
playsound('result.wav')

input('Press ENTER to EXIT or Close the Window. . .')
