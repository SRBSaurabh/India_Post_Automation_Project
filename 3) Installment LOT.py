import openpyxl
import shutil
import os
import sqlite3
import pyttsx3
import threading
import cv2
import pytesseract
import time
import tkinter as tk
import colorama
from PIL import Image
from colorama import Fore, Back, Style
from datetime import date, datetime
from timeit import default_timer as timer
from tkinter import messagebox, ttk
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.chrome.options import Options
from tqdm import tqdm
from playsound import playsound

lot_lis, inst_no = None, 0


def frontEnd():
    def onClick():
        global lot_lis, inst_no
        lot_lis = entry1.get()
        inst_no = v.get()

        label3 = tk.Label(root, text=" Let's Go...!!!", font=('helvetica', 15), bg='yellow', fg='#f00')
        canvas1.create_window(650, 250, window=label3)

        root.quit()

    def onPress(event):
        print("You hit ENTER.")
        onClick()

    root = tk.Tk()
    root.title('Created by SRBSaurabh  😎')
    root.iconbitmap('icon.ico')

    canvas1 = tk.Canvas(root, width=850, height=300, relief='raised', bg='#60408E')
    canvas1.pack()

    label1 = tk.Label(root, text='  *  India Post Virtual Agent  *  ', bg='#FF344C')
    label1.config(font=('helvetica', 14))
    canvas1.create_window(400, 25, window=label1)

    label2 = tk.Label(root, text='Enter the "Space" Separated LOT Numbers below: ', bg='#E6FF33')
    label2.config(font=('helvetica', 25))
    canvas1.create_window(400, 100, window=label2)

    # entry box
    entry1 = tk.Entry(root, font=('calibre', 35, 'normal'))
    canvas1.create_window(400, 170, window=entry1)

    button1 = tk.Button(text='START', command=onClick, bg='#45FF00', fg='red',
                        font=('helvetica', 20, 'bold'))
    canvas1.create_window(420, 250, window=button1)

    # Tkinter string variable able to store any string value
    v = tk.StringVar(root, "1")

    # # # Style class to add style to Radiobutton it can be used to style any ttk widget
    style = ttk.Style(root)
    style.configure("TRadiobutton", background="#00FFFF", foreground="red", font=("arial", 30, "bold"))

    ttk.Radiobutton(root, text=" 1st  Installment", variable=v, value=1).pack(side=tk.TOP, ipady=5)
    ttk.Radiobutton(root, text=" 2nd Installment", variable=v, value=2).pack(side=tk.TOP, ipady=5)

    root.bind('<Return>', onPress)
    tk.mainloop()
    return inst_no, lot_lis


def ColoredPrint(msg, color):
    colorama.init(autoreset=True)

    return (f"{color}{Style.BRIGHT}\n\n"
            "---------*********************************************************************************-----------\n\n"
            f"{Fore.LIGHTMAGENTA_EX}{Style.BRIGHT}                       {msg}                  \n\n"
            f"{color}{Style.BRIGHT}"
            f"---------********************************************************************************-----------\n\n")


def recCaptcha(img_path):
    # Read image with opencv
    img = cv2.imread(img_path)
    # Convert to gray
    img = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    cv2.imwrite("clear.png", img)
    result = pytesseract.image_to_string(Image.open("clear.png"))
    os.remove("clear.png")
    return result[:6]


def waitCall(speak):
    engine = pyttsx3.init()
    engine.setProperty('rate', 125)  # setting up new voice rate

    voices = engine.getProperty('voices')  # getting details of current voice
    engine.setProperty('voice', voices[1].id)  # changing index, changes voices. 1 for female

    engine.say(speak)
    engine.runAndWait()
    engine.stop()


# Getting Id & Password...
wbk = openpyxl.load_workbook(r"Portal.xlsx", read_only=True, data_only=True)
sheet = wbk["Summary"]
password = sheet['J1'].value
identity = sheet['Z1'].value
wbk.close()

print('$$$$$$$_________Fully Developed By "Saurabh Chaudhari" ______ :)')
# Chrome Download Path
chrome_path = os.getcwd()

# Variables Setup
dtStamp = datetime.now().strftime("%d-%m-%Y  %H:%M:%S")
today = date.today()
MM = datetime.now().strftime("%m")
YYYY = datetime.now().strftime("%Y")
d1 = today.strftime("%d-%m-%Y")
colorama.init(autoreset=True)
lots, c_nums, lot_names, cells = [], [], [], []
w8 = 0  # as waiting GATE 0-Close; 1-Open; 9-Destroy & Exit the code
verify = []  # Totality check for 20,000 rupee lot
verified = ''  # To show totality checks to User
a, b, installments_total = None, None, 0  # As First & Second Installments nickname
acc, inst, inst_verify = [], [], []


def user():
    """To take User Inputs & Validate Them Instantly"""
    global lots, lot_names, a, b, w8, sh1

    # Excel Path
    wb = openpyxl.load_workbook(r"Portal.xlsx", read_only=True, data_only=True)

    # getting data from Frontend and Unpacking it
    sht, rows = frontEnd()
    try:
        int(sht)
    except ValueError:
        print(ColoredPrint('--------------->>>    Installment No. Must Be a "NUMBER" !!!', Fore.LIGHTRED_EX))
        messagebox.showerror("Error", f">>>    Installment No. Must Be a 'NUMBER' !!! & NOT --> {sht}")
        w8 = 9
        quit()
    if int(sht) == 1:
        a = 'First'
        sh1 = wb["First"]
    elif int(sht) == 2:
        b = 'Second'
        sh1 = wb["Second"]
    else:
        print(ColoredPrint('Invalid Input', Fore.LIGHTRED_EX))
        messagebox.showerror("Error", f"Invalid Input : {sht}")
        w8 = 9
        quit()

    for items in rows.split():
        try:
            int(items)
        except ValueError:
            print(ColoredPrint('--------------->>>    LOT No. Must Be a "Numeric Value" !!!', Fore.LIGHTRED_EX))
            messagebox.showerror("Error", f">>>    LOT No. Must Be a 'Numeric Value' !!! & NOT --> {items}")
            w8 = 9
            quit()
        if int(items) > 1:
            cell = 'C' + items
            cells.append(cell)
            lots.append(sh1[cell].value)
            lot_name = sh1[f"B{cell[1:]}"].value
            lot_names.append(lot_name)
        else:
            print(ColoredPrint('Try By Entering All Greater Than 1 values', Fore.LIGHTRED_EX))
            messagebox.showerror("Error", "Invalid Input : Try By Entering All 'Greater Than 1' Values")
            w8 = 9
            quit()
    print(f"{Fore.GREEN}{Style.BRIGHT}\tOK...!!")
    lots = [i for i in lots if i]  # To eliminate the :None values
    lot_names = [i for i in lot_names if i]  # To eliminate the :None values
    if len(lot_names) < 1:
        print(ColoredPrint('NO LOTs FOUND ...!!!', Fore.LIGHTRED_EX))
        messagebox.showerror("Error", f"NO LOTs FOUND ...!!!")
        w8 = 9
        quit()

    def verify_amt():
        global verified
        conn = sqlite3.connect('Portal_Data.sqlite')
        cur = conn.cursor()
        for part in lots:
            xx = part.replace(" ", "")
            xx = xx.split(',')
            q_marks = len(xx) * '?,'
            lookup = cur.execute('SELECT SUM(Denomination) FROM Portal WHERE Account_No in (' + q_marks[:-1] + ')', xx)
            for totality in lookup:
                verify.append('₹' + str(list(totality)[0]))
        cur.close()
        packed = list(zip(lot_names, verify))
        verified = [one for one in packed]
        return print(verified)

    def view_accounts():
        """To View Details of Selected Accounts"""
        global installments_total, inst, acc
        v = 0
        conn = sqlite3.connect('Portal_Data.sqlite')
        cur = conn.cursor()
        for part in lots:
            xx = part.replace(" ", "")
            xx = xx.split(',')
            print(ColoredPrint(f'------> > "{lot_names[v]}" < <------', Fore.LIGHTRED_EX))
            zoo = []
            for iii in xx:
                lookup = cur.execute(
                    'SELECT Account_No, Account_Name, Denomination, Next_Installment_Due_Date, Pending_Installment, '
                    'Advance_Installment, Month_Paid_Upto FROM Portal WHERE Account_No in (?)', (iii,))
                for jk in lookup:
                    zoo.append(jk)
            installments_total = []
            reg_sum = 0
            reg_cnt = 0
            for i in zoo:
                print('\t\t' + str(i[:-3]), '===>{' + str(i[6]) + '} Months Paid:', end="------->> >>")
                if i[6] == 60:
                    print(f"{Fore.LIGHTRED_EX}{Style.BRIGHT}  All Months Are PAID : {i[6]}", end="---->> ")
                pending = i[4]
                advance = i[5]
                amt = int(i[2])
                account = i[0]
                if pending is not None:
                    print(f"{Fore.LIGHTRED_EX}{Style.BRIGHT}  Pending Installments = [{pending}]*")
                    installments_total.append(amt * (int(pending) + 1))
                    inst.append(pending)
                    acc.append(account)
                elif advance is not None:
                    print(f"{Fore.LIGHTBLUE_EX}{Style.BRIGHT}  Advance Installments = [{advance}]*")
                    installments_total.append(amt * int(advance))
                    inst.append(advance)
                    acc.append(account)
                else:
                    print(f'{Fore.GREEN}{Style.BRIGHT}  Regular_Inst')
                    reg_sum += i[2]
                    reg_cnt += 1
                    installments_total.append(amt * 1)
            print(
                f'{Fore.LIGHTRED_EX}\n\t\t{Fore.LIGHTGREEN_EX}@@@@@ ==>> "Total A/c Numbers":        "{len(zoo)} accounts", '
                f'\n\t\t@@@@@ ==>> "Regular Installment" A/c:  "{reg_cnt} accounts" Of total------> > "Rs. {reg_sum}"')
            print(f'{Back.MAGENTA}{Style.BRIGHT}'
                  f'Total Installment SUM of all Pending + Regular + Advance is = Rupees {sum(installments_total):,}/-')
            print(
                f'{Back.MAGENTA}{Style.BRIGHT}'
                f'Total Installment SUM of all Pending + Regular + Advance is = Rupees {sum(installments_total):,}/-')
            v += 1
            inst_verify.append('₹' + str(sum(installments_total)))
        return sum(installments_total)

    verify_amt()
    view_accounts()
    instant = list(zip(lot_names, inst_verify))
    check = [ek for ek in instant]

    answer = messagebox.askyesno("Question", f"{verified}\n Before Installments.....\n\n** After Installments "
                                             f"Verification : \n'-->{check}'\n\n\nDo you wants to proceed further ? ?")
    if answer:
        w8 = 1
    else:
        w8 = 9
        quit()

    print(ColoredPrint("Let's Go...........  :)   ", Fore.GREEN))


def web():
    """Opening WebPortal"""
    global w8, browser
    print('Running...')
    option = Options()
    preferences = {"download.default_directory": chrome_path, "safebrowsing.enabled": "false"}  # Chrome download Path
    option.add_experimental_option("prefs", preferences)
    option.add_argument("--log-level=3")
    option.add_argument("--disable-gpu")
    # =========to make Chrome Visible
    option.page_load_strategy = 'eager'
    option.add_argument("--window-size=1280,720")

    url = 'https://dopagent.indiapost.gov.in/corp/AuthenticationController?FORMSGROUP_ID__=AuthenticationFG' \
          '&__START_TRAN_FLAG__=Y&__FG_BUTTONS__=LOAD&ACTION.LOAD=Y&AuthenticationFG.LOGIN_FLAG=3&BANK_ID=DOP' \
          '&AGENT_FLAG=Y '
    try:
        browser = webdriver.Chrome(options=option)
        browser.minimize_window()
        browser.get(url)
        print(ColoredPrint('Opening India Post in Background.....', Fore.LIGHTBLUE_EX))
    except Exception as e:
        print(ColoredPrint('Internet Problem\n\n --> >' + str(e), Fore.LIGHTRED_EX))
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

            time.sleep(0.7)
            browser.find_element_by_id('IMAGECAPTCHA').screenshot('captchaaa.png')
            cap = recCaptcha('captchaaa.png')
            print(cap)
            browser.find_element_by_id('AuthenticationFG.VERIFICATION_CODE').clear()
            time.sleep(0.5)
            browser.find_element_by_id('AuthenticationFG.VERIFICATION_CODE').send_keys(cap)
            time.sleep(1.2)
            browser.find_element_by_id('VALIDATE_RM_PLUS_CREDENTIALS_CATCHA_DISABLED').click()
            browser.minimize_window()
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
                browser.minimize_window()
                browser.close()
                try:
                    browser = webdriver.Chrome(options=option)
                    browser.minimize_window()
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
                browser.minimize_window()

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
    except Exception as e:
        quit(ColoredPrint(f'Internet / Website is BUSY.....!!\n{e}', Fore.LIGHTRED_EX))

    browser.find_element_by_id('Accounts').click()
    browser.minimize_window()

    t3.start()
    t1.start()

    print('Going inside Agent Enquire & Update Screen...')
    browser.find_element_by_id('Agent Enquire & Update Screen').click()
    print("We Were Waiting In The 'INDIA POST'    (^_^)")
    t1.join()

    while w8 == 0:  # ----> Infinite loop to wait
        time.sleep(1)
    if not w8 == 1:
        print(f'{Fore.RED}{Style.BRIGHT}Loop Destroyed')
        quit()

    print(ColoredPrint("  @    Paying LOTs Installments........ . . !!!  ", Fore.LIGHTYELLOW_EX))
    browser.maximize_window()
    for i in tqdm(range(len(lots)), unit='LOT', desc=f'{Fore.RED}{Style.BRIGHT}Paying Installments',
                  bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.LIGHTGREEN_EX, Fore.RESET)):
        browser.find_element_by_id('CustomAgentRDAccountFG.PAY_MODE_SELECTED_FOR_TRN').click()
        browser.find_element_by_id('CustomAgentRDAccountFG.ACCOUNT_NUMBER_FOR_SEARCH').send_keys(lots[i])
        browser.find_element_by_id('Button3087042').click()
        try:
            def select(x, y):
                """To select all check boxes by"""
                for j in range(x, y + 1):
                    browser.find_element_by_id('CustomAgentRDAccountFG.SELECT_INDEX_ARRAY[' + str(j) + ']').click()
                else:
                    browser.find_element_by_id('Action.AgentRDActSummaryAllListing.GOTO_NEXT__').click()

            for st, ed in [(0, 9), (10, 19), (20, 29), (30, 39), (40, 49), (50, 59)]:
                try:
                    select(st, ed)
                except NoSuchElementException:
                    break

        except NameError:
            pass
        browser.implicitly_wait(5)
        browser.find_element_by_id('Button26553257').click()
        browser.find_element_by_id('Button11874602').click()
        browser.maximize_window()

        print(f"\n\n   {Fore.LIGHTRED_EX}{Style.BRIGHT}{aa}------------------ > > > {bb}\n")
        print(ColoredPrint(' (^_^)  ------- >> > "Click  ENTER" ', Fore.LIGHTMAGENTA_EX))
        waitCall("Hello.. Sunil Ji!!  Enter The Number Of installments Now..!!")
        waitCall("Hello.. Pratimaa Ji!!  Enter The Number Of installments Now..!!")
        time.sleep(8)
        waitCall(
            "Attention please: If You Are 'Done' with Installments; Click On 'Pay All Saved Installments..!!' button")

        browser.implicitly_wait(120)
        browser.find_element_by_id('dpScheduleBtn')  # just for waiting
        time.sleep(1.5)
        msg = browser.find_element_by_id('MessageDisplay_TABLE').get_attribute('innerText')
        c_num = msg[53:63]  # # ------> C_Number Stored
        c_nums.append(c_num)
        browser.find_element_by_id('CustomAgentRDAccountFG.ACCOUNT_NUMBER_FOR_SEARCH').clear()

    browser.find_element_by_id('Reports').click()
    print(ColoredPrint("  Downloading Reports... !!!  ", Fore.LIGHTYELLOW_EX))
    # # ------------------- >>>>>>>>>> To Remove RDs Directory/
    PDFs = []
    for directory, _, filename in os.walk(os.getcwd() + '\\RDs'):
        PDFs.extend(filename)
    for i in PDFs:
        os.remove(os.getcwd() + '\\RDs\\' + i)

    time.sleep(0.50)

    print(ColoredPrint(' Successfully Completed Cleaning Old "RDs" folder....! !', Fore.LIGHTMAGENTA_EX))
    time.sleep(1)
    z = 0
    for nums in tqdm(c_nums, unit='Report', desc=f'{Fore.RED}{Style.BRIGHT}Downloading',
                     bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.LIGHTGREEN_EX, Fore.RESET)):
        try:
            browser.implicitly_wait(4)
            browser.find_element_by_id('CustomAgentRDAccountFG.EBANKING_REF_NUMBER').send_keys(nums)
            browser.find_element_by_id('CustomAgentRDAccountFG.INSTALLMENT_STATUS').send_keys("SUC")
            browser.find_element_by_id('SearchBtn').click()
            browser.find_element_by_id('GENERATE_REPORT').click()
            time.sleep(9.2)
            # ============== Move & Rename
            src = chrome_path + "\\" + "RDInstallmentReport" + d1 + ".pdf"
            dest = chrome_path + "\\" + "RDs" + "\\" + "@" + lot_names[z] + "__" + nums + ".pdf"
            shutil.move(src, dest)
            time.sleep(1)
            os.startfile(os.getcwd() + "\\RDs\\" + "@" + lot_names[z] + "__" + nums + ".pdf", 'open')
            browser.find_element_by_id('CustomAgentRDAccountFG.EBANKING_REF_NUMBER').clear()
            z += 1
        except Exception as ee:
            print(ColoredPrint(" Unable to Search the Report" + str(nums) + "  !!!!!!!!!!!\n" + str(ee),
                               Fore.LIGHTRED_EX))
    print(ColoredPrint("SUCCESSFUL........!!!     :)  :)  :)", Fore.LIGHTGREEN_EX))
    try:
        browser.minimize_window()
    except Exception:
        pass


def music():
    playsound('start.wav')


def update_lot_history():
    """Update Completed LOTs in our DataBase"""
    connn = sqlite3.connect('Portal_Data.sqlite')
    curr = connn.cursor()
    # cur.execute('CREATE TABLE History(LOT TEXT, Accounts TEXT, Date TEXT)')
    for i in range(len(lot_names)):
        curr.execute('INSERT OR REPLACE INTO History (LOT, Accounts, Date) VALUES (?,?,?)',
                     (lot_names[i], lots[i], dtStamp))
    connn.commit()
    curr.close()


t1 = threading.Thread(target=user)
t2 = threading.Thread(target=web)
t3 = threading.Thread(target=music)

t2.start()
# t1.start()
#
# t1.join()
start = timer()  # starts after 1st thread
t2.join()
end = timer()  # Ends after 2nd thread
if not w8 == 1:
    print(ColoredPrint("Virtual Agent Closed.....!" + str(w8), Fore.LIGHTRED_EX))
    browser.quit()
    quit()

zipped = list(zip(lot_names, c_nums))
if len(zipped) >= 1:
    update_lot_history()
    playsound('result.wav')
    for k in zipped:
        print(f"{Fore.LIGHTBLUE_EX}{Style.BRIGHT}{k}")
    print(ColoredPrint(f"Your Work Is Completed In '{end - start}' Seconds", Fore.LIGHTGREEN_EX))
    time.sleep(5)
    input('Press ENTER to EXIT or Close the Window. . .')
else:
    print(ColoredPrint(f"Task 'FAILED'.......! ! !   Due to 'Unstable' Internet Connection", Fore.LIGHTRED_EX))
    input('Press ENTER to EXIT or Close the Window. . .')
