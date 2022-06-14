import os
import shutil
import threading
import time
import colorama
from colorama import Fore, Style
from datetime import date, datetime
from timeit import default_timer as timer
from selenium.common.exceptions import NoSuchElementException
from tqdm import tqdm
from playsound import playsound
import SRB_funcs
from SRB_funcs import ColoredPrint, portal_Login, user, update_lot_history

print('$$$$$$$$$_________Fully Developed By "Saurabh Chaudhari" ______ :)')
# Chrome Download Path
chrome_path = os.getcwd()

# Variables Setup
dtStamp = datetime.now().strftime("%d-%m-%Y  %H:%M:%S")

today = date.today()
d1 = today.strftime("%d-%m-%Y")
colorama.init(autoreset=True)
c_nums = []

start = timer()                                 # starts the main timer


def web():
    """Login into India Post-WebPortal"""
    browser, Msg, Days, Date = portal_Login()

    t3.start()
    t1.start()

    time.sleep(1)
    browser.implicitly_wait(7)
    browser.find_element_by_id('Accounts').click()

    print('Going inside Agent Enquire & Update Screen...')
    browser.find_element_by_id('Agent Enquire & Update Screen').click()
    print("We Are Waiting... In The 'INDIA POST'    (^_^)")
    t1.join()
    global w8, lot_names
    lots, lot_names, cells, w8, verify, verified, a, b, installments_total, acc, ins, inst_verify, lot_lis, inst_no = \
        SRB_funcs.lots, SRB_funcs.lot_names, SRB_funcs.cells, SRB_funcs.w8, SRB_funcs.verify, SRB_funcs.verified, \
        SRB_funcs.a, SRB_funcs.b, SRB_funcs.installments_total, SRB_funcs.acc, SRB_funcs.inst, SRB_funcs.inst_verify, \
        SRB_funcs.lot_lis, SRB_funcs.inst_no

    while w8 == 0:  # Infinite loop to wait
        time.sleep(.5)
    if not w8 == 1:
        print(f'{Fore.RED}{Style.BRIGHT}Virtual Agent Destroyed.....!')
        quit()
    time.sleep(.5)
    print(lots, lot_names)

    # # Doing Lots
    print(ColoredPrint("  @    Paying LOTs Installments........ . . !!!  ", Fore.LIGHTYELLOW_EX))
    browser.implicitly_wait(5)
    for i in tqdm(range(len(lots)), unit=' LOT', desc=f'{Fore.LIGHTRED_EX}{Style.BRIGHT}Paying Installments',
                  bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.LIGHTGREEN_EX, Fore.RESET)):
        browser.implicitly_wait(8)
        browser.find_element_by_xpath('//*[@id="CustomAgentRDAccountFG.PAY_MODE_SELECTED_FOR_TRN"]').click()
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
        # finally:
        #     print("Done Selection")
        browser.implicitly_wait(7)
        browser.find_element_by_xpath('//*[@id="Button26553257"]').click()
        browser.find_element_by_xpath('//*[@id="Button11874602"]').click()
        browser.find_element_by_xpath('//*[@id="PAY_ALL_SAVED_INSTALLMENTS"]').click()
        msg = browser.find_element_by_id('MessageDisplay_TABLE').get_attribute('innerText')

        c_num = msg[53:63]  # C_Number Stored
        c_nums.append(c_num)
        browser.find_element_by_id('CustomAgentRDAccountFG.ACCOUNT_NUMBER_FOR_SEARCH').clear()

    browser.find_element_by_id('Reports').click()
    time.sleep(1)

    # # ------------------- >>>>>>>>>> To Remove RDs Old Directory/
    PDFs = []
    for directory, _, filename in os.walk(os.getcwd() + '\\RDs'):
        PDFs.extend(filename)
    for i in PDFs:
        os.remove(os.getcwd() + '\\RDs\\' + i)

    time.sleep(0.30)

    print(f'{Fore.LIGHTMAGENTA_EX}{Style.BRIGHT}\n\nSuccessfully Completed Cleaning Old "RDs" folder....! !\n\n')

    print(ColoredPrint("  Downloading Reports... !!!  ", Fore.LIGHTYELLOW_EX))
    z = 0
    for nums in tqdm(c_nums, unit=' Report', desc=f'{Fore.LIGHTRED_EX}{Style.BRIGHT}Downloading',
                     bar_format="{l_bar}%s{bar}%s{r_bar}" % (Fore.LIGHTGREEN_EX, Fore.RESET)):
        try:
            browser.find_element_by_id('CustomAgentRDAccountFG.EBANKING_REF_NUMBER').send_keys(nums)
            browser.find_element_by_id('CustomAgentRDAccountFG.INSTALLMENT_STATUS').send_keys("SUC")
            browser.find_element_by_id('SearchBtn').click()
            browser.find_element_by_id('GENERATE_REPORT').click()
            # Required Maximum Time to Download
            time.sleep(10)
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
            time.sleep(1.5)
    print(ColoredPrint("All 'Reports Downloaded' SUCCESSFULLY........!!!     :)  :)  :)", Fore.LIGHTGREEN_EX))
    time.sleep(0.25)
    browser.quit()


def music():
    playsound('start.wav')


t1 = threading.Thread(target=user)
t2 = threading.Thread(target=web)
t3 = threading.Thread(target=music)

t2.start()
t2.join()

# t1.start()
# t1.join()
global w8, lot_names

end = timer()  # Ends after 2nd thread
if not w8 == 1:
    quit(ColoredPrint("Virtual Agent Closed.....!" + str(w8), Fore.LIGHTRED_EX))

zipped = list(zip(lot_names, c_nums))
if len(zipped) >= 1:
    update_lot_history()
    playsound('result.wav')
    for k in zipped:
        print(f"{Fore.LIGHTBLUE_EX}{Style.BRIGHT}{k}")
    print(ColoredPrint(f"Your Work Is Completed In '{end - start}' Seconds", Fore.LIGHTGREEN_EX))
    time.sleep(1)
else:
    print(ColoredPrint(f"Task 'FAILED'.......! ! !   Due to 'Unstable' Internet Connection", Fore.LIGHTRED_EX))
    time.sleep(20)
    input('Press ENTER to EXIT or Close the Window. . .')
