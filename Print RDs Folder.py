import os
import sys
import time
from timeit import default_timer as timer

import colorama
from colorama import Fore, Back, Style


def ColoredPrint(msg, color):
    colorama.init(autoreset=True)

    return (f"{color}{Style.BRIGHT}"
            "============================================================================================\n\n"
            f"{Fore.LIGHTMAGENTA_EX}{Style.BRIGHT}                       {msg}                  \n\n"
            f"{color}{Style.BRIGHT}"
            f"============================================================================================")


def FrontPagePrint(copies):
    """ Automatic Printing Of Only '2' PAGED PDFs On FRONT sides"""
    for pdf in PDFs:
        os.system(
            f'{os.getcwd()}\\PDFtoPrinter.exe {os.getcwd()}\\RDs\\{pdf} "HP LaserJet 1020" pages=1 copies={copies}')
        time.sleep(0.5)
        print(ColoredPrint(
            f'{Fore.LIGHTBLUE_EX}Printing Only {copies}-"FRONT" Pages of....\t{pdf}          \t -----------> > > '
            f'{Fore.LIGHTRED_EX}Ok', Fore.LIGHTGREEN_EX))
    return 0


def BackPagePrint(copies):
    """ Automatic Printing Of Only '2' PAGED PDFs On BACK sides"""
    for pdf in reversed(PDFs):
        os.system(
            f'{os.getcwd()}\\PDFtoPrinter.exe {os.getcwd()}\\RDs\\{pdf} "HP LaserJet 1020" pages=2 copies={copies}')
        time.sleep(0.5)
        print(ColoredPrint(
            f'{Fore.LIGHTBLUE_EX}Printing Only {copies}-"BACK" Pages of....\t{pdf}          \t -----------> > > '
            f'{Fore.LIGHTRED_EX}Ok', Fore.LIGHTGREEN_EX))
    return print(ColoredPrint("All 'Reports Printed' SUCCESSFULLY........!!!     :)  :)  :)", Fore.LIGHTGREEN_EX))


if __name__ == "__main__":
    colorama.init(autoreset=True)
    print('$$$$$$$$________Fully Developed By "Saurabh Chaudhari" ___ :)')
    print(f"{Fore.LIGHTRED_EX}To Cancel the Printing ...Just 'CLOSE the WINDOW' Else, ")
    PDFs = []
    for directory, _, filenames in os.walk(os.getcwd() + '\\RDs'):
        PDFs.extend(filenames)
    print(f""" 
    --------------------------------------------------------------------------------------------------------------
    \nFiles To Be Printed :- {PDFs}\n
    --------------------------------------------------------------------------------------------------------------\n""")

    print(f"{Fore.LIGHTGREEN_EX}{Style.BRIGHT}\t\t\t  Printing on HP LaserJet 1020..............! ! !")
    print(f'\n>>{Fore.LIGHTRED_EX} Are You READY to Print "FRONT Side" ... ? ?')
    answer = input('------------------ >> (y/n) :')
    start = timer()
    if 1 <= len(answer) < 4 and answer[0].lower() == 'y':
        FrontPagePrint(3)
    else:
        print(2 * ColoredPrint('You Choose "MANUAL" Printing for "FRONT SIDE" PAGES ..... ! ! !', Fore.LIGHTRED_EX))

    print(f'\n\n>>{Fore.LIGHTRED_EX} Are You READY to Print "BACK Side" ... ? ?')
    ans = input('------------------ >> (y/n) :')
    if 1 <= len(ans) < 4 and ans[0].lower() == 'y':
        BackPagePrint(3)
    else:
        print(2 * ColoredPrint('You Choose "MANUAL" Printing for "BACK SIDE" PAGES ..... ! ! !', Fore.LIGHTRED_EX))
    end = timer()
    print('\n\n', ColoredPrint(f"Your Printing Is Completed In Just '{end - start}' Seconds", Fore.LIGHTGREEN_EX))

    input('\n\nPress ENTER to EXIT or Close the Window. . .')
