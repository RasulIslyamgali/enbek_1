import win32com.client
import schedule
import time
import datetime
import pathlib, os
import asyncio
from multiprocessing import Process
from pyPythonRPA.Robot import bySelector
import keyboard
from multiprocessing import Process
import json

# from Sources.winlog import WinLog
from winlog import WinLog
winlog = WinLog("HCSBKKZ_robot")


def info_for_outlook():
    """
    :info for daily send report to outlook
    """

    dir_to_jsons_done = os.path.join(os.getcwd(), "jsons", "done")
    dir_to_jsons_failed = os.path.join(os.getcwd(), "jsons", "failed")

    done_jsons = list(pathlib.Path(dir_to_jsons_done).glob('*.json'))
    failed_jsons = list(pathlib.Path(dir_to_jsons_failed).glob('*.json'))

    count_perevod_done = 0
    count_perevod_failed = 0
    count_priem_done = 0
    count_priem_failed = 0
    count_dismissal_done = 0
    count_dismissal_failed = 0

    priem_who_is_failed = []
    perevod_who_is_failed = []
    dismissal_who_is_failed = []

    for file in done_jsons:
        if "perevod" in str(file):
            count_perevod_done += 1
        elif "priem" in str(file):
            count_priem_done += 1
        elif "dismissal" in str(file):
            count_dismissal_done += 1

    for file in failed_jsons:
        if "perevod" in str(file):
            count_perevod_failed += 1
            with open(file, encoding="utf-8") as file_:
                data = json.load(file_)
                data_list = data["data_for_perevod"][0]
                for t_num in data_list:
                    who_failed = "IIN " + data_list[t_num][-1]
                    perevod_who_is_failed.append(who_failed)
        elif "priem" in str(file):
            count_priem_failed += 1
            with open(file, encoding="utf-8") as file_:
                data = json.load(file_)
                who_failed = "IIN " + data["IIN"]
                priem_who_is_failed.append(who_failed)
        elif "dismissal" in str(file):
            count_dismissal_failed += 1
            with open(file, encoding="utf-8") as file_:
                data = json.load(file_)
                who_failed = "ФИО: " + data["data_for_dismissal"][0] + "\n" + " Предпологаемая причина: " + data["FAILED_REASON_DISMISSAL"].split("REASON:")[-1].split("EXCEPTION:")[0].strip()
                dismissal_who_is_failed.append(who_failed)

    print("count_perevod_failed", count_perevod_failed)
    print("count_priem_failed", count_priem_failed)
    print("count_perevod_done", count_perevod_done)
    print("count_dismissal_failed", count_dismissal_failed)
    print("count_priem_done", count_priem_done)
    print("count_dismissal_done", count_dismissal_done)

    print("priem_who_is_failed", priem_who_is_failed)
    print("perevod_who_is_failed", perevod_who_is_failed)
    print("dismissal_who_is_failed", dismissal_who_is_failed)

    outlook = win32com.client.Dispatch('outlook.application')

    mail = outlook.CreateItem(0)

    text_failed_priem = ""
    for employ in priem_who_is_failed:
        text_failed_priem += employ + "\n"

    text_failed_perevod = ""
    for employ in perevod_who_is_failed:
        text_failed_perevod += employ + "\n"

    text_failed_dismissal = ""
    for employ in dismissal_who_is_failed:
        text_failed_dismissal += employ + "\n"

    print(text_failed_priem)
    print(text_failed_perevod)
    print(text_failed_dismissal)

    mail.To = 'madieva.an@hcsbk.kz'
    mail.Subject = 'Enbek Colvir Robot'
    today = datetime.datetime.today().strftime("%Y-%m-%d")
    mail.HTMLBody = f'<h3>Отчет за {today}</h3>'
    mail.Body = f"""Создание договоров к приему на работу:
                    успешно: {count_priem_done}
                    вызвана ошибка: {count_priem_failed}
                    Подробнее:
                    {text_failed_priem}

                    Внутренние переводы:
                    успешно: {count_perevod_done}
                    вызвана ошибка: {count_perevod_failed}
                    Подробнее:
                    {text_failed_perevod}

                    Увольнения:
                    успешно: {count_dismissal_done}
                    вызвана ошибка: {count_dismissal_failed}
                    Подробнее:
                    {text_failed_dismissal}
"""
    # mail.Attachments.Add('c:\\sample.xlsx')
    # mail.Attachments.Add('c:\\sample2.xlsx')
    mail.CC = 'robot.drp@hcsbk.kz'

    mail.Send()


# schedule.every().day.at("08:00").do(info_for_outlook)
schedule.every().day.at("09:00").do(info_for_outlook)

print("INFO OUTLOOK START FIRST PART")

while True:
    print("INFO OUTLOOK START SECOND PART")
    schedule.run_pending()
    time.sleep(1)