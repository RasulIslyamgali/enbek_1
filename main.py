# import multiprocessing
# import shedule
# import colvir
# import enbek
# from colvir import colvir_priem
# from Sources.colvir import colvir_priem
# from Sources.enbek import enbek_priem, enbek_perevod, enbek_dismissal
# from Sources.colvir_dop import run_colvir_dop
# from Sources.colvir_dismissal import colvir_dismissal
from colvir import colvir_priem
from enbek import enbek_priem, enbek_perevod, enbek_dismissal
from colvir_dop import run_colvir_dop
from colvir_dismissal import colvir_dismissal

# from Sources.Enbek_Vacancies_Robot.Sources.main_vacancies import launch_vacancies
# from Sources.changePasswordColvir import change_pass_colvir_launcher
from Enbek_Vacancies_Robot.Sources.main_vacancies import launch_vacancies
from changePasswordColvir import change_pass_colvir_launcher
import datetime
import psutil
from time import sleep
# from Sources.winlog import WinLog
from winlog import WinLog
import json

winlog = WinLog("HCSBKKZ_robot")


def kill_colvir():
    for proc in psutil.process_iter():
        if proc.name() == 'EXCEL.EXE' or proc.name() == 'COLVIR.EXE' or proc.name() == "chrome.exe":
            proc.kill()


def colvir_enbek():
    while True:
        failed_proc = ""

        curr_hour_minute = datetime.datetime.now().strftime("%H:%M")
        print("i work colvir_enbek")
        if curr_hour_minute == "05:00":
            # # # TODO priem
            try:
                colvir_priem()
                # # func ecp_priem там неправильный пароль эцп стоит ( для теста так было сделано ) и надо пару инпутов убрать
                enbek_priem()
                winlog.info(f"{datetime.datetime.today()} {'-' * 10} The end priem")
            except Exception as e:
                winlog.info(f"{e}")
                print("PRIEM FAILED")
                failed_proc += "PRIEM "

            try:
                kill_colvir()
            except Exception as e:
                print(f"{e}")

            # # TODO dopka
            try:
                run_colvir_dop()
                # тут тоже неправильный пароль по эцп и надо пару инпутов убрать
                enbek_perevod()
            except Exception as e:
                winlog.info(f"{e}")
                print("PEREVOD FAILED")
                failed_proc += "\nPEREVOD"

            try:
                kill_colvir()
            except Exception as e:
                print(f"{e}")

            winlog.info(f"{datetime.datetime.today()} {'-' * 10} The end perevod")
            #
            # # TODO dismissal
            try:
                colvir_dismissal()
                # # тут тоже неправильный пароль эцп и есть пару инпутов
                enbek_dismissal()
            except Exception as e:
                winlog.info(f"{e}")
                print("DISMISSAL FAILED")
                failed_proc += "\nDISMISSAL"

            try:
                kill_colvir()
            except Exception as e:
                print(f"{e}")

            # TODO add vacancies
            try:
                launch_vacancies()
            except Exception as e:
                winlog.info(f"{e}")
                print("VACANCIES FAILED")
                failed_proc += "\nVACANCIES"

            try:
                kill_colvir()
            except Exception as e:
                print(f"{e}")

            # TODO change colvir password ( launch 2 time ( 1 and 15 date) on month )
            today = datetime.datetime.today().strftime("%Y%m%d")[-2:]
            if today == "14" or today == "28":
                try:
                    change_pass_colvir_launcher()
                except Exception as e:
                    winlog.info(f"{e}")
                    print("CHANGE PASSWORD FAILED")
                    failed_proc += "\nCHANGE PASSWORD"

                try:
                    kill_colvir()
                except Exception as e:
                    print(f"{e}")

            if failed_proc:
                print(f"FAILED PROCESS: {failed_proc}")
                dict_failed_proc = {"FAILED_PROC": failed_proc}
                with open("FAILED_PROCESSES.json", "w", encoding="utf-8") as file:
                    json.dump(dict_failed_proc, file, ensure_ascii=False, indent=4)
                raise ValueError("RAISE ERR FOR SEND MESSAGE ABOUT TO DEVELOPER")


