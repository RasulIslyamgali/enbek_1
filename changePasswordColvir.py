import os
import sys
from pyPythonRPA.Robot import bySelector, application
from pyPythonRPA.Robot import keyboard
import datetime
from time import sleep
# from Sources.winlog import WinLog
from winlog import WinLog

winlog = WinLog("HCSBKKZ_robot")


def finally_close():
    main_colvir_window_close_button = bySelector([{
        "title": "Банковская система COLVIR: ARM администратора  (COLVIR/COLVIR-18)",
        "class_name": "TfrmCssAppl", "backend": "uia"},
        {"depth_start": 2, "depth_end": 2, "title": "Close",
         "control_type": "Button"}])
    try:
        main_colvir_window_close_button.wait_appear()
        main_colvir_window_close_button.set_focus()
        sleep(0.5)
        main_colvir_window_close_button.click()
    except Exception as e:
        winlog.warning(f"finally_close main_colvir_window {e}")
        main_colvir_window_close_button.click()

    apply_da_main_window_close = bySelector(
        [{"title": "Подтверждение", "class_name": "TMessageForm", "backend": "win32"},
         {"depth_start": 1, "depth_end": 1, "title": "&Да"}])
    apply_da_main_window_close.wait_appear()
    apply_da_main_window_close.set_focus()
    apply_da_main_window_close.click()


def changePasswordColvir(new_password):
    application("C:\CBS_T_DRP\COLVIR.EXE").start()
    try:
        bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"uia"},{"ctrl_index":1},{"ctrl_index":0}]).wait_appear(10)
        bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"uia"},{"ctrl_index":0},{"ctrl_index":1}]).click()
        keyboard.write("colvir")
        bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"uia"},{"ctrl_index":0},{"ctrl_index":0}]).click()
        keyboard.write("colvir147")
        bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"uia"},{"ctrl_index":2},{"ctrl_index":0},{"title":"OK"}]).click()
        bySelector([{"title":"Выбор режима","class_name":"TfrmCssMenu","backend":"uia"},{"ctrl_index":2},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0}]).wait_appear(10)
        bySelector([{"title":"Выбор режима","class_name":"TfrmCssMenu","backend":"uia"},{"ctrl_index":2},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0}]).click()
        keyboard.write("TPSWRD")
        keyboard.press("enter")
        bySelector([{"title":"Смена пароля для COLVIR","class_name":"TfrmPswDialog","backend":"win32"},{"ctrl_index":1}]).click()
        keyboard.write(new_password)
        bySelector([{"title":"Смена пароля для COLVIR","class_name":"TfrmPswDialog","backend":"win32"},{"ctrl_index":0}]).click()
        keyboard.write(new_password)
        bySelector([{"title":"Смена пароля для COLVIR","class_name":"TfrmPswDialog","backend":"win32"},{"depth_start": 3, "depth_end": 3, "title":"OK"}]).click()
        bySelector([{"title":"Colvir Banking System","class_name":"#32770","backend":"win32"},{"ctrl_index":0},{"ctrl_index":12}]).wait_appear(5)
        bySelector([{"title":"Colvir Banking System","class_name":"#32770","backend":"win32"},{"ctrl_index":0},{"ctrl_index":12}]).click()
    except:
        exc_type, exc_obj, exc_tb = sys.exc_info()
        fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
        winlog.info(f"{exc_type}, {fname}, {exc_tb.tb_lineno}")
        winlog.info("Changing password in Colvir error")
    finally:
        finally_close()


def change_pass_colvir_launcher():
    date_today = datetime.datetime.today().strftime("%Y-%m-%d")
    today_day = date_today.split("-")[2]

    changePasswordColvir("colvir147" + date_today)
    # if today_day == "01" or today_day == "1" or today_day == "15":
    #     changePasswordColvir("colvir147" + date_today)
    # else:
    #     print("today is not 1 or 15")