"""Увольнение сотрудников Получаем список уволенных"""
from pyPythonRPA.Robot import bySelector, byImage, application, keyboard
from pyPythonRPA import byDesk
import json
import os
from time import sleep
import datetime
from os import listdir
from os.path import isfile, join
from xml.dom import minidom
import json
# from Sources.winlog import WinLog
from winlog import WinLog
winlog = WinLog("HCSBKKZ_robot")


class ColvirDismissal:
    def __init__(self):
        """тут можно application(/path/to/colvir.exe).start()"""
        application(r"C:\CBS_R\COLVIR.EXE").start()

    def send_pass(self):
        with open((os.path.join(os.getcwd(), 'settings.json'))) as colvir_json:
            keys_dict = json.load(colvir_json)

        login_colvir = keys_dict["login"]
        password_colvir = keys_dict["password"]
        instance_colvir = keys_dict["instance"]

        # login
        input_login = bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"win32"},{"ctrl_index":1},{"ctrl_index":1}])
        input_login.wait_appear(15)
        input_login.set_focus()
        input_login.set_text(login_colvir)
        sleep(0.5)

        # password
        input_password = bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"win32"},{"ctrl_index":1},{"ctrl_index":0}])
        input_password.set_text(password_colvir)
        sleep(0.5)

        # instance
        input_instance = bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"win32"},{"ctrl_index":4},{"ctrl_index":0},{"ctrl_index":0}])
        input_instance.set_text(instance_colvir)
        sleep(0.5)

        # press OK
        bySelector([{"title":"Вход в систему","class_name":"TfrmLoginDlg","backend":"win32"},{"title":"OK"}]).click()
        sleep(0.5)

    def set_text_prs(self):
        sleep(0.5)
        with open((os.path.join(os.getcwd(), 'settings.json'))) as colvir_json:
            code = json.load(colvir_json)
        code = code["code"]
        input_rps = bySelector([{"title":"Выбор режима","class_name":"TfrmCssMenu","backend":"win32"},{"class_name":"TClMaskEdit"}])
        input_rps.wait_appear()
        input_rps.set_focus()
        flag_set_text_prs = False
        count_set_text_prs = 0
        while not flag_set_text_prs:
            try:
                input_rps.set_edit_text(code)
                flag_set_text_prs = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} set prs text {e}")
                sleep(0.5)
                count_set_text_prs += 1
                if count_set_text_prs > 5:
                    flag_set_text_prs = True

        sleep(0.5)
        keyboard.press("enter")
        sleep(0.5)

    def downl_excel_dopka(self):
        with open((os.path.join(os.getcwd(), 'settings.json'))) as colvir_json:
            code = json.load(colvir_json)
        code = code["code"]

        flag_input_code = False
        flag_input_code_count = 0
        while not flag_input_code:
            try:
                flag_input_prs = False
                count_input_prs = 0
                while not flag_input_prs:
                    try:
                        input_rps = bySelector([{"title":"Выбор режима","class_name":"TfrmCssMenu","backend":"win32"},{"class_name":"TClMaskEdit"}])
                        input_rps.wait_appear()
                        input_rps.set_focus()
                        input_rps.set_text(code)
                        flag_input_prs = True
                    except Exception as e:
                        winlog.info(f"flag_input_prs {e}")
                        sleep(0.5)
                        count_input_prs += 1
                        if count_input_prs > 5:
                            flag_input_prs = True
                            winlog.info("flag_input_prs >>> not found")
                            raise ValueError
                sleep(0.5)
                keyboard.press("enter")
                sleep(1)
                if bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"win32"}]).is_exists():
                    flag_input_code = True
            except Exception as e:
                winlog.info(f"{datetime.datetime.today()} {'-' * 10} not entered code {e}")
                sleep(0.5)
                flag_input_code_count += 1
                if flag_input_code_count > 10:
                    flag_input_code = True
                    winlog.info(f"{datetime.datetime.today()} {'-' * 10} input code FAILED {e}")

        flag_input_random_tabel = False
        count_input_random_tabel = 0
        while not flag_input_random_tabel:
            try:
                input_random_tabel_number_for_get_access = bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"uia"},{"ctrl_index":1},{"ctrl_index":2},{"ctrl_index":0}])
                input_random_tabel_number_for_get_access.set_focus()
                sleep(0.5)
                # табельный номер нужен, просто чтобы пройти дальше и скачать отчет
                input_random_tabel_number_for_get_access.set_edit_text("014171")
                flag_input_random_tabel = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} flag_input_random_tabel {e}")
                sleep(0.5)
                count_input_random_tabel += 1
                if count_input_random_tabel > 5:
                    flag_input_random_tabel = True
                    raise ValueError("input form for random tabel number not found")
        sleep(0.5)
        ok_buton = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "uia"},
                               {"depth_start": 3, "depth_end": 3, "title": "OK", "control_type": "Button"}])
        ok_buton.click()
        layer_prs_gr4 = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "uia"}])
        layer_prs_gr4.wait_disappear(sec=300)
        sleep(1)

        # hide tools panel(bypass image conflict)
        bySelector([{"title": "Персонал", "class_name": "TfrmResPrsList", "backend": "uia"}, {"ctrl_index": 4},
                    {"ctrl_index": 2}]).click()
        sleep(0.5)
        keyboard.press("down")
        sleep(0.5)
        keyboard.press("enter")

        # report
        keyboard.press("F5")
        sleep(2)
        # Список принятых сотрудников
        bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}]).set_focus()
        sleep(0.5)
        byImage(os.path.join(os.getcwd(), "images", "set.PNG")).click()

        choose_priem = bySelector(
            [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 0},
             {"ctrl_index": 0}, {"ctrl_index": 0}])
        text_for_search = "Список уволенных сотрудников"
        choose_priem.set_text(text_for_search)

        ok_button = bySelector([{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"},
                                {"depth_start": 3, "depth_end": 3, "title": "OK", "control_type": "Button"}])
        ok_button.click()

        # it's normal
        ok_button = bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"},
                                {"depth_start": 3, "depth_end": 3, "title": "OK", "control_type": "Button"}])
        ok_button.wait_appear()
        ok_button.click()

        # filial
        filial_input = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"ctrl_index":0},{"ctrl_index":2},{"ctrl_index":0}])
        with open((os.path.join(os.getcwd(), 'settings.json'))) as colvir_json:
            code = json.load(colvir_json)
        filial_name = code["filial"]
        filial_input.wait_appear()
        filial_input.set_text(filial_name)

        # date from
        default_date = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0}])

        date_from_list = (datetime.datetime.now() - datetime.timedelta(20)).strftime('%d-%m-%Y').split('-')
        date_from = date_from_list[0] + "." + date_from_list[1] + "." + date_from_list[2][-2:]

        default_date.set_edit_text(date_from)

        date_to = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"ctrl_index":0},{"ctrl_index":1},{"ctrl_index":0}])
        today_list = today = datetime.datetime.today().strftime('%d-%m-%Y').split('-')
        today = today_list[0] + "." + today_list[1] + "." + today_list[2][-2:]

        date_to.set_edit_text(today)

        sleep(0.5)
        # OK
        ok_button = bySelector([{"title": "Параметры отчета ", "class_name": "TfrmRptPrmDialog", "backend": "uia"},
                                {"depth_start": 3, "depth_end": 3, "title": "OK", "control_type": "Button"}])
        ok_button.click()


    def dismissal_get_data_from_excel(self):
        """
        extract tabel numbers and write in json file
        """

        mypath = r"C:\Users\robot.drp\AppData\Local\Temp\4"
        # wait exist file
        file_exist = False
        while not file_exist:
            for file in os.listdir(mypath):
                if isfile(join(mypath, file)) == True and file.endswith(".xml") == True:
                    file_exist = True
                    winlog.info(f"end {file}")
                    break

        onlyfiles = [f for f in listdir(mypath) if isfile(join(mypath, f))]
        excel_priem = ""
        for file in onlyfiles:
            if file.endswith(".xml"):
                excel_priem = file
                break
        flag_parse_acces = False
        count_parse_acces = 0
        file = ""
        while not flag_parse_acces:
            try:
                file = minidom.parse(os.path.join(mypath, excel_priem))
                flag_parse_acces = True
            except Exception as e:
                winlog.warning(f"flag_parse_acces {e}")
                sleep(0.5)
                if count_parse_acces > 10:
                    flag_parse_acces = True

        items = file.getElementsByTagName('ss:Data')
        data_list = []
        for item in items:
            data_list.append(item.firstChild.data)
            winlog.info(item.firstChild.data)
        winlog.info(data_list)
        main_data_list_for_dismissal = []
        fio_counter = 1
        print(f"data list {data_list}")
        for i, data in enumerate(data_list):
            fio_date_reason_list = []
            if data == f"{fio_counter}" and not data_list[i + 1].startswith("п."):
                # fio
                fio_date_reason_list.append(data_list[i + 1])
                # date
                date_from_colvir = data_list[i + 5]
                cur_year = datetime.datetime.now().year
                date_list = date_from_colvir.split(".")
                date_for_enbek = date_list[0] + "." + date_list[1] + "." + str(cur_year)
                fio_date_reason_list.append(date_for_enbek)
                # reason
                if "расторжение трудового договора" in data_list[i + 8]:
                    reason_dismissal = data_list[i + 8].lower().split("расторжение трудового договора")[-1][1:-1]
                    fio_date_reason_list.append(reason_dismissal)
                else:
                    fio_date_reason_list.append("по соглашению сторон")
                fio_counter += 1
            if len(fio_date_reason_list) != 0:
                main_data_list_for_dismissal.append(fio_date_reason_list)

        winlog.info(main_data_list_for_dismissal)
        main_dict_for_perevod = {}
        main_dict_for_perevod["data_for_dismissal"] = main_data_list_for_dismissal

        winlog.info(f"main_dict_for_perevod {main_dict_for_perevod}")
        for each in main_dict_for_perevod["data_for_dismissal"]:
            with open(f"{os.getcwd()}\\jsons\\{each[0]}_dismissal.json", "w", encoding='utf-8') as file:
                dict_for_dismissal_json = {"data_for_dismissal": each}
                json.dump(dict_for_dismissal_json, file, ensure_ascii=False, indent=4)

        sleep(1)

        # close excel file window
        file_is_opened = True
        while file_is_opened:
            try:
                bySelector([{"class_name": "XLMAIN", "backend": "win32"}]).close()
                file_is_opened = False
            except Exception as e:
                winlog.warning(f"dismissal get data from excel file_is_opened {e}")
                sleep(0.5)

        # remove excel file
        remove_attempt = 0
        exist_file = True
        while exist_file:
            try:
                os.remove(os.path.join(mypath, excel_priem))
                exist_file = False
            except Exception as e:
                winlog.warning("+" * 20)
                winlog.warning(f"dismissal remove_attempt {e}")
                sleep(0.5)
                remove_attempt += 1
                if remove_attempt > 15:
                    exist_file = False

        winlog.info("-" * 30)
        winlog.info(f"file {os.path.join(mypath, excel_priem)} removed")

        # выбор отчета закрыть
        ch_report = bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}])
        ch_report.set_focus()
        sleep(0.5)
        ch_report.close()
        sleep(0.5)

        # window Персонал close
        window_personal = bySelector([{"title": "Персонал", "class_name": "TfrmResPrsList", "backend": "uia"}])
        window_personal.set_focus()
        sleep(0.2)
        window_personal.close()

        sleep(3)

    def finally_close(self):
        main_colvir_window_close_button = bySelector([{"title":"Банковская система COLVIR: !!! ДРП  (ROBOT_DRP/ROBOT_DRP-18)","class_name":"TfrmCssAppl","backend":"uia"},{"depth_start": 2, "depth_end": 2, "title":"Close", "control_type":"Button"}])
        try:
            main_colvir_window_close_button.wait_appear()
            main_colvir_window_close_button.set_focus()
            sleep(0.5)
            main_colvir_window_close_button.click()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} finally_close main_colvir_window {e}")
            main_colvir_window_close_button.click()

        apply_da_main_window_close = bySelector(
            [{"title": "Подтверждение", "class_name": "TMessageForm", "backend": "win32"},
             {"depth_start": 1, "depth_end": 1, "title": "&Да"}])
        apply_da_main_window_close.wait_appear()
        apply_da_main_window_close.set_focus()
        apply_da_main_window_close.click()


def colvir_dismissal():
    driver = ColvirDismissal()
    driver.send_pass()
    driver.downl_excel_dopka()
    driver.dismissal_get_data_from_excel()
    driver.finally_close()
    sleep(3)