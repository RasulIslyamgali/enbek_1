from pyPythonRPA.Robot import bySelector, keyboard, application, byImage
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


class Colvir:
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

    def downl_excel_priem(self):
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
                input_random_tabel_number_for_get_access = bySelector(
                    [{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 1},
                     {"ctrl_index": 2}, {"ctrl_index": 0}])
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
        ok_buton = bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"uia"},{"depth_start": 3, "depth_end": 3, "title":"OK", "control_type":"Button"}])
        ok_buton.click()
        layer_prs_gr4 = bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"uia"}])
        layer_prs_gr4.wait_disappear(sec=300)
        sleep(1)

        # hide tools panel(bypass image conflict)
        bySelector([{"title":"Персонал","class_name":"TfrmResPrsList","backend":"uia"},{"ctrl_index":4},{"ctrl_index":2}]).click()
        sleep(0.5)
        keyboard.press("down")
        sleep(0.5)
        keyboard.press("enter")

        # report
        keyboard.press("F5")
        sleep(2)
        # Список принятых сотрудников
        bySelector([{"title":"Выбор отчета","class_name":"TfrmRptLstRefer","backend":"uia"}]).set_focus()
        sleep(0.5)
        byImage(os.path.join(os.getcwd(), "images", "set.PNG")).click()
        choose_priem = bySelector([{"title":"Фильтр","class_name":"TfrmFilterParams","backend":"uia"},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0}])
        winlog.info("why???")
        text_priem = "Список принятых сотрудников"
        choose_priem.set_edit_text(text_priem)

        ok_button = bySelector([{"title":"Фильтр","class_name":"TfrmFilterParams","backend":"uia"},{"depth_start": 3, "depth_end": 3, "title":"OK", "control_type":"Button"}])
        ok_button.click()

        # it's normal
        ok_button = bySelector([{"title":"Выбор отчета","class_name":"TfrmRptLstRefer","backend":"uia"},{"depth_start": 3, "depth_end": 3, "title":"OK", "control_type":"Button"}])
        ok_button.wait_appear()
        ok_button.click()

        # filial
        filial_input = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0}])
        with open((os.path.join(os.getcwd(), 'settings.json'))) as colvir_json:
            code = json.load(colvir_json)
        filial_name = code["filial"]
        filial_input.wait_appear()
        filial_input.set_text(filial_name)

        # date from
        default_date = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"ctrl_index":0},{"ctrl_index":1},{"ctrl_index":0}])

        date_from_list = (datetime.datetime.now() - datetime.timedelta(30)).strftime('%d-%m-%Y').split('-')
        date_from = date_from_list[0] + "." + date_from_list[1] + "." + date_from_list[2][-2:]

        default_date.set_edit_text(date_from)
        sleep(0.5)
        # OK
        ok_button = bySelector([{"title":"Параметры отчета ","class_name":"TfrmRptPrmDialog","backend":"uia"},{"depth_start": 3, "depth_end": 3, "title":"OK", "control_type":"Button"}])
        ok_button.click()

    def priem_get_data_from_excel(self):
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
        flag_parse_acces = False
        count_parse_acces = 0
        file = ""
        while not flag_parse_acces:
            try:
                file = minidom.parse((mypath + os.sep + excel_priem))
                flag_parse_acces = True
            except Exception as e:
                winlog.warning(f"flag_parse_acces {e}")
                sleep(0.5)
                if count_parse_acces > 10:
                    flag_parse_acces = True

        items = file.getElementsByTagName('ss:Data')
        # items = file.getElementsByTagName("Cell")
        data_list = []
        for item in items:
            data_list.append(item.firstChild.data)
        # print(data_list)
        tabel_number_list = []
        for i, data in enumerate(data_list):
            if data.startswith("Производственный") or data.startswith("Вспомогательный"):
                tabel_number_list.append(data_list[i + 1])
        # print(tabel_number_list)  # ['014121', '014171', '014169', '014117', '014161', '014107', '014106', '014123', '014104', '014124', '014089', '014088', '014096', '014103', '014102', '014090']

        list_tabel_number = []
        dict_tabel_number = {}
        for i, elem in enumerate(tabel_number_list):
            dict_tabel_number[f"table_number{i + 1}"] = elem
        list_tabel_number.append(dict_tabel_number)

        dict_for_json = {}
        dict_for_json["all_tabel_numbers"] = list_tabel_number

        change_dict = json.dumps(dict_for_json, indent=4)

        with open("testt.json", "w", encoding="utf-8") as file:
            file.write(change_dict)
        sleep(1)
        with open((os.path.join(os.getcwd(), 'testt.json'))) as colvir_json:
            code = json.load(colvir_json)
        if code["all_tabel_numbers"][0]:
            t_num = code["all_tabel_numbers"][0]["table_number1"]  # 014121
        else:
            winlog.info(f"{__class__} Список принятых на работу пуст")

        # close excel file window
        file_is_opened = True
        while file_is_opened:
            try:
                bySelector([{"class_name":"XLMAIN","backend":"win32"}]).close()
                file_is_opened = False
            except Exception as e:
                winlog.warning(f"file_is_opened {e}")
                sleep(0.5)

        # remove excel file
        remove_attempt = 0
        exist_file = True
        while exist_file:
            try:
                os.remove(mypath + os.sep + excel_priem)
                exist_file = False
            except Exception as e:
                winlog.warning("+" * 20)
                winlog.warning(f"{e}")
                sleep(0.5)
                remove_attempt += 1
                if remove_attempt > 15:
                    exist_file = False

        winlog.info("-"*30)
        winlog.info(f"file {mypath + os.sep + excel_priem} removed")

        # close unnecessary windows
        try:
            unn_window = bySelector([{"title": "frmXlsUntBaseForm", "class_name": "TfrmXmlUntForm", "backend": "uia"}])
            unn_window.set_focus()
            sleep(0.5)
            unn_window.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} unn_window {e}")

        # выбор отчета закрыть
        ch_report = bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}])
        ch_report.set_focus()
        sleep(0.5)
        ch_report.close()
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
                input_rps.set_text(code)
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

    def get_data_from_colvir_with_tabel_numbers(self):
        self.set_text_prs()

        with open((os.path.join(os.getcwd(), 'testt.json'))) as colvir_json:
            code = json.load(colvir_json)
        code = code["all_tabel_numbers"]

        tabel_number = []
        for i, t_num in enumerate(code[0]):
            number = code[0][t_num]  # 014121
            tabel_number.append(number)
            if i > 1:
                break
        return tabel_number

    def get_data_from_colvir(self, tabel_number):
        list_data_with_tabel = []
        dict_data_with_tabel = {}

        el = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "win32"}, {"ctrl_index": 1},
                         {"ctrl_index": 4}, {"ctrl_index": 0}])
        el.wait_appear()
        el.set_focus()
        flag_tabel_number = False
        count_tabel_number = 0
        while not flag_tabel_number:
            try:
                el.set_edit_text(tabel_number)
                flag_tabel_number = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} flag_tabel_number {e}")
                sleep(0.5)
                count_tabel_number += 1
                if count_tabel_number > 5:
                    flag_tabel_number = True
                    winlog.info("input form tabel number not found")

        sleep(0.5)
        btn = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "win32"},
                          {"depth_start": 3, "depth_end": 3, "title": "OK"}])
        btn.click()

        # wait close PRSGR4 window
        prs_gr4_perevod = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "uia"}])
        prs_gr4_perevod.wait_disappear()

        # open client window
        sleep(1)
        keyboard.press("F3")

        # TODO run to page about person
        try:
            el = bySelector([{"title":"Карточка сотрудника","class_name":"TfrmaResPrsDtl","backend":"win32"}])
            el.wait_appear()
            el.set_focus()
            sleep(1)
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} Карточка сотрудника {e}")
            keyboard.press("F3")
            el = bySelector([{"title": "Карточка сотрудника", "class_name": "TfrmaResPrsDtl", "backend": "win32"}])
            el.wait_appear()
            el.set_focus()
            sleep(1)


        # TODO GET PODRAZDELENYE
        el = bySelector(
            [{"title": "Карточка сотрудника", "class_name": "TfrmaResPrsDtl", "backend": "uia"}, {"ctrl_index": 0},
             {"ctrl_index": 1}, {"ctrl_index": 1}, {"ctrl_index": 0}, {"backend": "uia"}, {"ctrl_index": 0},
             {"ctrl_index": 16}, {"ctrl_index": 0}])
        flag_podrazdelenye = False
        count_podrazdelenye = 0
        while not flag_podrazdelenye:
            try:
                text = el.get_value()
                winlog.info(text)
                dict_data_with_tabel["podrazdelenye"] = text
                flag_podrazdelenye = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} flag_podrazdelenye {e}")
                sleep(0.5)
                count_podrazdelenye += 1
                if count_podrazdelenye > 5:
                    flag_podrazdelenye = True
                    winlog.info("flag_podrazdelenye >>>>> not found")
        sleep(0.5)

        # TODO Get Department
        upravlenie = bySelector([{"title":"Карточка сотрудника","class_name":"TfrmaResPrsDtl","backend":"uia"},{"ctrl_index":0},{"ctrl_index":1},{"ctrl_index":1},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":17},{"ctrl_index":0}])

        with open("departments.json", encoding="utf-8") as file:
            departments = json.load(file)

        flag_get_department = False
        flag_get_department_count = 0
        while not flag_get_department:
            try:
                department = upravlenie.get_value()[:4]
                department_ = departments[department]
                dict_data_with_tabel["department"] = department_
                flag_get_department = True
            except Exception as e:
                winlog.info(f"{datetime.datetime.today()} {'-' * 10} exception get department {e}")
                sleep(0.5)
                if flag_get_department_count > 10:
                    flag_get_department = True
                    winlog.info("FAILED Get Department")


        # running man picture
        byImage(os.path.join(os.getcwd(), "images", "man.PNG")).click()



        # Идентификационные данные
        flag_iden_data = False
        count_iden_data = 0
        while not flag_iden_data:
            try:
                el = bySelector([{"class_name": "TfrmFJCliFizDtl", "backend": "uia"},
                                 {"depth_start": 5, "depth_end": 5, "title": "Идентификационные данные",
                                  "control_type": "TabItem"}])
                el.wait_appear()
                el.set_focus()
                sleep(0.5)
                el.click()
                flag_iden_data = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} flag_iden_data {e}")
                sleep(0.5)
                count_iden_data += 1
                byImage(os.path.join(os.getcwd(), "images", "man.PNG")).click()
                if count_iden_data > 5:
                    flag_iden_data = True

        # TODO get ИИН
        el = bySelector([{"class_name": "TfrmFJCliFizDtl", "backend": "uia"}, {"ctrl_index": 0}, {"ctrl_index": 0},
                         {"ctrl_index": 1}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2}, {"ctrl_index": 0},
                         {"ctrl_index": 3}, {"ctrl_index": 0}])
        text = el.get_value()
        dict_data_with_tabel["IIN"] = text
        dict_data_with_tabel["srok_dogovor"] = "на определенный срок не менее одного года"
        dict_data_with_tabel["work_type"] = "основная работа"

        winlog.info(text)

        # CLOSE UNNECESSARY WINDOW
        el = bySelector([{"class_name": "TfrmFJCliFizDtl", "backend": "uia"}])
        el.set_focus()
        sleep(0.3)
        el.close()
        sleep(0.3)
        # PRIKAZ
        el = bySelector([{"title": "Карточка сотрудника", "class_name": "TfrmaResPrsDtl", "backend": "win32"}])
        el.wait_appear()
        el.set_focus()
        sleep(0.5)
        byImage(os.path.join(os.getcwd(), "images", "prikaz.PNG")).click()

        el = bySelector([{"title": "Приазы сотрудника", "class_name": "TfrmHROrdLst", "backend": "win32"}])
        el.wait_appear()
        sleep(0.5)
        keyboard.press("F3")

        # TODO GET DOGOVOR NUMBER
        sleep(0.2)
        prikaz_window = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}])
        prikaz_window.wait_appear()
        prikaz_window.set_focus()
        sleep(0.2)

        flag_dog_number = False
        count_dog_number = 0
        while not flag_dog_number:
            try:
                el = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}, {"ctrl_index": 21}])
                el.wait_appear()
                # dogovor_numver = el.get_value()
                dogovor_number = el.texts()[0]
                dict_data_with_tabel["dogovor_number"] = dogovor_number
                winlog.info(dogovor_number[0])
                flag_dog_number = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} try get dogovor number {e}")
                sleep(0.5)
                count_dog_number += 1
                if count_dog_number > 5:
                    flag_dog_number = True

        # TODO GET Дата заключения договора
        flag_date_dog = False
        count_date_dog = 0
        while not flag_date_dog:
            try:
                sleep(0.2)
                prikaz_window.set_focus()
                sleep(0.2)
                el = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}, {"ctrl_index": 66}])
                date_dogovor = el.texts()[0]
                date_format_to_enbek = date_dogovor.split(".")
                cur_year = str(datetime.date.today().year)[:2]
                date_format_to_enbek = date_format_to_enbek[0] + "." + date_format_to_enbek[1] + "." + cur_year + date_format_to_enbek[2]
                dict_data_with_tabel["date_dogovor"] = date_format_to_enbek
                winlog.info(date_dogovor[0])
                flag_date_dog =True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} date dolzhonst {e}")
                sleep(0.5)
                count_date_dog += 1
                if count_date_dog > 5:
                    flag_date_dog = True

        # TODO GET Должность
        flag_dolzhnost = False
        count_dolzhnost = 0
        while not flag_dolzhnost:
            try:
                sleep(0.2)
                prikaz_window.set_focus()
                sleep(0.2)
                el = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}, {"ctrl_index": 63}])
                dolzhnost = el.texts()[0]
                dict_data_with_tabel["dolzhnost"] = dolzhnost
                winlog.info(dolzhnost)
                flag_dolzhnost = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} dolzhnost {e}")
                count_dolzhnost += 1
                sleep(0.5)
                if count_dolzhnost > 5:
                    flag_dolzhnost = True



        # TODO GET Место выполнения работы
        flag_mesto_raboty = False
        count_mesto_raboty = 0
        while not flag_mesto_raboty:
            try:
                sleep(0.2)
                prikaz_window.set_focus()
                sleep(0.2)
                el = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}, {"ctrl_index": 53}])
                mesto_raboty = el.texts()[0]
                dict_data_with_tabel["mesto_raboty"] = mesto_raboty
                winlog.warning(mesto_raboty)
                sleep(0.2)
                flag_mesto_raboty = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} mesto_raboty {e}")
                sleep(0.5)
                count_mesto_raboty += 1
                if count_mesto_raboty > 5:
                    flag_mesto_raboty = True

        # TODO for each tabel number will create separate .json file
        with open(f"{os.getcwd()}\\jsons\\{tabel_number}_priem.json", "w", encoding='utf-8') as file:
            json.dump(dict_data_with_tabel, file, ensure_ascii=False, indent=4)

        # list_data_with_tabel.append(dict_data_with_tabel)
        # dict_for_json = {}
        # try:
        #     with open("date_for_enbek.json", "r") as f:  # reading a file
        #         dict_for_json = json.load(f)  # deserialization
        # except Exception as e:
        #     print("load json", e)
        #
        # dict_for_json[f"{tabel_number}"] = list_data_with_tabel
        #
        # change_dict = json.dumps(dict_for_json, indent=4)
        #
        # with open("date_for_enbek.json", "w") as file:
        #     file.write(change_dict)

        # CLOSE WINDOW  Приказ
        try:
            el = bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "win32"}])
            el.set_focus()
            sleep(0.2)
            el.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} CLOSE WINDOW  Приказ {e}")

        # CLOSE WINDOW Приазы сотрудника(да, в момент на 29.09.21 с ошибкой название окна)
        try:
            sleep(0.2)
            el = bySelector([{"title": "Приазы сотрудника", "class_name": "TfrmHROrdLst", "backend": "win32"}])
            el.set_focus()
            sleep(0.2)
            el.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} CLOSE WINDOW Приазы сотрудника {e}")

        # CLOSE WINDOW Карточка сотрудника
        try:
            sleep(0.2)
            el = bySelector([{"title": "Карточка сотрудника", "class_name": "TfrmaResPrsDtl", "backend": "win32"}])
            el.set_focus()
            sleep(0.2)
            el.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} CLOSE WINDOW Карточка сотрудника {e}")

        # CLOSE WINDOW Персонал
        try:
            for i in range(2):
                sleep(0.2)
                el = bySelector([{"title": "Персонал", "class_name": "TfrmResPrsList", "backend": "win32"}])
                el.set_focus()
                sleep(0.2)
                el.close()
                sleep(0.2)
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} CLOSE WINDOW Персонал {e}")


        self.set_text_prs()

    def finally_close(self):
        prsg4 = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "win32"}])
        try:
            prsg4.wait_appear()
            prsg4.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} finally_close prsg4 {e}")
            prsg4.close()

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


def colvir_priem():
    winlog.info(f"{datetime.datetime.today()} {'-' * 10} Start colvir priem")
    driver_colvir = Colvir()
    driver_colvir.send_pass()
    driver_colvir.downl_excel_priem()
    driver_colvir.priem_get_data_from_excel()
    tn = driver_colvir.get_data_from_colvir_with_tabel_numbers()

    if tn:
        for number in tn:
            driver_colvir.get_data_from_colvir(number)
    else:
        winlog.info(f"{datetime.datetime.today()} colvir_priem Список принятых на работу пуст")
    sleep(1)
    driver_colvir.finally_close()




