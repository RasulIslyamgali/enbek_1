"""Дополнительное соглашение"""
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


class ColvirDop:
    def __init__(self):
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
                        input_rps = bySelector(
                            [{"title": "Выбор режима", "class_name": "TfrmCssMenu", "backend": "win32"},
                             {"class_name": "TClMaskEdit"}])
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
                if bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "win32"}]).is_exists():
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
        # Кадровые изменения (переводы сотрудников) ЖССБ
        bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}]).set_focus()
        sleep(0.5)
        byImage(os.path.join(os.getcwd(), "images", "set.PNG")).click()

        choose_priem = bySelector(
            [{"title": "Фильтр", "class_name": "TfrmFilterParams", "backend": "uia"}, {"ctrl_index": 0},
             {"ctrl_index": 0}, {"ctrl_index": 0}])
        text_priem = "Кадровые изменения (переводы сотрудников) ЖССБ"
        choose_priem.set_text(text_priem)

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
        today_list = datetime.datetime.today().strftime('%d-%m-%Y').split('-')
        today = today_list[0] + "." + today_list[1] + "." + today_list[2][-2:]

        date_to.set_edit_text(today)

        sleep(0.5)
        # OK
        ok_button = bySelector([{"title": "Параметры отчета ", "class_name": "TfrmRptPrmDialog", "backend": "uia"},
                                {"depth_start": 3, "depth_end": 3, "title": "OK", "control_type": "Button"}])
        ok_button.click()

    def perevod_get_data_from_excel(self):
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
                    winlog.info(f"{datetime.datetime.today()} {'-' * 10} end {file}")
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
                file = minidom.parse((mypath + os.sep + excel_priem))
                flag_parse_acces = True
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} flag_parse_acces {e}")
                sleep(0.5)
                if count_parse_acces > 10:
                    flag_parse_acces = True

        items = file.getElementsByTagName('ss:Data')
        data_list = []
        for item in items:
            data_list.append(item.firstChild.data)
        winlog.info(data_list)

        # TODO Get data from excel for perevod
        tabel_numbers_count = 1
        main_dict_for_perevod = {}
        list_for_main_dict = []
        data_perevod_dict = {}

        for i, data in enumerate(data_list):
            if data == f"{tabel_numbers_count}":
                if data_list[i + 1] != "2":
                    tabel_number = data_list[i + 1]
                    filial_upr_dolzhnost = []
                    filial_upr_dolzhnost = [
                        data_list[i + 3],
                        data_list[i + 4],
                        data_list[i + 5],
                        data_list[i + 7],
                        data_list[i + 8],
                        data_list[i + 9],
                        data_list[i + 10],
                    ]
                    tabel_numbers_count += 1
                    if data_list[i + 3] == data_list[i + 8] and data_list[i + 4] == data_list[i + 9] and data_list[
                        i + 5] == data_list[i + 10]:
                        continue
                    data_perevod_dict[tabel_number] = filial_upr_dolzhnost

        list_for_main_dict.append(data_perevod_dict)
        main_dict_for_perevod["data_for_perevod"] = list_for_main_dict

        # change_dict = json.dumps(main_dict_for_perevod, indent=4)

        with open(f"data_for_perevod.json", "w", encoding="utf-8") as file:
            json.dump(main_dict_for_perevod, file, ensure_ascii=False, indent=4)

        sleep(5)
        sleep(1)

        # close excel file window
        file_is_opened = True
        while file_is_opened:
            try:
                bySelector([{"class_name": "XLMAIN", "backend": "win32"}]).close()
                file_is_opened = False
            except Exception as e:
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10} file_is_opened {e}")
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

        winlog.info("-" * 30)
        winlog.info(f"file {mypath + os.sep + excel_priem} removed")

        # close unnecessary windows
        try:
            unn_window = bySelector([{"title": "frmXlsUntBaseForm", "class_name": "TfrmXmlUntForm", "backend": "uia"}])
            unn_window.set_focus()
            sleep(0.5)
            unn_window.close()
        except Exception as e:
            winlog.warning(f"{datetime.datetime.today()} {'-' * 10} perevod_get_data_from_excel unn_window {e}")
            winlog.warning("unn_window perevod_get_data_from_excel")

        # выбор отчета закрыть
        ch_report = bySelector([{"title": "Выбор отчета", "class_name": "TfrmRptLstRefer", "backend": "uia"}])
        ch_report.set_focus()
        sleep(0.5)
        ch_report.close()
        sleep(0.5)

        # window Персонал close
        window_personal = bySelector([{"title":"Персонал","class_name":"TfrmResPrsList","backend":"uia"}])
        window_personal.set_focus()
        sleep(0.2)
        window_personal.close()

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

    def get_iin_with_tabel_number_perevod(self, tabel_number):
        self.set_text_prs()

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
                    winlog.warning("input form tabel number not found")

        sleep(0.5)
        btn = bySelector([{"title": "PRS_GR4", "class_name": "TfrmFilterParams", "backend": "win32"},
                          {"depth_start": 3, "depth_end": 3, "title": "OK"}])
        btn.click()

        # wait close PRSGR4 window
        prs_gr4_perevod = bySelector([{"title":"PRS_GR4","class_name":"TfrmFilterParams","backend":"uia"}])
        prs_gr4_perevod.wait_disappear()

        # open client window
        sleep(1)
        keyboard.press("F3")

        # TODO run to page about person
        try:
            el = bySelector([{"title": "Карточка сотрудника", "class_name": "TfrmaResPrsDtl", "backend": "win32"}])
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

        byImage(os.path.join(os.getcwd(), "images", "ship_image.PNG")).click()
        flag_appear_prikazy_window = False
        flag_appear_prikazy_window_count = 0
        while not flag_appear_prikazy_window:
            if bySelector([{"title":"Приазы сотрудника","class_name":"TfrmHROrdLst","backend":"uia"}]).is_exists():
                flag_appear_prikazy_window = True
            else:
                sleep(0.5)
                flag_appear_prikazy_window_count += 1
                if flag_appear_prikazy_window_count > 20:
                    byImage(os.path.join(os.getcwd(), "images", "ship_image.PNG")).click()
                    sleep(0.5)
                if flag_appear_prikazy_window_count > 40:
                    flag_appear_prikazy_window = True
                    raise ValueError("Image ship_image was not click")
        sleep(1)

        keyboard.press("page down")
        sleep(0.7)
        keyboard.press("enter")

        flag_appear_window_one_prikaz = False
        flag_appear_window_one_prikaz_count = 0

        while not flag_appear_window_one_prikaz:
            if bySelector([{"title":"Приказ","class_name":"TfrmHROrdDtl","backend":"uia"}]).is_exists():
                flag_appear_window_one_prikaz = True
            else:
                sleep(0.5)
                flag_appear_window_one_prikaz_count += 1
                if flag_appear_window_one_prikaz_count > 15:
                    flag_appear_window_one_prikaz = True
                    raise ValueError("Window Prikaz do not appear. Maybe keyboard worked wrong")

        type_work = bySelector([{"title":"Приказ","class_name":"TfrmHROrdDtl","backend":"uia"},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":0},{"ctrl_index":5},{"ctrl_index":0},{"ctrl_index":21},{"ctrl_index":0}])

        get_type_work_flag = False
        get_type_work_flag_count = 0
        type_work_text = ""

        while not get_type_work_flag:
            try:
                bySelector([{"title":"Приказ","class_name":"TfrmHROrdDtl","backend":"uia"}]).wait_appear()
                type_work_text = type_work.get_value()
                get_type_work_flag = True
            except Exception as e:
                winlog.info(e)
                bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "uia"}]).close()
                bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "uia"}]).wait_disappear()
                sleep(1)
                keyboard.press("up")
                sleep(1)
                keyboard.press("enter")
                get_type_work_flag_count += 1
                if get_type_work_flag_count > 30:
                    get_type_work_flag = True
                    raise ValueError("Not Found Prikaz for perevod")
        print(type_work_text)
        print(type(type_work_text))
        sleep(0.5)

        with open((os.path.join(os.getcwd(), 'data_for_perevod.json')), encoding="utf-8") as colvir_json:
            code = json.load(colvir_json)

        data_perevod_dict = code['data_for_perevod'][0]
        data_perevod_dict[tabel_number].append(type_work_text)

        # close Prikaz window
        bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "uia"}]).close()
        bySelector([{"title": "Приказ", "class_name": "TfrmHROrdDtl", "backend": "uia"}]).wait_disappear()
        sleep(1)
        # close Prikazy window
        bySelector([{"title": "Приазы сотрудника", "class_name": "TfrmHROrdLst", "backend": "uia"}]).close()
        bySelector([{"title": "Приазы сотрудника", "class_name": "TfrmHROrdLst", "backend": "uia"}]).wait_disappear()

        sleep(1)
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
                winlog.warning(f"{datetime.datetime.today()} {'-' * 10}flag_iden_data {e}")
                sleep(0.5)
                count_iden_data += 1
                byImage(os.path.join(os.getcwd(), "images", "man.PNG")).click()
                if count_iden_data > 5:
                    flag_iden_data = True

        # TODO get ИИН
        el = bySelector([{"class_name": "TfrmFJCliFizDtl", "backend": "uia"}, {"ctrl_index": 0}, {"ctrl_index": 0},
                         {"ctrl_index": 1}, {"ctrl_index": 0}, {"ctrl_index": 0}, {"ctrl_index": 2},
                         {"ctrl_index": 0},
                         {"ctrl_index": 3}, {"ctrl_index": 0}])
        el.wait_appear()
        text_iin = el.get_value()

        print("text_iin", text_iin)

        data_perevod_dict[tabel_number].append(text_iin)

        dict_for_dopka = {"data_for_perevod": [{f"{tabel_number}": data_perevod_dict[tabel_number]}]}

        with open(f"{os.getcwd()}\\jsons\\{tabel_number}_perevod.json", "w", encoding='utf-8') as file:
            json.dump(dict_for_dopka, file, ensure_ascii=False, indent=4)
        #
        # change_dict = json.dumps(code, indent=4)
        #
        # with open("data_for_perevod.json", "w") as file:
        #     file.write(change_dict)

        sleep(1)

        # close window Карточка физического лица ...
        window_cart_fiz = bySelector([{"depth_start": 0, "depth_end": 0, "class_name":"TfrmFJCliFizDtl","backend":"uia"}])
        window_cart_fiz.close()
        window_cart_fiz.wait_disappear()

        # close window Карточка сотрудника
        window_cart_sotrudnik = bySelector([{"title":"Карточка сотрудника","class_name":"TfrmaResPrsDtl","backend":"uia"}])
        window_cart_sotrudnik.close()
        window_cart_sotrudnik.wait_disappear()

        # close window Персонал
        window_personal = bySelector([{"title":"Персонал","class_name":"TfrmResPrsList","backend":"uia"}])
        window_personal.close()
        window_personal.wait_disappear()
        sleep(0.5)

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


def run_colvir_dop():
    driver = ColvirDop()
    driver.send_pass()
    # driver.downl_excel_dopka()
    # driver.perevod_get_data_from_excel()

    with open((os.path.join(os.getcwd(), 'data_for_perevod.json')), encoding="utf-8") as colvir_json:
        code = json.load(colvir_json)

    data_perevod_dict = code['data_for_perevod'][0]
    if code['data_for_perevod'][0]:
        for i, tabel_num in enumerate(data_perevod_dict):
            driver.get_iin_with_tabel_number_perevod(tabel_num)
    else:
        winlog.info(f"{datetime.datetime.today()} {'-' * 10} No one employee has been moved, when get data from Colvir")
    driver.finally_close()
    sleep(3)




