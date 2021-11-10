import psutil
import smtplib
import datetime
import getpass
import json


USER = getpass.getuser()

proc_count = 0
proc_name = "HCSBKKZ_robot.exe"

send_from = "rasulds123@gmail.com"
send_to = "rasul.islyamgali@gmail.com"
password = "Rgbrands1289"




def send_message(send_from, send_to, BODY):
    server = smtplib.SMTP('smtp.gmail.com: 587')

    server.starttls()

    server.login(send_from, password)

    server.sendmail(send_from, send_to, BODY)

    server.quit()


while True:
    for proc in psutil.process_iter():
        try:
            # Get process name & pid from process object.
            processName = proc.name()
            if processName == proc_name:
                proc_count += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            pass





    print("proc_count", proc_count)
    if proc_count < 5:
        with open("FAILED_PROCESSES.json", encoding="utf-8") as file:
            failed_proc_str = json.load(file)["FAILED_PROC"]

        curr_time = datetime.datetime.today().strftime("%Y-%m-%d %H:%M:%S")

        title = f'SOMETHING HAPPENED WITH {proc_name[:len(proc_name) - 4]}'

        message_text = f"{title}\n\n" \
                       f"{curr_time}\n\n" \
                       f"USER: {USER}\n\n" \
                       f"ROBOT: {proc_name[:len(proc_name) - 4]}\n\n" \
                       f"FAILED PROCESSES: {failed_proc_str}"

        BODY = "\r\n".join((
            "From: %s" % send_from,
            "To: %s" % send_to,
            "Subject: %s" % title,
            "",
            message_text
        ))

        print(f"WAS SEND MESSAGE TO: {send_to}")
        send_message(send_from=send_from, send_to=send_to, BODY=BODY)
        break

    proc_count = 0






