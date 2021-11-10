from pyPythonRPA.Robot import bySelector,  keyboard
from time import sleep
import psutil


def start_HCSBK_cmd():
    sleep(1)

    keyboard.press("win + e")

    bySelector([{"title":"File Explorer","class_name":"CabinetWClass","backend":"win32"}]).wait_appear()

    input_dir_ = bySelector([{"title":"File Explorer","class_name":"CabinetWClass","backend":"win32"},{"ctrl_index":8},{"ctrl_index":0},{"ctrl_index":4},{"ctrl_index":0},{"ctrl_index":0}])
    input_dir_.set_focus()
    sleep(0.5)
    input_dir_.click()

    sleep(1)
    keyboard.write(r"C:\Users\robot.drp\Desktop\Rasul\HCSBKKZ_robot")
    sleep(1)

    keyboard.press("enter")

    bySelector([{"title":"HCSBKKZ_robot","class_name":"CabinetWClass","backend":"win32"}]).wait_appear()

    bySelector([{"title":"HCSBKKZ_robot","class_name":"CabinetWClass","backend":"win32"}]).set_focus()

    sleep(0.5)

    # start cmd
    HCSBK_CMD = bySelector([{"title":"HCSBKKZ_robot","class_name":"CabinetWClass","backend":"uia"},{"depth_start": 5, "depth_end": 5, "title":"HCSBKKZ_robot.cmd", "control_type":"ListItem"}])
    HCSBK_CMD.double_click()

    bySelector([{"title": "HCSBKKZ_robot", "class_name": "CabinetWClass", "backend": "win32"}]).wait_appear()

    bySelector([{"title": "HCSBKKZ_robot", "class_name": "CabinetWClass", "backend": "win32"}]).set_focus()

    sleep(1)

    check_process_list = bySelector([{"title":"HCSBKKZ_robot","class_name":"CabinetWClass","backend":"uia"},{"depth_start": 5, "depth_end": 5, "title":"check_process_list.cmd", "control_type":"ListItem"}])
    check_process_list.double_click()


proc_name = "HCSBKKZ_robot.exe"
proc_count = 0

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

    if proc_count < 6:
        print("start HCSBK_CMD")
        start_HCSBK_cmd()

    proc_count = 0

