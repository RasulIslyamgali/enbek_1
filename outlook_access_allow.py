from pyPythonRPA.Robot import bySelector
import time
# from Sources.winlog import WinLog
from winlog import WinLog
winlog = WinLog("HCSBKKZ_robot")


while True:
    root = {"title": "Microsoft Outlook", "class_name": "#32770", "backend": "win32"}
    button = {"title": "Разрешить", "depth_start": 1, "depth_end": 1}
    # bySelector([root]).wait_appear(15)
    print("i work outlook access")
    if bySelector([root]).is_exists():
        bySelector([root]).set_focus()
        time.sleep(6)
        bySelector([root, button]).click()
