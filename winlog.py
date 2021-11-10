import os
import win32evtlogutil
import win32evtlog
import win32security
import win32con
import win32api


class WinLog:
    def __init__(self, app_name):
        self.appName = app_name
        self.eventCategory = 0
        pt = win32security.OpenProcessToken(win32api.GetCurrentProcess(), win32con.TOKEN_READ)
        self.sid = win32security.GetTokenInformation(pt, win32security.TokenUser)[0]
        self.eventID = 0
        self.root_path = "".join(os.getcwd().split('Robot')[0])

    def info(self, strings: list):
        self.eventID += 1
        if type(strings) is str:
            strings = [strings]
        print(*strings)
        try:
            win32evtlogutil.ReportEvent(appName=self.appName, eventID=self.eventID, eventCategory=self.eventCategory,
                                        eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE, strings=strings, sid=self.sid)
        except Exception as e:
            print(e)

    def warning(self, strings: list):
        self.eventID += 1
        if type(strings) is str:
            strings = [strings]
        print(*strings)
        try:
            win32evtlogutil.ReportEvent(appName=self.appName, eventID=self.eventID, eventCategory=self.eventCategory,
                                        eventType=win32evtlog.EVENTLOG_WARNING_TYPE, strings=strings, sid=self.sid)
        except Exception as e:
            print(e)


# example
# log = WinLog("MyRobot")
# log.info("asdasd")
# log.info(["asdasd", "asdasdasd"])
# log.warning("asdasd")
# log.warning(["asdasd", "asdasdasd"])