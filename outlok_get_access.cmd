cd %~dp0Sources
copy /Y ..\Resources\Python_for_RPA\Resources\python-3.7.7.amd64\python.exe ..\Resources\Python_for_RPA\Resources\python-3.7.7.amd64\HCSBKKZ_robot.exe
..\Resources\Python_for_RPA\Resources\python-3.7.7.amd64\HCSBKKZ_robot.exe -m outlook_access_allow
pause >nul