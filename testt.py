import requests
from time import sleep
import xlrd
import openpyxl
from pathlib import Path
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
import logging
import glob
import pathlib
import getpass
from Sources.winlog import WinLog

winlog = WinLog("HCSBKKZ_robot")


