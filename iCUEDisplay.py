"""
Written by Benjamin Jack Cullen aka Holographic_Sol
"""
import os
import sys
import time
import win32com.client
import win32api
import win32process
import win32con
from PyQt5.QtCore import Qt, QThread, QSize, QPoint, QCoreApplication, QObject, QTimer, QEvent
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QDesktopWidget, QLineEdit
from PyQt5.QtGui import QIcon, QCursor, QFont
from PyQt5 import QtCore
from cuesdk import CueSdk
import GPUtil
import psutil
import pythoncom
import unicodedata
import shutil


def NFD(text):
    return unicodedata.normalize('NFD', text)


def canonical_caseless(text):
    return NFD(NFD(text).casefold())


print('-- initializing:')
if hasattr(Qt, 'AA_EnableHighDpiScaling'):
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    print('-- AA_EnableHighDpiScaling: True')
elif not hasattr(Qt, 'AA_EnableHighDpiScaling'):
    print('-- AA_EnableHighDpiScaling: False')
if hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)
    print('-- AA_UseHighDpiPixmaps: True')
elif not hasattr(Qt, 'AA_UseHighDpiPixmaps'):
    print('-- AA_UseHighDpiPixmaps: False')

priority_classes = [win32process.IDLE_PRIORITY_CLASS,
                    win32process.BELOW_NORMAL_PRIORITY_CLASS,
                    win32process.NORMAL_PRIORITY_CLASS,
                    win32process.ABOVE_NORMAL_PRIORITY_CLASS,
                    win32process.HIGH_PRIORITY_CLASS,
                    win32process.REALTIME_PRIORITY_CLASS]
pid = win32api.GetCurrentProcessId()
print('-- process id:', pid)
handle = win32api.OpenProcess(win32con.PROCESS_ALL_ACCESS, True, pid)
win32process.SetPriorityClass(handle, priority_classes[4])
print('-- win32process priority class:', priority_classes[4])

creat_new_startup_bool = False
out_of_bounds = False
prev_obj_eve = []

mon_threads = []
conf_thread = []

run_startup_bool = False
start_minimized_bool = False
configuration_read_complete = False
allow_display_application = False
hdd_startup_bool = False
cpu_startup_bool = False
dram_startup_bool = False
vram_startup_bool = False
allow_mon_threads_bool = False
connected_bool = None
connected_bool_prev = None

sdk = CueSdk(os.path.join(os.getcwd(), 'bin\\CUESDK.x64_2017.dll'))
k95_rgb_platinum = []
k95_rgb_platinum_selected = 0

alpha_led = [38,
             55,
             53,
             40,
             28,
             41,
             42,
             43,
             33,
             44,
             45,
             46,
             57,
             56,
             34,
             35,
             26,
             29,
             39,
             30,
             32,
             54,
             27,
             52,
             31,
             51]
alpha_str = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
             'u', 'v', 'w', 'x', 'y', 'z']
hdd_led_color = [255, 255, 255]
hdd_led_color_off = [0, 0, 0]
hdd_led_time_on = 0.5
hdd_initiation = False
hdd_led_item = []
hdd_led_off_item = []
hdd_display_key_bool = []
i = 0
for _ in alpha_led:
    itm = {alpha_led[i]: hdd_led_color}
    hdd_led_item.append(itm)
    i += 1
i = 0
for _ in alpha_led:
    itm = {alpha_led[i]: hdd_led_color_off}
    hdd_led_off_item.append(itm)
    i += 1
for _ in alpha_led:
    hdd_display_key_bool.append(False)

cpu_stat = ()
cpu_led_color = [255, 255, 255]
cpu_led_color_off = [0, 0, 0]
cpu_led_time_on = 1.0
cpu_led_item = [({116: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 2
    ({113: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 5
    ({109: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 8
    ({103: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])})]   # /
cpu_initiation = False
cpu_display_key_bool = [False, False, False, False]
cpu_led_off_item = [({116: (0, 0, 0)}),
                ({113: (0, 0, 0)}),
                ({109: (0, 0, 0)}),
                ({103: (0, 0, 0)})]

dram_stat = ()
dram_led_color = [255, 255, 255]
dram_led_color_off = [0, 0, 0]
dram_led_time_on = 1.0
dram_led_item = [({117: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 2
    ({114: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 5
    ({110: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 8
    ({104: (dram_led_color[0], dram_led_color[1], dram_led_color[2])})]   # /
dram_initiation = False
dram_display_key_bool = [False, False, False, False]
dram_led_off_item = [({117: (0, 0, 0)}),
                ({114: (0, 0, 0)}),
                ({110: (0, 0, 0)}),
                ({104: (0, 0, 0)})]

gpu_num = ()
vram_stat = ()
vram_led_color = [255, 255, 255]
vram_led_color_off = [0, 0, 0]
vram_led_time_on = 1.0
vram_led_item = [({118: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 3
    ({115: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 6
    ({111: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 9
    ({105: (vram_led_color[0], vram_led_color[1], vram_led_color[2])})]  # keypad_asterisk
vram_led_off_item = [({118: (0, 0, 0)}),
                ({115: (0, 0, 0)}),
                ({111: (0, 0, 0)}),
                ({105: (0, 0, 0)})]
exclusive_access_bool = False
vram_initiation = False
vram_display_key_bool = [False, False, False, False]


def create_new():
    global creat_new_startup_bool
    # create config if not exist
    # create .vbs .bat & shortcut if not exist
    print('-- running: create_new')
    if not os.path.exists('./config.dat'):
        print('-- installing')
        print('-- creating new: configuration file')
        open('./config.dat', 'w').close()
        with open('./config.dat', 'a') as fo:
            fo.writelines('cpu_led_color: 255,0,255\n')
            fo.writelines('cpu_led_color_off: 0,0,0\n')
            fo.writelines('cpu_led_time_on: 1.0\n')
            fo.writelines('cpu_startup: true\n')
            fo.writelines('dram_led_color: 255,0,255\n')
            fo.writelines('dram_led_color_off: 0,0,0\n')
            fo.writelines('dram_led_time_on: 1\n')
            fo.writelines('dram_startup: true\n')
            fo.writelines('vram_led_color: 255,0,255\n')
            fo.writelines('vram_led_color_off: 0,0,0\n')
            fo.writelines('vram_led_time_on: 1\n')
            fo.writelines('vram_startup: true\n')
            fo.writelines('gpu_num: 0\n')
            fo.writelines('hdd_led_color: 255,0,255\n')
            fo.writelines('hdd_led_color_off: 0,0,0\n')
            fo.writelines('hdd_led_time_on: 0.1\n')
            fo.writelines('hdd_startup: true\n')
            fo.writelines('exclusive_access: true\n')
            fo.writelines('start_minimized: false\n')
            fo.writelines('run_startup: false')
        fo.close()

    if not os.path.exists('./iCUEDisplay.vbs') or not os.path.exists('./iCUEDisplay.bat'):
        print('-- installing')
        cwd = os.getcwd()
        print(cwd)

        path_for_in_bat = os.path.join('"' + cwd + '\\iCUEDisplay.exe"')
        path_to_bat = cwd + '\\iCUEDisplay.bat'
        path_for_in_vbs = 'WshShell.Run chr(34) & "' + path_to_bat + '" & Chr(34), 0'

        print('-- creating new: batch script')
        open('./iCUEDisplay.bat', 'w').close()

        print('-- creating new: vbs script')
        open('./iCUEDisplay.vbs', 'w').close()

        with open('./iCUEDisplay.bat', 'a') as fo:
            fo.writelines(path_for_in_bat)
        fo.close()

        with open('./iCUEDisplay.vbs', 'a') as fo:
            fo.writelines('Set WshShell = CreateObject("WScript.Shell")\n')
            fo.writelines(path_for_in_vbs + '\n')
            fo.writelines('Set WshShell = Nothing\n')
        fo.close()

        try:
            path = os.path.join(cwd + '\\iCUEDisplay.lnk')
            target = cwd + '\\iCUEDisplay.vbs'
            icon = cwd + './icon.ico'
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(path)
            shortcut.Targetpath = target
            shortcut.WorkingDirectory = cwd
            shortcut.IconLocation = icon
            shortcut.save()
        except Exception as e:
            print('-- Error in create_new:', e)

        time.sleep(1)

        print('-- checking existence of created files')
        if os.path.exists('./iCUEDisplay.exe') and os.path.exists(path_to_bat) and os.path.exists('./iCUEDisplay.vbs'):
            print('-- checking existence of created files: files exist')
            if os.path.exists('./iCUEDisplay.lnk'):
                print('-- starting program')
                os.startfile(cwd+'./iCUEDisplay.lnk')
                time.sleep(2)
                creat_new_startup_bool = True


class App(QMainWindow):
    cursorMove = QtCore.pyqtSignal(object)

    def __init__(self):
        super(App, self).__init__()
        global creat_new_startup_bool

        create_new()
        if creat_new_startup_bool is True:
            sys.exit()

        self.cursorMove.connect(self.handleCursorMove)
        self.timer = QTimer(self)
        self.timer.setInterval(50)
        self.timer.timeout.connect(self.pollCursor)
        self.timer.start()
        self.cursor = None

        self.setWindowIcon(QIcon('./icon.ico'))
        self.title = 'iCUE Display'
        print('-- setting self.title as:', self.title)
        self.setWindowTitle(self.title)

        self.width = 372
        self.height = 200
        self.prev_pos = self.pos()
        self.pos_w = ((QDesktopWidget().availableGeometry().width() / 2) - (self.width / 2))
        self.pos_h = ((QDesktopWidget().availableGeometry().height() / 2) - (self.height / 2))
        self.pos_w = int(self.pos_w)
        self.pos_h = int(self.pos_h)
        print('-- setting window dimensions:', self.width, self.height)
        print('-- setting window position:', self.pos_w, self.pos_h)
        self.setGeometry(int(self.pos_w), int(self.pos_h), self.width, self.height)

        p = self.palette()
        p.setColor(self.backgroundRole(), Qt.black)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setPalette(p)

        self.font_s8b = QFont("Segoe UI", 8, QFont.Bold)
        self.font_s9b = QFont("Segoe UI", 10, QFont.Bold)

        self.lbl_data_style = """QLabel {background-color: rgb(35, 35, 35);
           color: rgb(200, 200, 200);
           border:0px solid rgb(35, 35, 35);}"""

        self.btn_enabled_style = """QPushButton{background-color: rgb(50, 50, 50);
                   color: rgb(200, 200, 200);
                   border:1px solid rgb(35, 35, 35);}"""

        self.btn_disabled_style = """QPushButton{background-color: rgb(0, 0, 0);
                           color: rgb(200, 200, 200);
                           border:2px solid rgb(35, 35, 35);}"""

        self.qle_selected = """QLineEdit{background-color: rgb(0, 0, 0);
               color: rgb(200, 200, 200);
               border:1px solid rgb(35, 35, 35);}"""

        self.qle_unselected = """QLineEdit{background-color: rgb(0, 0, 0);
                       color: rgb(200, 200, 200);
                       border:2px solid rgb(35, 35, 35);}"""

        self.title_bar_h = 20
        self.title_bar_btn_w = (self.title_bar_h * 1.5)
        self.monitor_btn_h = 20
        self.monitor_btn_w = 72
        self.monitor_btn_pos_w = 2
        self.monitor_btn_pos_h = 25

        self.btn_title_logo = QPushButton(self)
        self.btn_title_logo.move(0, 0)
        self.btn_title_logo.resize(self.title_bar_h, self.title_bar_h)
        # self.btn_title_logo.setIcon(QIcon("./image/dev_target_25x25.png"))
        self.btn_title_logo.setIconSize(QSize(self.title_bar_btn_w, self.title_bar_btn_w))
        self.btn_title_logo.setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        self.btn_title_logo.setEnabled(False)
        print('-- created:', self.btn_title_logo)

        self.lbl_title_bg = QLabel(self)
        self.lbl_title_bg.move(self.title_bar_btn_w, 0)
        self.lbl_title_bg.resize((self.width - (self.title_bar_btn_w * 2)), self.title_bar_h)
        self.lbl_title_bg.setStyleSheet(
            """QLabel {background-color: rgb(0, 0, 0);
           border:0px solid rgb(35, 35, 35);}"""
        )
        print('-- created:', self.lbl_title_bg)

        self.lbl_title = QLabel(self)
        self.lbl_title.move((self.width / 2) - 44, 15)
        self.lbl_title.resize(86, 30)
        self.lbl_title.setFont(self.font_s9b)
        self.lbl_title.setText('iCUE-DISPLAY')
        self.lbl_title.setStyleSheet("""QLabel {background-color: rgb(0, 0, 0);
                           color: rgb(200, 200, 200);
                           border:0px solid rgb(35, 35, 35);}""")
        print('-- created:', self.lbl_title)

        self.lbl_con_stat = QLabel(self)
        # self.lbl_con_stat.move(self.width - 10, 25)
        self.lbl_con_stat.move(2, 2)
        self.lbl_con_stat.resize(8, 8)
        self.lbl_con_stat.setStyleSheet("""QLabel {background-color: rgb(255, 0, 0);
                                   color: rgb(0, 0, 0);
                                   border:2px solid rgb(35, 35, 35);}""")
        print('-- created:', self.lbl_con_stat)

        self.btn_quit = QPushButton(self)
        self.btn_quit.move((self.width - 20), 0)
        self.btn_quit.resize(20, 20)
        self.btn_quit.setIcon(QIcon("./image/img_close.png"))
        self.btn_quit.setIconSize(QSize(8, 8))
        self.btn_quit.clicked.connect(QCoreApplication.instance().quit)
        self.btn_quit.setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        print('-- created:', self.btn_quit)

        self.btn_minimize = QPushButton(self)
        self.btn_minimize.move((self.width - 40), 0)
        self.btn_minimize.resize(20, 20)
        self.btn_minimize.setIcon(QIcon("./image/img_minimize.png"))
        self.btn_minimize.setIconSize(QSize(20, 20))
        self.btn_minimize.clicked.connect(self.showMinimized)
        self.btn_minimize.setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        print('-- created self.btn_minimize', self.btn_minimize)

        self.lbl_cpu_mon = QLabel(self)
        self.lbl_cpu_mon.move(2, (self.monitor_btn_pos_h * 2))
        self.lbl_cpu_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_cpu_mon.setFont(self.font_s8b)
        self.lbl_cpu_mon.setText(' CPU')
        self.lbl_cpu_mon.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_cpu_mon)

        self.btn_cpu_mon = QPushButton(self)
        self.btn_cpu_mon.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 2))
        self.btn_cpu_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_cpu_mon.setFont(self.font_s8b)
        self.btn_cpu_mon.setFont(self.font_s8b)
        self.btn_cpu_mon.setStyleSheet(self.btn_disabled_style)
        self.btn_cpu_mon.clicked.connect(self.btn_cpu_mon_function)
        print('-- created:', self.btn_cpu_mon)

        self.btn_cpu_mon_rgb_on = QLineEdit(self)
        self.btn_cpu_mon_rgb_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_cpu_mon_rgb_on.move((self.monitor_btn_w * 2) + 6, (self.monitor_btn_pos_h * 2))
        self.btn_cpu_mon_rgb_on.setFont(self.font_s8b)
        self.btn_cpu_mon_rgb_on.returnPressed.connect(self.btn_cpu_mon_rgb_on_function)
        self.btn_cpu_mon_rgb_on.setStyleSheet(self.qle_unselected)

        self.btn_cpu_mon_rgb_off = QLineEdit(self)
        self.btn_cpu_mon_rgb_off.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_cpu_mon_rgb_off.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 2))
        self.btn_cpu_mon_rgb_off.setFont(self.font_s8b)
        self.btn_cpu_mon_rgb_off.returnPressed.connect(self.btn_cpu_mon_rgb_off_function)
        self.btn_cpu_mon_rgb_off.setStyleSheet(self.qle_unselected)

        self.btn_cpu_led_time_on = QLineEdit(self)
        self.btn_cpu_led_time_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_cpu_led_time_on.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 2))
        self.btn_cpu_led_time_on.setFont(self.font_s8b)
        self.btn_cpu_led_time_on.returnPressed.connect(self.btn_cpu_led_time_on_function)
        self.btn_cpu_led_time_on.setStyleSheet(self.qle_unselected)

        self.lbl_dram_mon = QLabel(self)
        self.lbl_dram_mon.move(2, (self.monitor_btn_pos_h * 3))
        self.lbl_dram_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_dram_mon.setFont(self.font_s8b)
        self.lbl_dram_mon.setText(' DRAM')
        self.lbl_dram_mon.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_dram_mon)

        self.btn_dram_mon = QPushButton(self)
        self.btn_dram_mon.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 3))
        self.btn_dram_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_dram_mon.setFont(self.font_s8b)
        self.btn_dram_mon.setStyleSheet(self.btn_disabled_style)
        self.btn_dram_mon.clicked.connect(self.btn_dram_mon_function)
        print('-- created:', self.btn_dram_mon)

        self.btn_dram_mon_rgb_on = QLineEdit(self)
        self.btn_dram_mon_rgb_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_dram_mon_rgb_on.move((self.monitor_btn_w * 2) + 6, (self.monitor_btn_pos_h * 3))
        self.btn_dram_mon_rgb_on.setFont(self.font_s8b)
        self.btn_dram_mon_rgb_on.returnPressed.connect(self.btn_dram_mon_rgb_on_function)
        self.btn_dram_mon_rgb_on.setStyleSheet(self.qle_unselected)

        self.btn_dram_mon_rgb_off = QLineEdit(self)
        self.btn_dram_mon_rgb_off.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_dram_mon_rgb_off.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 3))
        self.btn_dram_mon_rgb_off.setFont(self.font_s8b)
        self.btn_dram_mon_rgb_off.returnPressed.connect(self.btn_dram_mon_rgb_off_function)
        self.btn_dram_mon_rgb_off.setStyleSheet(self.qle_unselected)

        self.btn_dram_led_time_on = QLineEdit(self)
        self.btn_dram_led_time_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_dram_led_time_on.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 3))
        self.btn_dram_led_time_on.setFont(self.font_s8b)
        self.btn_dram_led_time_on.returnPressed.connect(self.btn_dram_led_time_on_function)
        self.btn_dram_led_time_on.setStyleSheet(self.qle_unselected)

        self.lbl_vram_mon = QLabel(self)
        self.lbl_vram_mon.move(2, (self.monitor_btn_pos_h * 4))
        self.lbl_vram_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_vram_mon.setFont(self.font_s8b)
        self.lbl_vram_mon.setText(' VRAM')
        self.lbl_vram_mon.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_vram_mon)

        self.btn_vram_mon = QPushButton(self)
        self.btn_vram_mon.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 4))
        self.btn_vram_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_vram_mon.setFont(self.font_s8b)
        self.btn_vram_mon.setStyleSheet(self.btn_disabled_style)
        self.btn_vram_mon.clicked.connect(self.btn_vram_mon_function)
        print('-- created:', self.btn_vram_mon)

        self.btn_vram_mon_rgb_on = QLineEdit(self)
        self.btn_vram_mon_rgb_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_vram_mon_rgb_on.move((self.monitor_btn_w * 2) + 6, (self.monitor_btn_pos_h * 4))
        self.btn_vram_mon_rgb_on.setFont(self.font_s8b)
        self.btn_vram_mon_rgb_on.returnPressed.connect(self.btn_vram_mon_rgb_on_function)
        self.btn_vram_mon_rgb_on.setStyleSheet(self.qle_unselected)

        self.btn_vram_mon_rgb_off = QLineEdit(self)
        self.btn_vram_mon_rgb_off.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_vram_mon_rgb_off.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 4))
        self.btn_vram_mon_rgb_off.setFont(self.font_s8b)
        self.btn_vram_mon_rgb_off.returnPressed.connect(self.btn_vram_mon_rgb_off_function)
        self.btn_vram_mon_rgb_off.setStyleSheet(self.qle_unselected)

        self.btn_vram_led_time_on = QLineEdit(self)
        self.btn_vram_led_time_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_vram_led_time_on.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 4))
        self.btn_vram_led_time_on.setFont(self.font_s8b)
        self.btn_vram_led_time_on.returnPressed.connect(self.btn_vram_led_time_on_function)
        self.btn_vram_led_time_on.setStyleSheet(self.qle_unselected)

        self.lbl_hdd_mon = QLabel(self)
        self.lbl_hdd_mon.move(2, (self.monitor_btn_pos_h * 5))
        self.lbl_hdd_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_hdd_mon.setFont(self.font_s8b)
        self.lbl_hdd_mon.setText(' HDD')
        self.lbl_hdd_mon.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_hdd_mon)

        self.btn_hdd_mon = QPushButton(self)
        self.btn_hdd_mon.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 5))
        self.btn_hdd_mon.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_hdd_mon.setFont(self.font_s8b)
        self.btn_hdd_mon.setStyleSheet(self.btn_disabled_style)
        self.btn_hdd_mon.clicked.connect(self.btn_hdd_mon_function)
        print('-- created:', self.btn_hdd_mon)

        self.btn_hdd_mon_rgb_on = QLineEdit(self)
        self.btn_hdd_mon_rgb_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_hdd_mon_rgb_on.move((self.monitor_btn_w * 2) + 6, (self.monitor_btn_pos_h * 5))
        self.btn_hdd_mon_rgb_on.setFont(self.font_s8b)
        self.btn_hdd_mon_rgb_on.returnPressed.connect(self.btn_hdd_mon_rgb_on_function)
        self.btn_hdd_mon_rgb_on.setStyleSheet(self.qle_unselected)

        self.btn_hdd_mon_rgb_off = QLineEdit(self)
        self.btn_hdd_mon_rgb_off.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_hdd_mon_rgb_off.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 5))
        self.btn_hdd_mon_rgb_off.setFont(self.font_s8b)
        self.btn_hdd_mon_rgb_off.returnPressed.connect(self.btn_hdd_mon_rgb_off_function)
        self.btn_hdd_mon_rgb_off.setStyleSheet(self.qle_unselected)

        self.btn_hdd_led_time_on = QLineEdit(self)
        self.btn_hdd_led_time_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_hdd_led_time_on.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 5))
        self.btn_hdd_led_time_on.setFont(self.font_s8b)
        self.btn_hdd_led_time_on.returnPressed.connect(self.btn_hdd_led_time_on_function)
        self.btn_hdd_led_time_on.setStyleSheet(self.qle_unselected)

        self.lbl_exclusive_con = QLabel(self)
        self.lbl_exclusive_con.move(2, (self.monitor_btn_pos_h * 6))
        self.lbl_exclusive_con.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_exclusive_con.setFont(self.font_s8b)
        self.lbl_exclusive_con.setText(' EX_CON')
        self.lbl_exclusive_con.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_exclusive_con)

        self.btn_exclusive_con = QPushButton(self)
        self.btn_exclusive_con.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 6))
        self.btn_exclusive_con.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_exclusive_con.setFont(self.font_s8b)
        self.btn_exclusive_con.setStyleSheet(self.btn_disabled_style)
        self.btn_exclusive_con.clicked.connect(self.btn_exclusive_con_function)
        print('-- created:', self.btn_exclusive_con)

        self.lbl_start_minimized = QLabel(self)
        self.lbl_start_minimized.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 6))
        self.lbl_start_minimized.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_start_minimized.setFont(self.font_s8b)
        self.lbl_start_minimized.setText(' STRT_MIN')
        self.lbl_start_minimized.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_start_minimized)

        self.btn_start_minimized = QPushButton(self)
        self.btn_start_minimized.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 6))
        self.btn_start_minimized.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_start_minimized.setFont(self.font_s8b)
        self.btn_start_minimized.setStyleSheet(self.btn_disabled_style)
        self.btn_start_minimized.clicked.connect(self.btn_start_minimized_function)
        print('-- created:', self.btn_start_minimized)

        self.lbl_run_startup = QLabel(self)
        self.lbl_run_startup.move(2, (self.monitor_btn_pos_h * 7))
        self.lbl_run_startup.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.lbl_run_startup.setFont(self.font_s8b)
        self.lbl_run_startup.setText(' RUN_STRT')
        self.lbl_run_startup.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_run_startup)

        self.btn_run_startup = QPushButton(self)
        self.btn_run_startup.move(self.monitor_btn_w + 4, (self.monitor_btn_pos_h * 7))
        self.btn_run_startup.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_run_startup.setFont(self.font_s8b)
        self.btn_run_startup.setStyleSheet(self.btn_disabled_style)
        self.btn_run_startup.clicked.connect(self.btn_run_startup_function)
        print('-- created:', self.btn_run_startup)

        self.cpu_led_color_str = ""
        self.dram_led_color_str = ""
        self.vram_led_color_str = ""
        self.hdd_led_color_str = ""

        self.cpu_led_color_off_str = ""
        self.dram_led_color_off_str = ""
        self.vram_led_color_off_str = ""
        self.hdd_led_color_off_str = ""

        self.cpu_led_time_on_str = ""
        self.dram_led_time_on_str = ""
        self.vram_led_time_on_str = ""
        self.hdd_led_time_on_str = ""

        self.highlight_obj = []

        self.write_var = ''
        self.write_var_bool = False
        self.write_var_key = -1

        self.initUI()

    def btn_run_startup_function(self):
        global run_startup_bool
        self.setFocus()
        cwd = os.getcwd()
        shortcut_in = os.path.join(cwd + '\\iCUEDisplay.vbs')
        shortcut_out = os.path.join(os.path.expanduser('~') + '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/iCUEDisplay.lnk')

        # Remove shortcut
        if run_startup_bool is True:
            run_startup_bool = False
            print('-- setting run_startup_bool:', run_startup_bool)
            self.write_var = 'run_startup: false'
            self.write_changes()
            print('-- searching for:', shortcut_out)
            if os.path.exists(shortcut_out):
                print('-- removing:', shortcut_out)
                try:
                    os.remove(shortcut_out)
                except Exception as e:
                    print('-- btn_run_startup_function:', e)
            self.btn_run_startup.setText('DISABLED')
            self.btn_run_startup.setStyleSheet(self.btn_disabled_style)

        # Create shortcut
        elif run_startup_bool is False:
            if os.path.exists(shortcut_in):
                print('-- copying:', shortcut_in)
                try:
                    target = os.path.join(cwd + '\\iCUEDisplay.vbs')
                    icon = cwd + './icon.ico'
                    shell = win32com.client.Dispatch("WScript.Shell")
                    shortcut = shell.CreateShortCut(shortcut_out)
                    shortcut.Targetpath = target
                    shortcut.WorkingDirectory = cwd
                    shortcut.IconLocation = icon
                    shortcut.save()
                except Exception as e:
                    print('-- Error creating new shortcut file:', e)

            # Check new shortcut file exists
            if os.path.exists(shortcut_out):
                print('-- run at startup: file copied successfully')
                self.btn_run_startup.setText('ENABLED')
                self.btn_run_startup.setStyleSheet(self.btn_enabled_style)
                run_startup_bool = True
                print('-- setting run_startup_bool:', run_startup_bool)
                self.write_var = 'run_startup: true'
                self.write_changes()

            elif not os.path.exists(shortcut_out):
                print('-- run at startup: shortcut file failed to be created')
                self.btn_run_startup.setText('DISABLED')
                self.btn_run_startup.setStyleSheet(self.btn_disabled_style)

    def btn_start_minimized_function(self):
        self.setFocus()
        global start_minimized_bool
        if start_minimized_bool is True:
            start_minimized_bool = False
            print('-- setting start_minimized_bool:', start_minimized_bool)
            self.write_var = 'start_minimized: false'
            self.write_changes()
            self.btn_start_minimized.setText('DISABLED')
            self.btn_start_minimized.setStyleSheet(self.btn_disabled_style)

        elif start_minimized_bool is False:
            start_minimized_bool = True
            print('-- setting start_minimized_bool:', start_minimized_bool)
            self.write_var = 'start_minimized: true'
            self.write_changes()
            self.btn_start_minimized.setText('ENABLED')
            self.btn_start_minimized.setStyleSheet(self.btn_enabled_style)

    def sanitize_rgb_values(self):
        global cpu_led_color, dram_led_color, vram_led_color, hdd_led_color
        global cpu_led_color_off, dram_led_color_off, vram_led_color_off, hdd_led_color_off
        print('-- attempting to sanitize input:', self.btn_cpu_mon_rgb_on.text())
        var_str = self.write_var
        var_str = var_str.replace(' ', '')
        var_str = var_str.split(',')

        self.write_var_bool = False
        if len(var_str) == 3:
            if len(var_str[0]) >= 1 and len(var_str[0]) <= 3:
                if len(var_str[1]) >= 1 and len(var_str[1]) <= 3:
                    if len(var_str[2]) >= 1 and len(var_str[2]) <= 3:
                        if var_str[0].isdigit():
                            if var_str[1].isdigit():
                                if var_str[2].isdigit():
                                    var_int_0 = int(var_str[0])
                                    var_int_1 = int(var_str[1])
                                    var_int_2 = int(var_str[2])
                                    if var_int_0 >= 0 and var_int_0 <= 255:
                                        if var_int_1 >= 0 and var_int_1 <= 255:
                                            if var_int_2 >= 0 and var_int_2 <= 255:
                                                self.write_var_bool = True
                                                if self.write_var_key == 0:
                                                    cpu_led_color = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'cpu_led_color: ' + self.btn_cpu_mon_rgb_on.text().replace(' ', '')
                                                elif self.write_var_key == 1:
                                                    dram_led_color = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'dram_led_color: ' + self.btn_dram_mon_rgb_on.text().replace(' ', '')
                                                elif self.write_var_key == 2:
                                                    vram_led_color = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'vram_led_color: ' + self.btn_vram_mon_rgb_on.text().replace(' ', '')
                                                elif self.write_var_key == 3:
                                                    hdd_led_color = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'hdd_led_color: ' + self.btn_hdd_mon_rgb_on.text().replace(' ', '')
                                                elif self.write_var_key == 4:
                                                    cpu_led_color_off = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'cpu_led_color_off: ' + self.btn_cpu_mon_rgb_off.text().replace(' ', '')
                                                elif self.write_var_key == 5:
                                                    dram_led_color_off = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'dram_led_color_off: ' + self.btn_dram_mon_rgb_off.text().replace(' ', '')
                                                elif self.write_var_key == 6:
                                                    vram_led_color_off = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'vram_led_color_off: ' + self.btn_vram_mon_rgb_off.text().replace(' ', '')
                                                elif self.write_var_key == 7:
                                                    hdd_led_color_off = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'hdd_led_color_off: ' + self.btn_hdd_mon_rgb_off.text().replace(' ', '')

    def sanitize_interval(self):
        global cpu_led_time_on, dram_led_time_on, vram_led_time_on, hdd_led_time_on
        self.write_var_bool = False
        self.write_var = self.write_var.replace(' ', '')
        print('sanitize_interval:', self.write_var)
        try:
            self.write_var_float = float(float(self.write_var))
            print('write_var: is float')
            if float(self.write_var_float) >= 0.1 and float(self.write_var_float) <= 5:
                if self.write_var_key == 0:
                    cpu_led_time_on = self.write_var_float
                    self.write_var = 'cpu_led_time_on: ' + self.write_var
                    self.write_var_bool = True
                elif self.write_var_key == 1:
                    dram_led_time_on = self.write_var_float
                    self.write_var = 'dram_led_time_on: ' + self.write_var
                    self.write_var_bool = True
                elif self.write_var_key == 2:
                    vram_led_time_on = self.write_var_float
                    self.write_var = 'vram_led_time_on: ' + self.write_var
                    self.write_var_bool = True
                elif self.write_var_key == 3:
                    hdd_led_time_on = self.write_var_float
                    self.write_var = 'hdd_led_time_on: ' + self.write_var
                    self.write_var_bool = True
        except Exception as e:
            print('sanitize_interval:', e)

    def read_only_true(self):
        self.btn_cpu_mon_rgb_on.setReadOnly(True)
        self.btn_dram_mon_rgb_on.setReadOnly(True)
        self.btn_vram_mon_rgb_on.setReadOnly(True)
        self.btn_hdd_mon_rgb_on.setReadOnly(True)

        self.btn_cpu_mon_rgb_off.setReadOnly(True)
        self.btn_dram_mon_rgb_off.setReadOnly(True)
        self.btn_vram_mon_rgb_off.setReadOnly(True)
        self.btn_hdd_mon_rgb_off.setReadOnly(True)

        self.btn_cpu_led_time_on.setReadOnly(True)
        self.btn_dram_led_time_on.setReadOnly(True)
        self.btn_vram_led_time_on.setReadOnly(True)
        self.btn_hdd_led_time_on.setReadOnly(True)

    def read_only_false(self):
        self.btn_cpu_mon_rgb_on.setReadOnly(False)
        self.btn_dram_mon_rgb_on.setReadOnly(False)
        self.btn_vram_mon_rgb_on.setReadOnly(False)
        self.btn_hdd_mon_rgb_on.setReadOnly(False)

        self.btn_cpu_mon_rgb_off.setReadOnly(False)
        self.btn_dram_mon_rgb_off.setReadOnly(False)
        self.btn_vram_mon_rgb_off.setReadOnly(False)
        self.btn_hdd_mon_rgb_off.setReadOnly(False)

        self.btn_cpu_led_time_on.setReadOnly(False)
        self.btn_dram_led_time_on.setReadOnly(False)
        self.btn_vram_led_time_on.setReadOnly(False)
        self.btn_hdd_led_time_on.setReadOnly(False)

    def write_changes(self):
        self.read_only_true()
        write_var_split = self.write_var.split()
        write_var_split_key = write_var_split[0]
        self.write_var = self.write_var.strip()
        print(write_var_split)

        new_config_data = []
        print('-- writing changes to configuration file:', self.write_var)
        with open('./config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                print('0', line)
                if line.startswith(write_var_split_key):
                    print('1', line)
                    new_config_data.append(self.write_var)
                    print(self.write_var)
                else:
                    new_config_data.append(line)
        fo.close()

        open('./config.dat', 'w').close()

        with open('./config.dat', 'a') as fo:
            i = 0
            for new_config_datas in new_config_data:
                fo.writelines(new_config_data[i] + '\n')
                i += 1
        fo.close()

        self.read_only_false()

    def btn_cpu_led_time_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 0
        self.write_var = self.btn_cpu_led_time_on.text()
        self.sanitize_interval()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_cpu_led_time_on.text())
            self.write_changes()
            self.cpu_led_time_on_str = self.btn_cpu_led_time_on.text().replace(' ', '')
            self.btn_cpu_led_time_on.setText(self.cpu_led_time_on_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_cpu_led_time_on.text())
            self.cpu_led_time_on_str = str(cpu_led_time_on).replace(' ', '')
            self.btn_cpu_led_time_on.setText(self.cpu_led_time_on_str)
            conf_thread[0].start()

    def btn_dram_led_time_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 1
        self.write_var = self.btn_dram_led_time_on.text()
        self.sanitize_interval()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_dram_led_time_on.text())
            self.write_changes()
            self.dram_led_time_on_str = self.btn_dram_led_time_on.text().replace(' ', '')
            self.btn_dram_led_time_on.setText(self.dram_led_time_on_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_dram_led_time_on.text())
            self.dram_led_time_on_str = str(dram_led_time_on).replace(' ', '')
            self.btn_dram_led_time_on.setText(self.dram_led_time_on_str)
            conf_thread[0].start()

    def btn_vram_led_time_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 2
        self.write_var = self.btn_vram_led_time_on.text()
        self.sanitize_interval()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_vram_led_time_on.text())
            self.write_changes()
            self.vram_led_time_on_str = self.btn_vram_led_time_on.text().replace(' ', '')
            self.btn_vram_led_time_on.setText(self.vram_led_time_on_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_vram_led_time_on.text())
            self.vram_led_time_on_str = str(vram_led_time_on).replace(' ', '')
            self.btn_vram_led_time_on.setText(self.vram_led_time_on_str)
            conf_thread[0].start()

    def btn_hdd_led_time_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 3
        self.write_var = self.btn_hdd_led_time_on.text()
        self.sanitize_interval()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_hdd_led_time_on.text())
            self.write_changes()
            self.hdd_led_time_on_str = self.btn_hdd_led_time_on.text().replace(' ', '')
            self.btn_hdd_led_time_on.setText(self.hdd_led_time_on_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_hdd_led_time_on.text())
            self.hdd_led_time_on_str = str(hdd_led_time_on).replace(' ', '')
            self.btn_hdd_led_time_on.setText(self.hdd_led_time_on_str)
            conf_thread[0].start()

    def btn_cpu_mon_rgb_off_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 4
        self.write_var = self.btn_cpu_mon_rgb_off.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_cpu_mon_rgb_off.text())
            self.write_changes()
            self.cpu_led_color_off_str = self.btn_cpu_mon_rgb_off.text().replace(' ', '')
            self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace(',', ', ')
            self.btn_cpu_mon_rgb_off.setText(self.cpu_led_color_off_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_cpu_mon_rgb_off.text())
            self.cpu_led_color_off_str = str(cpu_led_color_off).replace('[', '')
            self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace(']', '')
            self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace(' ', '')
            self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace(',', ', ')
            self.btn_cpu_mon_rgb_off.setText(self.cpu_led_color_off_str)
            conf_thread[0].start()

    def btn_dram_mon_rgb_off_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 5
        self.write_var = self.btn_dram_mon_rgb_off.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_dram_mon_rgb_off.text())
            self.write_changes()
            self.dram_led_color_off_str = self.btn_dram_mon_rgb_off.text().replace(' ', '')
            self.dram_led_color_off_str = self.dram_led_color_off_str.replace(',', ', ')
            self.btn_dram_mon_rgb_off.setText(self.dram_led_color_off_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_dram_mon_rgb_off.text())
            self.dram_led_color_off_str = str(dram_led_color_off).replace('[', '')
            self.dram_led_color_off_str = self.dram_led_color_off_str.replace(']', '')
            self.dram_led_color_off_str = self.dram_led_color_off_str.replace(' ', '')
            self.dram_led_color_off_str = self.dram_led_color_off_str.replace(',', ', ')
            self.btn_dram_mon_rgb_off.setText(self.dram_led_color_off_str)
            conf_thread[0].start()

    def btn_vram_mon_rgb_off_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 6
        self.write_var = self.btn_vram_mon_rgb_off.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_vram_mon_rgb_off.text())
            self.write_changes()
            self.vram_led_color_off_str = self.btn_vram_mon_rgb_off.text().replace(' ', '')
            self.vram_led_color_off_str = self.vram_led_color_off_str.replace(',', ', ')
            self.btn_vram_mon_rgb_off.setText(self.vram_led_color_off_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_vram_mon_rgb_off.text())
            self.vram_led_color_off_str = str(vram_led_color_off).replace('[', '')
            self.vram_led_color_off_str = self.vram_led_color_off_str.replace(']', '')
            self.vram_led_color_off_str = self.vram_led_color_off_str.replace(' ', '')
            self.vram_led_color_off_str = self.vram_led_color_off_str.replace(',', ', ')
            self.btn_vram_mon_rgb_off.setText(self.vram_led_color_off_str)
            conf_thread[0].start()

    def btn_hdd_mon_rgb_off_function(self):
        global hdd_led_color, conf_thread
        self.setFocus()
        self.write_var_key = 7
        self.write_var = self.btn_hdd_mon_rgb_off.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_hdd_mon_rgb_off.text())
            self.write_changes()
            self.hdd_led_color_off_str = self.btn_hdd_mon_rgb_off.text().replace(' ', '')
            self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace(',', ', ')
            self.btn_hdd_mon_rgb_off.setText(self.hdd_led_color_off_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_hdd_mon_rgb_off.text())
            self.hdd_led_color_off_str = str(hdd_led_color_off).replace('[', '')
            self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace(']', '')
            self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace(' ', '')
            self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace(',', ', ')
            self.btn_hdd_mon_rgb_off.setText(self.hdd_led_color_off_str)
            conf_thread[0].start()

    def btn_cpu_mon_rgb_on_function(self):
        global cpu_led_color, conf_thread
        self.setFocus()
        self.write_var_key = 0
        self.write_var = self.btn_cpu_mon_rgb_on.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_cpu_mon_rgb_on.text())
            self.write_changes()
            self.cpu_led_color_str = self.btn_cpu_mon_rgb_on.text().replace(' ', '')
            self.cpu_led_color_str = self.cpu_led_color_str.replace(',', ', ')
            self.btn_cpu_mon_rgb_on.setText(self.cpu_led_color_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_cpu_mon_rgb_on.text())
            self.cpu_led_color_str = str(cpu_led_color).replace('[', '')
            self.cpu_led_color_str = self.cpu_led_color_str.replace(']', '')
            self.cpu_led_color_str = self.cpu_led_color_str.replace(' ', '')
            self.cpu_led_color_str = self.cpu_led_color_str.replace(',', ', ')
            self.btn_cpu_mon_rgb_on.setText(self.cpu_led_color_str)
            conf_thread[0].start()

    def btn_dram_mon_rgb_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 1
        self.write_var = self.btn_dram_mon_rgb_on.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_dram_mon_rgb_on.text())
            self.write_changes()
            self.dram_led_color_str = self.btn_dram_mon_rgb_on.text().replace(' ', '')
            self.dram_led_color_str = self.dram_led_color_str.replace(',', ', ')
            self.btn_dram_mon_rgb_on.setText(self.dram_led_color_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_dram_mon_rgb_on.text())
            self.dram_led_color_str = str(dram_led_color).replace('[', '')
            self.dram_led_color_str = self.dram_led_color_str.replace(']', '')
            self.dram_led_color_str = self.dram_led_color_str.text().replace(' ', '')
            self.dram_led_color_str = self.dram_led_color_str.replace(',', ', ')
            self.btn_dram_mon_rgb_on.setText(self.dram_led_color_str)
            conf_thread[0].start()

    def btn_vram_mon_rgb_on_function(self):
        global conf_thread
        self.setFocus()
        self.write_var_key = 2
        self.write_var = self.btn_vram_mon_rgb_on.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_vram_mon_rgb_on.text())
            self.write_changes()
            self.vram_led_color_str = self.btn_vram_mon_rgb_on.text().replace(' ', '')
            self.vram_led_color_str = self.vram_led_color_str.replace(',', ', ')
            self.btn_vram_mon_rgb_on.setText(self.vram_led_color_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_vram_mon_rgb_on.text())
            self.vram_led_color_str = str(vram_led_color).replace('[', '')
            self.vram_led_color_str = self.vram_led_color_str.replace(']', '')
            self.vram_led_color_str = self.vram_led_color_str.replace(' ', '')
            self.vram_led_color_str = self.vram_led_color_str.replace(',', ', ')
            self.btn_vram_mon_rgb_on.setText(self.vram_led_color_str)
            conf_thread[0].start()

    def btn_hdd_mon_rgb_on_function(self):
        global hdd_led_color, conf_thread
        self.setFocus()
        self.write_var_key = 3
        self.write_var = self.btn_hdd_mon_rgb_on.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_hdd_mon_rgb_on.text())
            self.write_changes()
            self.hdd_led_color_str = self.btn_hdd_mon_rgb_on.text().replace(' ', '')
            self.hdd_led_color_str = self.hdd_led_color_str.replace(',', ', ')
            self.btn_hdd_mon_rgb_on.setText(self.hdd_led_color_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_hdd_mon_rgb_on.text())
            self.hdd_led_color_str = str(hdd_led_color).replace('[', '')
            self.hdd_led_color_str = self.hdd_led_color_str.replace(']', '')
            self.hdd_led_color_str = self.hdd_led_color_str.replace(' ', '')
            self.hdd_led_color_str = self.hdd_led_color_str.replace(',', ', ')
            self.btn_hdd_mon_rgb_on.setText(self.hdd_led_color_str)
            conf_thread[0].start()

    def btn_cpu_mon_function(self):
        print('-- clicked: btn_cpu_mon')
        self.setFocus()
        global cpu_startup_bool, mon_threads
        if cpu_startup_bool is True:
            print('-- stopping: cpu monitor')
            self.btn_cpu_mon.setStyleSheet(self.btn_disabled_style)
            self.write_var = 'cpu_startup: false'
            mon_threads[1].stop()
            cpu_startup_bool = False
            self.btn_cpu_mon.setText('DISABLED')
        elif cpu_startup_bool is False:
            print('-- starting: cpu monitor')
            self.btn_cpu_mon.setStyleSheet(self.btn_enabled_style)
            self.write_var = 'cpu_startup: true'
            mon_threads[1].start()
            cpu_startup_bool = True
            self.btn_cpu_mon.setText('ENABLED')
        self.write_changes()

    def btn_dram_mon_function(self):
        print('-- clicked: btn_dram_mon')
        self.setFocus()
        global dram_startup_bool
        if dram_startup_bool is True:
            print('-- stopping: dram monitor')
            self.btn_dram_mon.setStyleSheet(self.btn_disabled_style)
            self.write_var = 'dram_startup: false'
            mon_threads[2].stop()
            dram_startup_bool = False
            self.btn_dram_mon.setText('DISABLED')
        elif dram_startup_bool is False:
            print('-- starting: dram monitor')
            self.btn_dram_mon.setStyleSheet(self.btn_enabled_style)
            self.write_var = 'dram_startup: true'
            mon_threads[2].start()
            dram_startup_bool = True
            self.btn_dram_mon.setText('ENABLED')
        self.write_changes()

    def btn_vram_mon_function(self):
        print('-- clicked: btn_vram_mon')
        self.setFocus()
        global vram_startup_bool
        if vram_startup_bool is True:
            print('-- stopping: vram monitor')
            self.btn_vram_mon.setStyleSheet(self.btn_disabled_style)
            self.write_var = 'vram_startup: false'
            mon_threads[3].stop()
            vram_startup_bool = False
            self.btn_vram_mon.setText('DISABLED')
        elif vram_startup_bool is False:
            print('-- starting: vram monitor')
            self.btn_vram_mon.setStyleSheet(self.btn_enabled_style)
            self.write_var = 'vram_startup: true'
            mon_threads[3].start()
            vram_startup_bool = True
            self.btn_vram_mon.setText('ENABLED')
        self.write_changes()

    def btn_hdd_mon_function(self):
        print('-- clicked: btn_hdd_mon')
        self.setFocus()
        global hdd_startup_bool
        if hdd_startup_bool is True:
            print('-- stopping: hdd monitor')
            self.btn_hdd_mon.setStyleSheet(self.btn_disabled_style)
            self.write_var = 'hdd_startup: false'
            mon_threads[0].stop()
            hdd_startup_bool = False
            self.btn_hdd_mon.setText('DISABLED')
        elif hdd_startup_bool is False:
            print('-- starting: hdd monitor')
            self.btn_hdd_mon.setStyleSheet(self.btn_enabled_style)
            self.write_var = 'hdd_startup: true'
            mon_threads[0].start()
            hdd_startup_bool = True
            self.btn_hdd_mon.setText('ENABLED')
        self.write_changes()

    def btn_exclusive_con_function(self):
        print('clicked: btn_exclusive_con_function')
        self.setFocus()
        global exclusive_access_bool
        if exclusive_access_bool is True:
            print('-- exclusive access request changed: requesting control')
            self.write_var = 'exclusive_access: true'
            sdk.request_control()
            exclusive_access_bool = False
            self.btn_exclusive_con.setStyleSheet(self.btn_enabled_style)
            self.btn_exclusive_con.setText('ENABLED')

        elif exclusive_access_bool is False:
            print('-- exclusive access request changed: releasing control')
            self.write_var = 'exclusive_access: false'
            sdk.release_control()
            exclusive_access_bool = True
            self.btn_exclusive_con.setStyleSheet(self.btn_disabled_style)
            self.btn_exclusive_con.setText('DISABLED')
        self.write_changes()

    def initUI(self):
        global mon_threads, conf_thread, allow_display_application, exclusive_access_bool
        global cpu_startup_bool, dram_startup_bool, vram_startup_bool, hdd_startup_bool
        global cpu_led_color, dram_led_color, vram_led_color, hdd_led_color
        global cpu_led_color_off, dram_led_color_off, vram_led_color_off, hdd_led_color_off
        global start_minimized_bool
        compile_devices_thread = CompileDevicesClass(self.lbl_con_stat)
        compile_devices_thread.start()
        read_configuration_thread = ReadConfigurationClass()
        read_configuration_thread.start()
        conf_thread.append(read_configuration_thread)
        hdd_mon_thread = HddMonClass()
        mon_threads.append(hdd_mon_thread)
        cpu_mon_thread = CpuMonClass()
        mon_threads.append(cpu_mon_thread)
        dram_mon_thread = DramMonClass()
        mon_threads.append(dram_mon_thread)
        vram_mon_thread = VramMonClass()
        mon_threads.append(vram_mon_thread)
        print('\n-- displaying application:')
        while allow_display_application is False:
            time.sleep(1)

        if run_startup_bool is True:
            self.btn_run_startup.setText('ENABLED')
            self.btn_run_startup.setStyleSheet(self.btn_enabled_style)
        elif run_startup_bool is False:
            self.btn_run_startup.setText('DISABLED')
            self.btn_run_startup.setStyleSheet(self.btn_disabled_style)

        if start_minimized_bool is True:
            self.showMinimized()
            self.btn_start_minimized.setText('ENABLED')
            self.btn_start_minimized.setStyleSheet(self.btn_enabled_style)
        elif start_minimized_bool is False:
            self.btn_start_minimized.setText('DISABLED')
            self.btn_start_minimized.setStyleSheet(self.btn_disabled_style)
        if cpu_startup_bool is True:
            self.btn_cpu_mon.setText('ENABLED')
            self.btn_cpu_mon.setStyleSheet(self.btn_enabled_style)
        elif cpu_startup_bool is False:
            self.btn_cpu_mon.setText('DISABLED')
            self.btn_cpu_mon.setStyleSheet(self.btn_disabled_style)
        if dram_startup_bool is True:
            self.btn_dram_mon.setText('ENABLED')
            self.btn_dram_mon.setStyleSheet(self.btn_enabled_style)
        elif dram_startup_bool is False:
            self.btn_dram_mon.setText('DISABLED')
            self.btn_dram_mon.setStyleSheet(self.btn_disabled_style)
        if vram_startup_bool is True:
            self.btn_vram_mon.setText('ENABLED')
            self.btn_vram_mon.setStyleSheet(self.btn_enabled_style)
        elif vram_startup_bool is False:
            self.btn_vram_mon.setText('DISABLED')
            self.btn_vram_mon.setStyleSheet(self.btn_disabled_style)
        if hdd_startup_bool is True:
            self.btn_hdd_mon.setText('ENABLED')
            self.btn_hdd_mon.setStyleSheet(self.btn_enabled_style)
        elif hdd_startup_bool is False:
            self.btn_hdd_mon.setText('DISABLED')
            self.btn_hdd_mon.setStyleSheet(self.btn_disabled_style)
        if exclusive_access_bool is True:
            self.btn_exclusive_con.setText('ENABLED')
            exclusive_access_bool = True
            self.btn_exclusive_con.setStyleSheet(self.btn_enabled_style)
        elif exclusive_access_bool is False:
            self.btn_exclusive_con.setText('DISABLED')
            exclusive_access_bool = False
            self.btn_exclusive_con.setStyleSheet(self.btn_disabled_style)
        if connected_bool is False:
            self.lbl_con_stat.setStyleSheet("""QLabel {background-color: rgb(255, 0, 0);
                                               color: rgb(0, 0, 0);
                                               border:2px solid rgb(35, 35, 35);}""")
        elif connected_bool is True:
            self.lbl_con_stat.setStyleSheet("""QLabel {background-color: rgb(0, 255, 0);
                                               color: rgb(0, 0, 0);
                                               border:2px solid rgb(35, 35, 35);}""")

        self.cpu_led_color_str = str(cpu_led_color).strip()
        self.cpu_led_color_str = self.cpu_led_color_str.replace('[', '')
        self.cpu_led_color_str = self.cpu_led_color_str.replace(']', '')
        self.btn_cpu_mon_rgb_on.setText(self.cpu_led_color_str)

        self.dram_led_color_str = str(dram_led_color).strip()
        self.dram_led_color_str = self.dram_led_color_str.replace('[', '')
        self.dram_led_color_str = self.dram_led_color_str.replace(']', '')
        self.btn_dram_mon_rgb_on.setText(self.dram_led_color_str)

        self.vram_led_color_str = str(vram_led_color).strip()
        self.vram_led_color_str = self.vram_led_color_str.replace('[', '')
        self.vram_led_color_str = self.vram_led_color_str.replace(']', '')
        self.btn_vram_mon_rgb_on.setText(self.vram_led_color_str)

        self.hdd_led_color_str = str(hdd_led_color).strip()
        self.hdd_led_color_str = self.hdd_led_color_str.replace('[', '')
        self.hdd_led_color_str = self.hdd_led_color_str.replace(']', '')
        self.btn_hdd_mon_rgb_on.setText(self.hdd_led_color_str)

        self.cpu_led_color_off_str = str(cpu_led_color_off).strip()
        self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace('[', '')
        self.cpu_led_color_off_str = self.cpu_led_color_off_str.replace(']', '')
        self.btn_cpu_mon_rgb_off.setText(self.cpu_led_color_off_str)

        self.dram_led_color_off_str = str(dram_led_color_off).strip()
        self.dram_led_color_off_str = self.dram_led_color_off_str.replace('[', '')
        self.dram_led_color_off_str = self.dram_led_color_off_str.replace(']', '')
        self.btn_dram_mon_rgb_off.setText(self.dram_led_color_off_str)

        self.vram_led_color_off_str = str(vram_led_color_off).strip()
        self.vram_led_color_off_str = self.vram_led_color_off_str.replace('[', '')
        self.vram_led_color_off_str = self.vram_led_color_off_str.replace(']', '')
        self.btn_vram_mon_rgb_off.setText(self.vram_led_color_off_str)

        self.hdd_led_color_off_str = str(hdd_led_color_off).strip()
        self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace('[', '')
        self.hdd_led_color_off_str = self.hdd_led_color_off_str.replace(']', '')
        self.btn_hdd_mon_rgb_off.setText(self.hdd_led_color_off_str)

        self.cpu_led_time_on_str = str(cpu_led_time_on).strip()
        self.btn_cpu_led_time_on.setText(self.cpu_led_time_on_str)

        self.dram_led_time_on_str = str(dram_led_time_on).strip()
        self.btn_dram_led_time_on.setText(self.dram_led_time_on_str)

        self.vram_led_time_on_str = str(vram_led_time_on).strip()
        self.btn_vram_led_time_on.setText(self.vram_led_time_on_str)

        self.hdd_led_time_on_str = str(hdd_led_time_on).strip()
        self.btn_hdd_led_time_on.setText(self.hdd_led_time_on_str)

        self.show()

    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                print('MINIMIZED')
                self.setFocus()
            else:
                print('DISPLAYED')
                self.setFocus()

    def mousePressEvent(self, event):
        self.prev_pos = event.globalPos()
        self.setFocus()

    def mouseMoveEvent(self, event):
        delta = QPoint(event.globalPos() - self.prev_pos)
        self.move(self.x() + delta.x(), self.y() + delta.y())
        self.prev_pos = event.globalPos()

    def pollCursor(self):
        pos = QCursor.pos()
        if pos != self.cursor:
            self.cursor = pos
            self.cursorMove.emit(pos)

    def handleCursorMove(self, pos):
        global out_of_bounds
        out_of_bounds = True
        if pos.x() > self.x():
            if pos.x() < (self.x() + self.width):
                if pos.y() < (self.y() + self.height):
                    if pos.y() > self.y():
                        if self.isMinimized() is False:
                            out_of_bounds = False


class CompileDevicesClass(QThread):
    def __init__(self, lbl_con_stat):
        QThread.__init__(self)
        self.lbl_con_stat = lbl_con_stat

    def run(self):
        global sdk, k95_rgb_platinum
        global connected_bool, connected_bool_prev
        global mon_threads, conf_thread
        global hdd_startup_bool, cpu_startup_bool, dram_startup_bool, vram_startup_bool, exclusive_access_bool
        global configuration_read_complete
        time.sleep(3)
        while True:
            try:
                if configuration_read_complete is False:
                    while configuration_read_complete is False:
                        print('-- waiting for configuration file to be read')
                        time.sleep(2)
                connected = sdk.connect()
                if not connected:
                    connected_bool = False
                    err = sdk.get_last_error()
                elif connected:
                    device = sdk.get_devices()
                    i = 0
                    for _ in device:
                        target_name = str(device[i])
                        if 'K95 RGB PLATINUM' in target_name:
                            k95_rgb_platinum.append(i)
                        i += 1
                    if len(k95_rgb_platinum) >= 1:
                        connected_bool = True
                if connected_bool is False and connected_bool != connected_bool_prev:
                    print('-- stopping threads:', )
                    conf_thread[0].stop()
                    mon_threads[0].stop()
                    mon_threads[1].stop()
                    mon_threads[2].stop()
                    mon_threads[3].stop()
                    if exclusive_access_bool is True:
                        sdk.request_control()
                        exclusive_access_bool = False
                    elif exclusive_access_bool is False:
                        sdk.release_control()
                        exclusive_access_bool = True
                    connected_bool_prev = False
                elif connected_bool is True and connected_bool != connected_bool_prev:
                    print('-- starting threads:', )
                    print('compile hdd_startup_bool:', hdd_startup_bool)
                    print('compile cpu_startup_bool:', cpu_startup_bool)
                    print('compile dram_startup_bool:', dram_startup_bool)
                    print('compile vram_startup_bool:', vram_startup_bool)
                    if hdd_startup_bool is True:
                        mon_threads[0].start()
                    if cpu_startup_bool is True:
                        mon_threads[1].start()
                    if dram_startup_bool is True:
                        mon_threads[2].start()
                    if vram_startup_bool is True:
                        mon_threads[3].start()
                    if exclusive_access_bool is True:
                        sdk.request_control()
                        sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], {1: (255, 0, 0)})
                        exclusive_access_bool = False
                    elif exclusive_access_bool is False:
                        sdk.release_control()
                        sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], {1: (255, 0, 0)})
                        exclusive_access_bool = True
                    connected_bool_prev = True

                if connected_bool is False:
                    self.lbl_con_stat.setStyleSheet("""QLabel {background-color: rgb(255, 0, 0);
                                                                       color: rgb(0, 0, 0);
                                                                       border:2px solid rgb(35, 35, 35);}""")
                elif connected_bool is True:
                    self.lbl_con_stat.setStyleSheet("""QLabel {background-color: rgb(0, 255, 0);
                                                                               color: rgb(0, 0, 0);
                                                                               border:2px solid rgb(35, 35, 35);}""")
            except Exception as e:
                print('[NAME]: CompileDevicesClass [FUNCTION]: run [EXCEPTION]:', e)
            time.sleep(3)


class ReadConfigurationClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.sanitize_passed = False
        self.sanitize_str = ''

    def sanitize_rgb_values(self):
        print('-- attempting to sanitize input:', self.sanitize_str)
        var_str = self.sanitize_str
        self.sanitize_passed = False
        if len(var_str) == 3:
            if len(var_str[0]) >= 1 and len(var_str[0]) <= 3:
                if len(var_str[1]) >= 1 and len(var_str[1]) <= 3:
                    if len(var_str[2]) >= 1 and len(var_str[2]) <= 3:
                        if var_str[0].isdigit():
                            if var_str[1].isdigit():
                                if var_str[2].isdigit():
                                    var_int_0 = int(var_str[0])
                                    var_int_1 = int(var_str[1])
                                    var_int_2 = int(var_str[2])
                                    if var_int_0 >= 0 and var_int_0 <= 255:
                                        if var_int_1 >= 0 and var_int_1 <= 255:
                                            if var_int_2 >= 0 and var_int_2 <= 255:
                                                self.sanitize_passed = True

    def startup_settings(self):
        global start_minimized_bool, run_startup_bool
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line == 'start_minimized: true':
                    start_minimized_bool = True
                    print('-- setting start_minimized_bool:', start_minimized_bool)
                elif line == 'start_minimized_bool: false':
                    start_minimized_bool = False
                    print('-- setting start_minimized_bool:', start_minimized_bool)

                if line == 'run_startup: true' and os.path.exists(os.path.join(os.path.expanduser('~') + '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/iCUEDisplay.lnk')):
                    run_startup_bool = True
                    print('-- setting run_startup:', run_startup_bool)
                elif line == 'run_startup: false' or not os.path.exists(os.path.join(os.path.expanduser('~') + '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/iCUEDisplay.lnk')):
                    run_startup_bool = False
                    print('-- setting run_startup:', run_startup_bool)

    def exclusive_access(self):
        global exclusive_access_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('exclusive_access: '):
                    if line == 'exclusive_access: true':
                        print('-- setting exclusive_access true')
                        exclusive_access_bool = True
                    elif line == 'exclusive_access: false':
                        print('-- setting exclusive_access false')
                        exclusive_access_bool = False

    def hdd_sanitize(self):
        global hdd_led_color, hdd_led_color_off, hdd_led_time_on, hdd_led_item, hdd_led_off_item
        global hdd_startup_bool
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()

                if line == 'hdd_startup: true':
                    hdd_startup_bool = True
                    print('hdd_startup_bool:', hdd_startup_bool)
                if line == 'hdd_startup: false':
                    hdd_startup_bool = False
                    print('hdd_startup_bool:', hdd_startup_bool)

                if line.startswith('hdd_led_color: '):
                    hdd_led_color = line.replace('hdd_led_color: ', '')
                    hdd_led_color = hdd_led_color.split(',')
                    self.sanitize_str = hdd_led_color
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        hdd_led_color[0] = int(hdd_led_color[0])
                        hdd_led_color[1] = int(hdd_led_color[1])
                        hdd_led_color[2] = int(hdd_led_color[2])
                        i = 0
                        hdd_led_item = []
                        for _ in alpha_led:
                            itm = {alpha_led[i]: hdd_led_color}
                            hdd_led_item.append(itm)
                            i += 1
                    elif self.sanitize_passed is False:
                        hdd_led_color = [255, 255, 255]

                if line.startswith('hdd_led_color_off: '):
                    hdd_led_color_off = line.replace('hdd_led_color_off: ', '')
                    hdd_led_color_off = hdd_led_color_off.split(',')
                    self.sanitize_str = hdd_led_color_off
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        hdd_led_color_off[0] = int(hdd_led_color_off[0])
                        hdd_led_color_off[1] = int(hdd_led_color_off[1])
                        hdd_led_color_off[2] = int(hdd_led_color_off[2])
                        i = 0
                        hdd_led_off_item = []
                        for _ in alpha_led:
                            itm = {alpha_led[i]: hdd_led_color_off}
                            hdd_led_off_item.append(itm)
                            i += 1
                    elif self.sanitize_passed is False:
                        hdd_led_color_off = [0, 0, 0]

                if line.startswith('hdd_led_time_on: '):
                    line = line.replace('hdd_led_time_on: ', '')
                    try:
                        line = float(float(line))
                        if line >= 0.1 and line <= 5:
                            hdd_led_time_on = line
                    except Exception as e:
                        hdd_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: hdd_led_time_on [EXCEPTION]:', e)

    def cpu_sanitize(self):
        global cpu_led_color, cpu_led_time_on, cpu_led_item, cpu_startup_bool
        global cpu_led_color_off, cpu_led_off_item
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()

                if line == 'cpu_startup: true':
                    cpu_startup_bool = True
                    print('cpu_startup_bool:', cpu_startup_bool)

                if line == 'cpu_startup: false':
                    cpu_startup_bool = False
                    print('cpu_startup_bool:', cpu_startup_bool)

                if line.startswith('cpu_led_color: '):
                    cpu_led_color = line.replace('cpu_led_color: ', '')
                    cpu_led_color = cpu_led_color.split(',')
                    self.sanitize_str = cpu_led_color
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        cpu_led_color[0] = int(cpu_led_color[0])
                        cpu_led_color[1] = int(cpu_led_color[1])
                        cpu_led_color[2] = int(cpu_led_color[2])
                        cpu_led_item = [
                            ({116: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 1
                            ({113: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 4
                            ({109: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 7
                            ({103: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])})]  # num
                    elif self.sanitize_passed is False:
                        cpu_led_color = [255, 255, 255]

                if line.startswith('cpu_led_color_off: '):
                    cpu_led_color_off = line.replace('cpu_led_color_off: ', '')
                    cpu_led_color_off = cpu_led_color_off.split(',')
                    self.sanitize_str = cpu_led_color_off
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        print('cpu_led_color_off: passed sanitization')
                        cpu_led_color_off[0] = int(cpu_led_color_off[0])
                        cpu_led_color_off[1] = int(cpu_led_color_off[1])
                        cpu_led_color_off[2] = int(cpu_led_color_off[2])
                        cpu_led_off_item = [
                            ({116: (cpu_led_color_off[0], cpu_led_color_off[1], cpu_led_color_off[2])}),  # 1
                            ({113: (cpu_led_color_off[0], cpu_led_color_off[1], cpu_led_color_off[2])}),  # 4
                            ({109: (cpu_led_color_off[0], cpu_led_color_off[1], cpu_led_color_off[2])}),  # 7
                            ({103: (cpu_led_color_off[0], cpu_led_color_off[1], cpu_led_color_off[2])})]  # num
                    elif self.sanitize_passed is False:
                        print('cpu_led_color_off: failed sanitization')
                        cpu_led_color_off = [0, 0, 0]

                if line.startswith('cpu_led_time_on: '):
                    line = line.replace('cpu_led_time_on: ', '')
                    try:
                        line = float(float(line))
                        if line >= 0.1 and line <= 5:
                            cpu_led_time_on = line
                    except Exception as e:
                        cpu_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: cpu_sanitize [EXCEPTION]:', e)

    def dram_sanitize(self):
        global dram_led_color, dram_led_time_on, dram_led_item, dram_startup_bool, dram_led_color_off, dram_led_off_item
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()

                if line == 'dram_startup: true':
                    dram_startup_bool = True
                    print('dram_startup_bool:', dram_startup_bool)
                if line == 'dram_startup: false':
                    dram_startup_bool = False
                    print('dram_startup_bool:', dram_startup_bool)

                if line.startswith('dram_led_color: '):
                    dram_led_color = line.replace('dram_led_color: ', '')
                    dram_led_color = dram_led_color.split(',')
                    self.sanitize_str = dram_led_color
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        print('dram_led_color: sanitize_passed')
                        dram_led_color[0] = int(dram_led_color[0])
                        dram_led_color[1] = int(dram_led_color[1])
                        dram_led_color[2] = int(dram_led_color[2])
                        dram_led_item = [({117: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 2
                                         ({114: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 5
                                         ({110: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 8
                                         ({104: (dram_led_color[0], dram_led_color[1], dram_led_color[2])})]  # /
                        print(dram_led_item)
                    elif self.sanitize_passed is False:
                        print('dram_led_color: sanitize_failed')
                        dram_led_color = [255, 255, 255]

                if line.startswith('dram_led_color_off: '):
                    dram_led_color_off = line.replace('dram_led_color_off: ', '')
                    dram_led_color_off = dram_led_color_off.split(',')
                    self.sanitize_str = dram_led_color_off
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        dram_led_color_off[0] = int(dram_led_color_off[0])
                        dram_led_color_off[1] = int(dram_led_color_off[1])
                        dram_led_color_off[2] = int(dram_led_color_off[2])
                        dram_led_off_item = [({117: (dram_led_color_off[0], dram_led_color_off[1], dram_led_color_off[2])}),  # 2
                                         ({114: (dram_led_color_off[0], dram_led_color_off[1], dram_led_color_off[2])}),  # 5
                                         ({110: (dram_led_color_off[0], dram_led_color_off[1], dram_led_color_off[2])}),  # 8
                                         ({104: (dram_led_color_off[0], dram_led_color_off[1], dram_led_color_off[2])})]  # /
                    elif self.sanitize_passed is False:
                        dram_led_color_off = [0, 0, 0]

                if line.startswith('dram_led_time_on: '):
                    line = line.replace('dram_led_time_on: ', '')
                    try:
                        line = float(float(line))
                        if line >= 0.1 and line <= 5:
                            dram_led_time_on = line
                    except Exception as e:
                        dram_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: dram_led_time_on [EXCEPTION]:', e)

    def vram_sanitize(self):
        global vram_led_color, vram_led_time_on, gpu_num, vram_led_item, vram_startup_bool, vram_led_off_item, vram_led_color_off
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()

                if line == 'vram_startup: true':
                    vram_startup_bool = True
                    print('vram_startup_bool:', vram_startup_bool)
                if line == 'vram_startup: false':
                    vram_startup_bool = False
                    print('vram_startup_bool:', vram_startup_bool)

                if line.startswith('vram_led_color: '):
                    vram_led_color = line.replace('vram_led_color: ', '')
                    vram_led_color = vram_led_color.split(',')
                    self.sanitize_str = vram_led_color
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        print('vram_led_color: sanitize_passed')
                        vram_led_color[0] = int(vram_led_color[0])
                        vram_led_color[1] = int(vram_led_color[1])
                        vram_led_color[2] = int(vram_led_color[2])
                        vram_led_item = [({118: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 3
                                         ({115: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 6
                                         ({111: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 9
                                         ({105: (vram_led_color[0], vram_led_color[1], vram_led_color[2])})]  # *
                        print(vram_led_item)
                    elif self.sanitize_passed is False:
                        print('vram_led_color: sanitize_failed')
                        vram_led_color = [255, 255, 255]

                if line.startswith('vram_led_color_off: '):
                    vram_led_color_off = line.replace('vram_led_color_off: ', '')
                    vram_led_color_off = vram_led_color_off.split(',')
                    self.sanitize_str = vram_led_color_off
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        vram_led_color_off[0] = int(vram_led_color_off[0])
                        vram_led_color_off[1] = int(vram_led_color_off[1])
                        vram_led_color_off[2] = int(vram_led_color_off[2])
                        vram_led_off_item = [({118: (vram_led_color_off[0], vram_led_color_off[1], vram_led_color_off[2])}),  # 3
                                         ({115: (vram_led_color_off[0], vram_led_color_off[1], vram_led_color_off[2])}),  # 6
                                         ({111: (vram_led_color_off[0], vram_led_color_off[1], vram_led_color_off[2])}),  # 9
                                         ({105: (vram_led_color_off[0], vram_led_color_off[1], vram_led_color_off[2])})]  # *
                    elif self.sanitize_passed is False:
                        vram_led_color_off = [0, 0, 0]

                elif line.startswith('vram_led_time_on: '):
                    vram_led_time_on_tmp = line.replace('vram_led_time_on: ', '')
                    try:
                        vram_led_time_on_tmp = float(float(vram_led_time_on_tmp))
                        if vram_led_time_on_tmp >= 0.1 and vram_led_time_on_tmp <= 5:
                            vram_led_time_on = vram_led_time_on_tmp
                    except Exception as e:
                        vram_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: dram_led_time_on [EXCEPTION]:', e)

                if line.startswith('gpu_num: '):
                    gpu_num_tmp = line.replace('gpu_num: ', '')
                    if gpu_num_tmp.isdigit():
                        gpu_num_tmp = int(gpu_num_tmp)
                        gpus = GPUtil.getGPUs()
                        if len(gpus) >= gpu_num_tmp:
                            gpu_num = gpu_num_tmp
                        else:
                            print('-- gpu_num: may exceed gpus currently active on the system. using default value')
                            gpu_num = 0
                    else:
                        gpu_num = 0

    def run(self):
        print('-- thread started: ReadConfigurationClass(QThread).run(self)')
        global allow_mon_threads_bool, exclusive_access_bool, allow_display_application, connected_bool
        global configuration_read_complete
        try:
            configuration_read_complete = False
            self.startup_settings()
            self.exclusive_access()
            self.hdd_sanitize()
            self.cpu_sanitize()
            self.dram_sanitize()
            self.vram_sanitize()
            allow_mon_threads_bool = True
            allow_display_application = True
            configuration_read_complete = True
        except Exception as e:
            print('[NAME]: ReadConfigurationClass [FUNCTION]: run [EXCEPTION]:', e)

    def stop(self):
        print('-- stopping: ReadConfigurationClass')
        self.terminate()


class HddMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        pythoncom.CoInitialize()
        global k95_rgb_platinum, allow_mon_threads_bool
        print('-- thread started: HddMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(hdd_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global hdd_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected, hdd_led_off_item
        global hdd_led_item
        self.get_stat()
        hdd_i = 0
        for _ in hdd_display_key_bool:
            if hdd_display_key_bool[hdd_i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], hdd_led_item[hdd_i])
            elif hdd_display_key_bool[hdd_i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], hdd_led_off_item[hdd_i])
            hdd_i += 1
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global hdd_display_key_bool, alpha_str, hdd_display_key_bool
        try:
            hdd_display_key_bool = []
            for _ in alpha_led:
                hdd_display_key_bool.append(False)
            obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            obj_swbem_services = obj_wmi_service.ConnectServer(".", "root\\cimv2")
            col_items = obj_swbem_services.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk")
            for objItem in col_items:
                if objItem.DiskBytesPersec != None:
                    if '_Total' not in objItem.Name:
                        var = objItem.Name.split()
                        try:
                            if len(var) >= 2:
                                disk_letter = var[1]
                                disk_letter = disk_letter.replace(':', '')
                                if len(disk_letter) == 1:
                                    if int(objItem.DiskBytesPersec) > 0:
                                        i = 0
                                        for _ in alpha_str:
                                            if canonical_caseless(disk_letter) == canonical_caseless(alpha_str[i]):
                                                hdd_display_key_bool[i] = True
                                            i += 1
                        except Exception as e:
                            print('[NAME]: HddMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
        except Exception as e:
            print('[NAME]: HddMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()

    def stop(self):
        print('-- stopping: HddMonClass')
        global sdk, k95_rgb_platinum, k95_rgb_platinum_selected, hdd_led_off_item
        hdd_i = 0
        for _ in hdd_led_off_item:
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], hdd_led_off_item[hdd_i])
            hdd_i += 1
        sdk.set_led_colors_flush_buffer()
        self.terminate()


class CpuMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.cpu_display_key_bool_tmp_0 = [True, False, False, False]
        self.cpu_display_key_bool_tmp_1 = [True, True, False, False]
        self.cpu_display_key_bool_tmp_2 = [True, True, True, False]
        self.cpu_display_key_bool_tmp_3 = [True, True, True, True]

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: CpuMonClass(QThread).run(self)')

        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(cpu_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global cpu_initiation, cpu_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        global cpu_led_item, cpu_led_off_item
        self.get_stat()
        cpu_i = 0
        for _ in cpu_led_item:
            if cpu_display_key_bool[cpu_i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], cpu_led_item[cpu_i])
            elif cpu_display_key_bool[cpu_i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], cpu_led_off_item[cpu_i])
            cpu_i += 1
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global cpu_stat, cpu_display_key_bool
        try:
            cpu_stat = psutil.cpu_percent(0.1)
            if cpu_stat < 25:
                cpu_display_key_bool = self.cpu_display_key_bool_tmp_0
            elif cpu_stat >= 25 and cpu_stat < 50:
                cpu_display_key_bool = self.cpu_display_key_bool_tmp_1
            elif cpu_stat >= 50 and cpu_stat < 75:
                cpu_display_key_bool = self.cpu_display_key_bool_tmp_2
            elif cpu_stat >= 75:
                cpu_display_key_bool = self.cpu_display_key_bool_tmp_3
        except Exception as e:
            print('[NAME]: CpuMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()

    def stop(self):
        print('-- stopping: CpuMonClass')
        global sdk, k95_rgb_platinum, k95_rgb_platinum_selected, cpu_led_off_item
        cpu_i = 0
        for _ in cpu_led_off_item:
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], cpu_led_off_item[cpu_i])
            cpu_i += 1
        sdk.set_led_colors_flush_buffer()
        self.terminate()


class DramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.dram_display_key_bool_tmp_0 = [True, False, False, False]
        self.dram_display_key_bool_tmp_1 = [True, True, False, False]
        self.dram_display_key_bool_tmp_2 = [True, True, True, False]
        self.dram_display_key_bool_tmp_3 = [True, True, True, True]

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: DramMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(dram_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global dram_initiation, dram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        dram_i = 0
        for _ in dram_led_item:
            if dram_display_key_bool[dram_i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], dram_led_item[dram_i])
            elif dram_display_key_bool[dram_i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], dram_led_off_item[dram_i])
            dram_i += 1
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global dram_stat, dram_display_key_bool
        try:
            dram_stat = psutil.virtual_memory().percent
            if dram_stat < 25:
                dram_display_key_bool = self.dram_display_key_bool_tmp_0
            elif dram_stat >= 25 and dram_stat < 50:
                dram_display_key_bool = self.dram_display_key_bool_tmp_1
            elif dram_stat >= 50 and dram_stat < 75:
                dram_display_key_bool = self.dram_display_key_bool_tmp_2
            elif dram_stat >= 75:
                dram_display_key_bool = self.dram_display_key_bool_tmp_3
        except Exception as e:
            print('[NAME]: DramMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()

    def stop(self):
        print('-- stopping: DramMonClass')
        global sdk, k95_rgb_platinum, k95_rgb_platinum_selected, dram_led_off_item
        dram_i = 0
        for _ in dram_led_off_item:
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], dram_led_off_item[dram_i])
            dram_i += 1
        sdk.set_led_colors_flush_buffer()
        self.terminate()


class VramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.vram_display_key_bool_tmp_0 = [True, False, False, False]
        self.vram_display_key_bool_tmp_1 = [True, True, False, False]
        self.vram_display_key_bool_tmp_2 = [True, True, True, False]
        self.vram_display_key_bool_tmp_3 = [True, True, True, True]

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: VramMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(vram_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global vram_initiation, vram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        vram_i = 0
        for _ in vram_led_item:
            if vram_display_key_bool[vram_i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], vram_led_item[vram_i])
            elif vram_display_key_bool[vram_i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], vram_led_off_item[vram_i])
            vram_i += 1
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global vram_stat, vram_display_key_bool, gpu_num
        try:
            gpus = GPUtil.getGPUs()
            if len(gpus) >= 0:
                vram_stat = float(f"{gpus[gpu_num].load * 100}")
                vram_stat = float(float(vram_stat))
                vram_stat = int(vram_stat)
                if vram_stat < 25:
                    vram_display_key_bool = self.vram_display_key_bool_tmp_0
                elif vram_stat >= 25 and vram_stat < 50:
                    vram_display_key_bool = self.vram_display_key_bool_tmp_1
                elif vram_stat >= 50 and vram_stat < 75:
                    vram_display_key_bool = self.vram_display_key_bool_tmp_2
                elif vram_stat >= 75:
                    vram_display_key_bool = self.vram_display_key_bool_tmp_3
        except Exception as e:
            print('[NAME]: VramMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()

    def stop(self):
        print('-- stopping: VramMonClass')
        global sdk, k95_rgb_platinum, k95_rgb_platinum_selected, vram_led_off_item
        vram_i = 0
        for _ in vram_led_off_item:
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], vram_led_off_item[vram_i])
            vram_i += 1
        sdk.set_led_colors_flush_buffer()
        self.terminate()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
