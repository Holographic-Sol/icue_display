"""
Written by Benjamin Jack Cullen aka Holographic_Sol
"""
import os
import sys
import time
import GPUtil
import psutil
import pythoncom
import unicodedata
import win32con
import win32api
import win32process
import win32com.client
import subprocess
from cuesdk import CueSdk
from PyQt5 import QtCore
from PyQt5.QtGui import QIcon, QCursor, QFont
from PyQt5.QtCore import Qt, QThread, QSize, QPoint, QCoreApplication, QTimer, QEvent
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QDesktopWidget, QLineEdit, QComboBox


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

sdk = CueSdk(os.path.join(os.getcwd(), 'bin\\CUESDK.x64_2017.dll'))
exclusive_access_bool = False
enum_compile_m_bool = True
enum_compile_kb_bool = True
install_bool = False
run_startup_bool = False
start_minimized_bool = False
allow_display_application = False
allow_mon_threads_bool = False
connected_bool = None
connected_bool_prev = None
configuration_read_complete = False
hdd_startup_bool = False
cpu_startup_bool = False
dram_startup_bool = False
vram_startup_bool = False
network_adapter_name_bool = False
network_adapter_startup_bool = False
ping_test_thread_startup = True
compile_devices_thread = []
mon_threads = []
conf_thread = []
key_board = []
key_board_selected = 0
mouse_device = []
mouse_device_selected = 0
prev_device = []
key_name = []
key_id = []
m_key_id = []
corsair_mouse_led_id = ['CorsairLedId.M_1', 'CorsairLedId.M_2', 'CorsairLedId.M_2', 'CorsairLedId.M_3', 'CorsairLedId.M_4']
cpu_led_id = ['CorsairLedId.K_Keypad1',
              'CorsairLedId.K_Keypad4',
              'CorsairLedId.K_Keypad7',
              'CorsairLedId.K_NumLock']
dram_led_id = ['CorsairLedId.K_Keypad2',
               'CorsairLedId.K_Keypad5',
               'CorsairLedId.K_Keypad8',
               'CorsairLedId.K_KeypadSlash']
vram_led_id = ['CorsairLedId.K_Keypad3',
               'CorsairLedId.K_Keypad6',
               'CorsairLedId.K_Keypad9',
               'CorsairLedId.K_KeypadAsterisk']
corsair_led_id_alpha = ['CorsairLedId.K_A', 'CorsairLedId.K_B', 'CorsairLedId.K_C', 'CorsairLedId.K_D',
                        'CorsairLedId.K_E', 'CorsairLedId.K_F', 'CorsairLedId.K_G', 'CorsairLedId.K_H',
                        'CorsairLedId.K_I', 'CorsairLedId.K_J', 'CorsairLedId.K_K', 'CorsairLedId.K_L',
                        'CorsairLedId.K_M', 'CorsairLedId.K_N', 'CorsairLedId.K_O', 'CorsairLedId.K_P',
                        'CorsairLedId.K_Q', 'CorsairLedId.K_R', 'CorsairLedId.K_S', 'CorsairLedId.K_T',
                        'CorsairLedId.K_U', 'CorsairLedId.K_V', 'CorsairLedId.K_W', 'CorsairLedId.K_X',
                        'CorsairLedId.K_Y', 'CorsairLedId.K_Z']
corsair_led_id_1_9 = ['CorsairLedId.K_1', 'CorsairLedId.K_2', 'CorsairLedId.K_3',
                      'CorsairLedId.K_4', 'CorsairLedId.K_5', 'CorsairLedId.K_6',
                      'CorsairLedId.K_7', 'CorsairLedId.K_8', 'CorsairLedId.K_9']
corsair_led_id_f1_f9 = ['CorsairLedId.K_F1', 'CorsairLedId.K_F2', 'CorsairLedId.K_F3',
                        'CorsairLedId.K_F4', 'CorsairLedId.K_F5', 'CorsairLedId.K_F6',
                        'CorsairLedId.K_F7', 'CorsairLedId.K_F8', 'CorsairLedId.K_F9']
corsair_led_id_f10 = 'CorsairLedId.K_F10'
corsair_led_id_f11 = 'CorsairLedId.K_F11'
corsair_led_id_f12 = 'CorsairLedId.K_F12'
corsair_led_id_0 = 'CorsairLedId.K_0'
alpha_str = ['a', 'b', 'c', 'd', 'e', 'f', 'g', 'h', 'i', 'j', 'k', 'l', 'm', 'n', 'o', 'p', 'q', 'r', 's', 't',
             'u', 'v', 'w', 'x', 'y', 'z']
mouse_led = ['', '', '', '']
ping_test_f11 = ()
ping_test_f12 = ()
net_adapter_connected_item = []
cpu_stat = ()
cpu_led_time_on = 1.0
cpu_led_color = [255, 255, 255]
cpu_led_color_off = [0, 0, 0]
cpu_led = [116, 113, 109, 103]
cpu_led_item = []
cpu_led_off_item = []
cpu_display_key_bool = [False, False, False, False]
dram_stat = ()
dram_led_time_on = 1.0
dram_led_color = [255, 255, 255]
dram_led_color_off = [0, 0, 0]
dram_led = [117, 114, 110, 104]
dram_led_item = []
dram_led_off_item = []
dram_display_key_bool = [False, False, False, False]
gpu_num = ()
vram_stat = ()
vram_led_time_on = 1.0
vram_led_color = [255, 255, 255]
vram_led_color_off = [0, 0, 0]
vram_led = [118, 115, 111, 105]
vram_led_item = []
vram_led_off_item = []
vram_display_key_bool = [False, False, False, False]
hdd_led_time_on = 0.5
hdd_led_color = [255, 255, 255]
hdd_led_color_off = [0, 0, 0]
hdd_led_item = []
hdd_led_off_item = []
hdd_display_key_bool = []
alpha_led_id = []
for _ in alpha_str:
    alpha_led_id.append('pending')
    hdd_display_key_bool.append(False)
network_adapter_name = ""
network_adapter_time_on = 0.5
net_rcv_led = [14, 15, 16, 17, 18, 19, 20, 21, 22]
net_snt_led = [2, 3, 4, 5, 6, 7, 8, 9, 10]
network_adapter_color_bytes = [255, 0, 0]
network_adapter_color_kb = [0, 255, 0]
network_adapter_color_mb = [0, 0, 255]
network_adapter_color_gb = [0, 255, 255]
network_adapter_color_tb = [255, 255, 255]
network_adapter_led_off_snt_item = []
network_adapter_led_off_rcv_item = []
network_adapter_led_rcv_item_bytes = []
network_adapter_led_snt_item_bytes = []
network_adapter_led_rcv_item_kb = []
network_adapter_led_snt_item_kb = []
network_adapter_led_rcv_item_mb = []
network_adapter_led_snt_item_mb = []
network_adapter_led_rcv_item_gb = []
network_adapter_led_snt_item_gb = []
network_adapter_led_rcv_item_tb = []
network_adapter_led_snt_item_tb = []
network_adapter_led_rcv_item_unit = []
network_adapter_led_snt_item_unit = []
network_adapter_led_rcv_item_unit_led = ()
network_adapter_led_snt_item_unit_led = ()
network_adapter_display_snt_bool = []
network_adapter_display_rcv_bool = []
i = 0
for _ in net_rcv_led:
    network_adapter_display_rcv_bool.append(False)
    network_adapter_display_snt_bool.append(False)
    i += 1


def create_new():
    global install_bool
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
            fo.writelines('run_startup: false\n')
            fo.writelines('network_adapter_color_off: 0,0,0\n')
            fo.writelines('network_adapter_time_on: 1\n')
            fo.writelines('network_adapter_startup: false\n')
            fo.writelines('network_adapter_name: ')
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
                install_bool = True


class App(QMainWindow):
    cursorMove = QtCore.pyqtSignal(object)

    def __init__(self):
        super(App, self).__init__()
        global install_bool

        create_new()
        if install_bool is True:
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
        self.height = 272
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
           border-top:1px solid rgb(40, 40, 40);
           border-left:1px solid rgb(40, 40, 40);}"""

        self.btn_enabled_style = """QPushButton{background-color: rgb(0, 0, 0);
                   color: rgb(200, 200, 200);
                   border-bottom:1px solid rgb(35, 35, 35);
                   border-right:1px solid rgb(35, 35, 35);
                   border-top:1px solid rgb(40, 40, 40);
                   border-left:1px solid rgb(40, 40, 40);}"""

        self.btn_disabled_style = """QPushButton{background-color: rgb(35, 35, 35);
                           color: rgb(200, 200, 200);
                           border-bottom:1px solid rgb(35, 35, 35);
                           border-right:1px solid rgb(35, 35, 35);
                           border-top:1px solid rgb(40, 40, 40);
                           border-left:1px solid rgb(40, 40, 40);}"""

        self.qle_selected = """QLineEdit{background-color: rgb(0, 0, 0);
               color: rgb(200, 200, 200);
               border-bottom:1px solid rgb(35, 35, 35);
               border-right:1px solid rgb(35, 35, 35);
               border-top:1px solid rgb(40, 40, 40);
               border-left:1px solid rgb(40, 40, 40);}"""

        self.qle_unselected = """QLineEdit{background-color: rgb(0, 0, 0);
                       color: rgb(200, 200, 200);
                       border-bottom:1px solid rgb(35, 35, 35);
                       border-right:1px solid rgb(35, 35, 35);
                       border-top:1px solid rgb(40, 40, 40);
                       border-left:1px solid rgb(40, 40, 40);}"""

        self.lbl_con_stat_false = """QLabel {background-color: rgb(255, 0, 0);
                                               color: rgb(0, 0, 0);
                                               border:2px solid rgb(35, 35, 35);}"""
        self.lbl_con_stat_true = """QLabel {background-color: rgb(0, 255, 0);
                                                       color: rgb(0, 0, 0);
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
        self.lbl_con_stat.move(2, 2)
        self.lbl_con_stat.resize(8, 8)
        self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_false)
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

        self.lbl_network_adapter = QLabel(self)
        self.lbl_network_adapter.move(2, (self.monitor_btn_pos_h * 6))
        self.lbl_network_adapter.resize((self.monitor_btn_w * 2) + 2, self.monitor_btn_h)
        self.lbl_network_adapter.setFont(self.font_s8b)
        self.lbl_network_adapter.setText(' NETWORK TRAFFIC')
        self.lbl_network_adapter.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_network_adapter)

        self.btn_network_adapter = QPushButton(self)
        self.btn_network_adapter.move((self.monitor_btn_w * 2) + 6, (self.monitor_btn_pos_h * 6))
        self.btn_network_adapter.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_network_adapter.setFont(self.font_s8b)
        self.btn_network_adapter.setStyleSheet(self.btn_disabled_style)
        self.btn_network_adapter.clicked.connect(self.btn_network_adapter_function)
        print('-- created:', self.btn_network_adapter)

        self.btn_network_adapter_rgb_off = QLineEdit(self)
        self.btn_network_adapter_rgb_off.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_network_adapter_rgb_off.move((self.monitor_btn_w * 3) + 8, (self.monitor_btn_pos_h * 6))
        self.btn_network_adapter_rgb_off.setFont(self.font_s8b)
        self.btn_network_adapter_rgb_off.returnPressed.connect(self.btn_network_adapter_mon_rgb_off_function)
        self.btn_network_adapter_rgb_off.setStyleSheet(self.qle_unselected)

        self.btn_network_adapter_led_time_on = QLineEdit(self)
        self.btn_network_adapter_led_time_on.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_network_adapter_led_time_on.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 6))
        self.btn_network_adapter_led_time_on.setFont(self.font_s8b)
        self.btn_network_adapter_led_time_on.returnPressed.connect(self.btn_network_adapter_led_time_on_function)
        self.btn_network_adapter_led_time_on.setStyleSheet(self.qle_unselected)

        self.cmb_network_adapter_name = QComboBox(self)
        self.cmb_network_adapter_name.resize(226, self.monitor_btn_h)
        self.cmb_network_adapter_name.move(self.monitor_btn_w + 2, (self.monitor_btn_pos_h * 7))
        self.cmb_network_adapter_name.setStyleSheet("""QComboBox {background-color: rgb(0, 0, 0);
                           color: rgb(200, 200, 200);
                           border:0px solid rgb(35, 35, 35);}""")
        self.cmb_network_adapter_name.activated[str].connect(self.cmb_network_adapter_name_function)

        self.btn_network_adapter_refresh = QPushButton(self)
        self.btn_network_adapter_refresh.move((self.monitor_btn_w * 4) + 10, (self.monitor_btn_pos_h * 7))
        self.btn_network_adapter_refresh.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_network_adapter_refresh.setFont(self.font_s8b)
        self.btn_network_adapter_refresh.setText('REFRESH')
        self.btn_network_adapter_refresh.setStyleSheet(self.btn_enabled_style)
        self.btn_network_adapter_refresh.clicked.connect(self.btn_network_adapter_refresh_function)
        print('-- created:', self.btn_network_adapter_refresh)

        self.lbl_exclusive_con = QLabel(self)
        self.lbl_exclusive_con.move(2, (self.height - ((self.monitor_btn_pos_h * 3) - 6)))
        self.lbl_exclusive_con.resize((self.monitor_btn_w * 2) + 2, self.monitor_btn_h)
        self.lbl_exclusive_con.setFont(self.font_s8b)
        self.lbl_exclusive_con.setText(' EXCLUSIVE CONTROL')
        self.lbl_exclusive_con.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_exclusive_con)

        self.btn_exclusive_con = QPushButton(self)
        self.btn_exclusive_con.move((self.monitor_btn_w * 2) + 6, (self.height - ((self.monitor_btn_pos_h * 3) - 6)))
        self.btn_exclusive_con.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_exclusive_con.setFont(self.font_s8b)
        self.btn_exclusive_con.setStyleSheet(self.btn_disabled_style)
        self.btn_exclusive_con.clicked.connect(self.btn_exclusive_con_function)
        print('-- created:', self.btn_exclusive_con)

        self.lbl_run_startup = QLabel(self)
        self.lbl_run_startup.move(2, (self.height - ((self.monitor_btn_pos_h * 2) - 4)))
        self.lbl_run_startup.resize((self.monitor_btn_w * 2) + 2, self.monitor_btn_h)
        self.lbl_run_startup.setFont(self.font_s8b)
        self.lbl_run_startup.setText(' AUTOMATIC STARTUP')
        self.lbl_run_startup.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_run_startup)

        self.btn_run_startup = QPushButton(self)
        self.btn_run_startup.move((self.monitor_btn_w * 2) + 6, (self.height - ((self.monitor_btn_pos_h * 2) - 4)))
        self.btn_run_startup.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_run_startup.setFont(self.font_s8b)
        self.btn_run_startup.setStyleSheet(self.btn_disabled_style)
        self.btn_run_startup.clicked.connect(self.btn_run_startup_function)
        print('-- created:', self.btn_run_startup)

        self.lbl_start_minimized = QLabel(self)
        self.lbl_start_minimized.move(2, (self.height - (self.monitor_btn_pos_h - 2)))
        self.lbl_start_minimized.resize((self.monitor_btn_w * 2 + 2), self.monitor_btn_h)
        self.lbl_start_minimized.setFont(self.font_s8b)
        self.lbl_start_minimized.setText(' START MINIMIZED')
        self.lbl_start_minimized.setStyleSheet(self.lbl_data_style)
        print('-- created:', self.lbl_start_minimized)

        self.btn_start_minimized = QPushButton(self)
        self.btn_start_minimized.move((self.monitor_btn_w * 2) + 6, (self.height - (self.monitor_btn_pos_h - 2)))
        self.btn_start_minimized.resize(self.monitor_btn_w, self.monitor_btn_h)
        self.btn_start_minimized.setFont(self.font_s8b)
        self.btn_start_minimized.setStyleSheet(self.btn_disabled_style)
        self.btn_start_minimized.clicked.connect(self.btn_start_minimized_function)
        print('-- created:', self.btn_start_minimized)

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

        self.network_adapter_led_color_off_str = ""
        self.network_adapter_led_time_on_str = ""

        self.write_var = ''
        self.write_var_bool = False
        self.write_var_key = -1

        self.initUI()

    def cmb_network_adapter_name_function(self, text):
        global network_adapter_name, conf_thread
        network_adapter_name = text
        print('-- setting network_adapter_name:', network_adapter_name)
        self.setFocus()
        self.write_var = 'network_adapter_name: ' + network_adapter_name
        self.write_changes()
        conf_thread[0].start()

    def btn_network_adapter_refresh_function(self, text):
        print('-- cmb_network_adapter_name_update_function:', text)
        pythoncom.CoInitialize()
        self.cmb_network_adapter_name.clear()
        try:
            obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            obj_swbem_services = obj_wmi_service.ConnectServer(".", "root\\cimv2")
            col_items = obj_swbem_services.ExecQuery('SELECT * FROM Win32_PerfFormattedData_Tcpip_NetworkAdapter')
            for objItem in col_items:
                if objItem.Name != None:
                    print('Found:', objItem.Name)
                    self.cmb_network_adapter_name.addItem(objItem.Name)
        except Exception as e:
            print('-- Exception in Function: network_adapter_sanitize.', e)

    def btn_network_adapter_mon_rgb_off_function(self):
        print('-- btn_network_adapter_mon_rgb_off_function')
        global conf_thread
        self.setFocus()
        self.write_var_key = 8
        self.write_var = self.btn_network_adapter_rgb_off.text()
        self.sanitize_rgb_values()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_network_adapter_rgb_off.text())
            self.write_changes()
            self.network_adapter_led_color_off_str = self.btn_network_adapter_rgb_off.text().replace(' ', '')
            self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace(',', ', ')
            self.btn_network_adapter_rgb_off.setText(self.network_adapter_led_color_off_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_network_adapter_rgb_off.text())
            self.network_adapter_led_color_off_str = str(network_adapter_color_off).replace('[', '')
            self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace(']', '')
            self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace(' ', '')
            self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace(',', ', ')
            self.btn_network_adapter_rgb_off.setText(self.network_adapter_led_color_off_str)
            conf_thread[0].start()
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

    def btn_network_adapter_led_time_on_function(self):
        print('-- btn_network_adapter_led_time_on_function')
        global conf_thread
        self.setFocus()
        self.write_var_key = 4
        self.write_var = self.btn_network_adapter_led_time_on.text()
        self.sanitize_interval()
        if self.write_var_bool is True:
            print('-- self.write_var passed sanitization checks:', self.btn_network_adapter_led_time_on.text())
            self.write_changes()
            self.network_adapter_led_time_on_str = self.btn_network_adapter_led_time_on.text().replace(' ', '')
            self.btn_network_adapter_led_time_on.setText(self.network_adapter_led_time_on_str)
            conf_thread[0].start()
        else:
            print('-- self.write_var failed sanitization checks:', self.btn_network_adapter_led_time_on.text())
            self.network_adapter_led_time_on_str = str(network_adapter_time_on).replace(' ', '')
            self.btn_network_adapter_led_time_on.setText(self.network_adapter_led_time_on_str)
            conf_thread[0].start()

    def btn_network_adapter_function(self):
        global network_adapter_startup_bool
        self.setFocus()
        if network_adapter_startup_bool is True:
            network_adapter_startup_bool = False
            print('-- setting network_adapter_startup_bool:', network_adapter_startup_bool)
            self.write_var = 'network_adapter_startup: false'
            mon_threads[4].stop()
            self.write_changes()
            self.btn_network_adapter.setText('DISABLED')
            self.btn_network_adapter.setStyleSheet(self.btn_disabled_style)
        elif network_adapter_startup_bool is False:
            network_adapter_startup_bool = True
            print('-- setting network_adapter_startup_bool:', network_adapter_startup_bool)
            self.write_var = 'network_adapter_startup: true'
            mon_threads[4].start()
            self.write_changes()
            self.btn_network_adapter.setText('ENABLED')
            self.btn_network_adapter.setStyleSheet(self.btn_enabled_style)

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
        global cpu_led_color_off, dram_led_color_off, vram_led_color_off, hdd_led_color_off, network_adapter_color_off
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
                                                elif self.write_var_key == 8:
                                                    network_adapter_color_off = [var_int_0, var_int_1, var_int_2]
                                                    self.write_var = 'network_adapter_color_off: ' + self.btn_network_adapter_rgb_off.text().replace(' ', '')

    def sanitize_interval(self):
        global cpu_led_time_on, dram_led_time_on, vram_led_time_on, hdd_led_time_on, network_adapter_time_on
        self.write_var_bool = False
        self.write_var = self.write_var.replace(' ', '')
        print('-- sanitize_interval:', self.write_var)
        try:
            self.write_var_float = float(float(self.write_var))
            print('-- write_var: is float', self.write_var_float)
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
                elif self.write_var_key == 4:
                    network_adapter_time_on = self.write_var_float
                    self.write_var = 'network_adapter_time_on: ' + self.write_var
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
            for _ in new_config_data:
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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global compile_devices_thread
        compile_devices_thread[0].stop()
        compile_devices_thread[0].start()

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
        global start_minimized_bool, compile_devices_thread
        compile_devices_thread_0 = CompileDevicesClass(self.lbl_con_stat, self.lbl_con_stat_false, self.lbl_con_stat_true)
        compile_devices_thread.append(compile_devices_thread_0)
        compile_devices_thread[0].start()
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
        network_mon_thread = NetworkMonClass()
        mon_threads.append(network_mon_thread)
        ping_test_thread = PingTestClass()
        mon_threads.append(ping_test_thread)
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
            self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_false)
        elif connected_bool is True:
            self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_true)
        if network_adapter_startup_bool is False:
            self.btn_network_adapter.setStyleSheet(self.btn_disabled_style)
            self.btn_network_adapter.setText('DISABLED')
        elif network_adapter_startup_bool is True:
            self.btn_network_adapter.setStyleSheet(self.btn_enabled_style)
            self.btn_network_adapter.setText('ENABLED')

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

        self.network_adapter_led_color_off_str = str(network_adapter_color_off).strip()
        self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace('[', '')
        self.network_adapter_led_color_off_str = self.network_adapter_led_color_off_str.replace(']', '')
        self.btn_network_adapter_rgb_off.setText(self.network_adapter_led_color_off_str)

        self.cpu_led_time_on_str = str(cpu_led_time_on).strip()
        self.btn_cpu_led_time_on.setText(self.cpu_led_time_on_str)

        self.dram_led_time_on_str = str(dram_led_time_on).strip()
        self.btn_dram_led_time_on.setText(self.dram_led_time_on_str)

        self.vram_led_time_on_str = str(vram_led_time_on).strip()
        self.btn_vram_led_time_on.setText(self.vram_led_time_on_str)

        self.hdd_led_time_on_str = str(hdd_led_time_on).strip()
        self.btn_hdd_led_time_on.setText(self.hdd_led_time_on_str)

        self.hdd_led_time_on_str = str(network_adapter_time_on).strip()
        self.btn_network_adapter_led_time_on.setText(self.hdd_led_time_on_str)

        self.cmb_network_adapter_name.addItem(network_adapter_name)

        self.show()

    def changeEvent(self, event):
        if event.type() == QEvent.WindowStateChange:
            if self.windowState() & Qt.WindowMinimized:
                self.setFocus()
            else:
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
        pass


class CompileDevicesClass(QThread):
    def __init__(self, lbl_con_stat, lbl_con_stat_false, lbl_con_stat_true):
        QThread.__init__(self)
        self.lbl_con_stat = lbl_con_stat
        self.lbl_con_stat_false = lbl_con_stat_false
        self.lbl_con_stat_true = lbl_con_stat_true
        self.device_str = ''
        self.device_index = ()
        self.compile_dicts_key = ()

    def compile_dicts(self):
        global alpha_led_id, hdd_led_item, hdd_led_off_item
        global hdd_led_color, hdd_led_color_off, cpu_led_color, cpu_led_color_off
        global dram_led_color, dram_led_color_off, vram_led_color, vram_led_color_off
        global network_adapter_color_off, network_adapter_color_bytes, network_adapter_color_kb
        global network_adapter_color_mb, network_adapter_color_gb, network_adapter_color_tb
        global net_rcv_led, net_snt_led, network_adapter_color_bytes, network_adapter_led_rcv_item_bytes, network_adapter_led_snt_item_bytes
        global network_adapter_color_kb, network_adapter_led_rcv_item_kb, network_adapter_led_snt_item_kb
        global network_adapter_color_mb, network_adapter_led_rcv_item_mb, network_adapter_led_snt_item_mb
        global network_adapter_color_gb, network_adapter_led_rcv_item_gb, network_adapter_led_snt_item_gb
        global network_adapter_color_tb, network_adapter_led_rcv_item_tb, network_adapter_led_snt_item_tb
        global network_adapter_color_off, network_adapter_led_off_rcv_item, network_adapter_led_off_snt_item
        global network_adapter_led_rcv_item_unit, network_adapter_led_snt_item_unit
        global cpu_led_item, dram_led_item, vram_led_item, cpu_led, dram_led, vram_led
        global cpu_led_off_item, dram_led_off_item, vram_led_off_item, net_adapter_connected_item

        if self.compile_dicts_key == 0:
            print('-- compiling list of dictionaries for keyboard:')
            hdd_led_item = []
            hdd_led_off_item = []
            cpu_led_item = []
            cpu_led_off_item = []
            dram_led_item = []
            dram_led_off_item = []
            vram_led_item = []
            vram_led_off_item = []

            i_1 = 0
            for _ in alpha_str:
                itm = {alpha_led_id[i_1]: hdd_led_color}
                hdd_led_item.append(itm)
                itm = {alpha_led_id[i_1]: hdd_led_color_off}
                hdd_led_off_item.append(itm)
                i_1 += 1
            print('Compiled hdd_led_item:', hdd_led_item)
            print('Compiled hdd_led_off_item:', hdd_led_off_item)

            i_2 = 0
            for _ in cpu_led:
                itm = {cpu_led[i_2]: cpu_led_color}
                cpu_led_item.append(itm)
                itm = {cpu_led[i_2]: cpu_led_color_off}
                cpu_led_off_item.append(itm)
                i_2 += 1
            print('Compiled cpu_led_item:', cpu_led_item, len(cpu_led_item))
            print('Compiled cpu_led_off_item:', cpu_led_off_item)

            i_3 = 0
            for _ in dram_led:
                itm = {dram_led[i_3]: dram_led_color}
                dram_led_item.append(itm)
                itm = {dram_led[i_3]: dram_led_color_off}
                dram_led_off_item.append(itm)
                i_3 += 1
            print('Compiled dram_led_item:', dram_led_item)
            print('Compiled dram_led_off_item:', dram_led_off_item)

            i_4 = 0
            for _ in vram_led:
                itm = {vram_led[i_4]: vram_led_color}
                vram_led_item.append(itm)
                itm = {vram_led[i_4]: vram_led_color_off}
                vram_led_off_item.append(itm)
                i_4 += 1
            print('Compiled vram_led_item:', vram_led_item)
            print('Compiled vram_led_off_item:', vram_led_off_item)

            network_adapter_led_rcv_item_bytes = []
            network_adapter_led_snt_item_bytes = []
            network_adapter_led_rcv_item_kb = []
            network_adapter_led_snt_item_kb = []
            network_adapter_led_rcv_item_mb = []
            network_adapter_led_snt_item_mb = []
            network_adapter_led_rcv_item_gb = []
            network_adapter_led_snt_item_gb = []
            network_adapter_led_rcv_item_tb = []
            network_adapter_led_snt_item_tb = []
            network_adapter_led_off_rcv_item = []
            network_adapter_led_off_snt_item = []

            i_5 = 0
            for _ in net_rcv_led:
                itm = {net_rcv_led[i_5]: network_adapter_color_bytes}
                network_adapter_led_rcv_item_bytes.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_bytes}
                network_adapter_led_snt_item_bytes.append(itm)

                itm = {net_rcv_led[i_5]: network_adapter_color_kb}
                network_adapter_led_rcv_item_kb.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_kb}
                network_adapter_led_snt_item_kb.append(itm)

                itm = {net_rcv_led[i_5]: network_adapter_color_mb}
                network_adapter_led_rcv_item_mb.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_mb}
                network_adapter_led_snt_item_mb.append(itm)

                itm = {net_rcv_led[i_5]: network_adapter_color_gb}
                network_adapter_led_rcv_item_gb.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_gb}
                network_adapter_led_snt_item_gb.append(itm)

                itm = {net_rcv_led[i_5]: network_adapter_color_tb}
                network_adapter_led_rcv_item_tb.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_tb}
                network_adapter_led_snt_item_tb.append(itm)

                itm = {net_rcv_led[i_5]: network_adapter_color_off}
                network_adapter_led_off_rcv_item.append(itm)
                itm = {net_snt_led[i_5]: network_adapter_color_off}
                network_adapter_led_off_snt_item.append(itm)
                i_5 += 1
            print('Compiled network_adapter_led_rcv_item_bytes:', network_adapter_led_rcv_item_bytes)
            print('Compiled network_adapter_led_snt_item_bytes:', network_adapter_led_snt_item_bytes)

            print('Compiled network_adapter_led_rcv_item_kb:', network_adapter_led_rcv_item_kb)
            print('Compiled network_adapter_led_snt_item_kb:', network_adapter_led_snt_item_kb)

            print('Compiled network_adapter_led_rcv_item_mb:', network_adapter_led_rcv_item_mb)
            print('Compiled network_adapter_led_snt_item_mb:', network_adapter_led_snt_item_mb)

            print('Compiled network_adapter_led_rcv_item_gb:', network_adapter_led_rcv_item_gb)
            print('Compiled network_adapter_led_snt_item_gb:', network_adapter_led_snt_item_gb)

            print('Compiled network_adapter_led_rcv_item_tb:', network_adapter_led_rcv_item_tb)
            print('Compiled network_adapter_led_snt_item_tb:', network_adapter_led_snt_item_tb)

            print('Compiled network_adapter_led_off_rcv_item:', network_adapter_led_off_rcv_item)
            print('Compiled network_adapter_led_off_snt_item:', network_adapter_led_off_snt_item)

            print('network_adapter_led_rcv_item_unit_led:', network_adapter_led_rcv_item_unit_led)
            print('network_adapter_led_snt_item_unit_led:', network_adapter_led_snt_item_unit_led)
            network_adapter_led_rcv_item_unit = [({network_adapter_led_rcv_item_unit_led: (255, 0, 0)}),
                                                 ({network_adapter_led_rcv_item_unit_led: (0, 0, 255)}),
                                                 ({network_adapter_led_rcv_item_unit_led: (0, 255, 255)}),
                                                 ({network_adapter_led_rcv_item_unit_led: (255, 255, 255)}),
                                                 ({network_adapter_led_rcv_item_unit_led: (0, 0, 0)})]
            network_adapter_led_snt_item_unit = [({network_adapter_led_snt_item_unit_led: (255, 0, 0)}),
                                                 ({network_adapter_led_snt_item_unit_led: (0, 0, 255)}),
                                                 ({network_adapter_led_snt_item_unit_led: (0, 255, 255)}),
                                                 ({network_adapter_led_snt_item_unit_led: (255, 255, 255)}),
                                                 ({network_adapter_led_snt_item_unit_led: (0, 0, 0)})]
            print('network_adapter_led_rcv_item_unit:', network_adapter_led_rcv_item_unit)
            print('network_adapter_led_snt_item_unit:', network_adapter_led_snt_item_unit)
        elif self.compile_dicts_key == 1:
            print('-- compiling list of dictionaries for mouse:')

    def enumerate_device(self):
        global key_board, mouse_device
        global key_name, key_id, m_key_id
        global corsair_led_id_alpha, alpha_led_id
        global corsair_led_id_1_9, corsair_led_id_f1_f9, net_rcv_led, net_snt_led, corsair_led_id_f10, corsair_led_id_0
        global network_adapter_led_rcv_item_unit_led, network_adapter_led_snt_item_unit_led
        global cpu_led_id, dram_led_id, vram_led_id, cpu_led, dram_led, vram_led
        global enum_compile_kb_bool, enum_compile_m_bool, ping_test_f11, ping_test_f12
        global corsair_led_id_f11, corsair_led_id_f12
        print('-- enumerating device:', self.device_str)

        # 1. Get Key Names & Key IDs
        led_position = sdk.get_led_positions_by_device_index(self.device_index)
        led_position_str = str(led_position).split('), ')
        led_position_str_tmp = led_position_str[0].split()
        print(led_position_str_tmp)

        if 'CorsairLedId.K_' in led_position_str_tmp[0]:
            if enum_compile_kb_bool is True:
                self.compile_dicts_key = 0
                enum_compile_kb_bool = False
                key_name = []
                key_id = []
                key_board.append(self.device_index)
                for _ in led_position_str:
                    var = _.split()
                    var_0 = var[0]
                    var_0 = var_0.replace('{', '')
                    var_0 = var_0.replace('<', '')
                    var_0 = var_0.replace(':', '')
                    var_1 = var[1].replace('>:', '')
                    key_name.append(var_0)
                    key_id.append(int(var_1))
                    i_led = 0
                    for _ in corsair_led_id_alpha:
                        if var_0 == corsair_led_id_alpha[i_led]:
                            alpha_led_id[i_led] = int(var_1)
                        i_led += 1
                    i_led = 0
                    for _ in corsair_led_id_1_9:
                        if var_0 == corsair_led_id_1_9[i_led]:
                            net_rcv_led[i_led] = int(var_1)
                        i_led += 1
                    i_led = 0
                    for _ in corsair_led_id_f1_f9:
                        if var_0 == corsair_led_id_f1_f9[i_led]:
                            net_snt_led[i_led] = int(var_1)
                        i_led += 1
                    if var_0 == corsair_led_id_f10:
                        network_adapter_led_snt_item_unit_led = int(var_1)
                    if var_0 == corsair_led_id_f11:
                        ping_test_f11 = int(var_1)
                    if var_0 == corsair_led_id_f12:
                        ping_test_f12 = int(var_1)
                    if var_0 == corsair_led_id_0:
                        network_adapter_led_rcv_item_unit_led = int(var_1)
                    i_led = 0
                    for _ in cpu_led_id:
                        if var_0 == cpu_led_id[i_led]:
                            cpu_led[i_led] = int(var_1)
                        i_led += 1
                    i_led = 0
                    for _ in dram_led_id:
                        if var_0 == dram_led_id[i_led]:
                            dram_led[i_led] = int(var_1)
                        i_led += 1
                    i_led = 0
                    for _ in vram_led_id:
                        if var_0 == vram_led_id[i_led]:
                            vram_led[i_led] = int(var_1)
                        i_led += 1
                self.compile_dicts()

        elif 'CorsairLedId.M_' in led_position_str_tmp[0]:
            self.compile_dicts_key = 1
            if enum_compile_m_bool is True:
                enum_compile_m_bool = False
                m_key_id = ['', '', '', '']
                mouse_device.append(int(self.device_index))
                for _ in led_position_str:
                    var = _.split()
                    var_0 = var[0]
                    var_0 = var_0.replace('{', '')
                    var_0 = var_0.replace('<', '')
                    var_0 = var_0.replace(':', '')
                    var_1 = var[1].replace('>:', '')
                    print('Corsair LED ID NAME:', var_0, '  LED ID:', var_1)
                    if 'CorsairLedId.M_1' == var_0:
                        m_key_id[0] = int(var_1)
                    elif 'CorsairLedId.M_2' == var_0:
                        m_key_id[1] = int(var_1)
                    elif 'CorsairLedId.M_3' == var_0:
                        m_key_id[2] = int(var_1)
                    elif 'CorsairLedId.M_4' == var_0:
                        m_key_id[3] = int(var_1)

    def stop_all_threads(self):
        global exclusive_access_bool, connected_bool_prev, conf_thread, mon_threads
        print('-- stopping all threads:', )
        conf_thread[0].stop()
        mon_threads[0].stop()
        mon_threads[1].stop()
        mon_threads[2].stop()
        mon_threads[3].stop()
        mon_threads[4].stop()
        if exclusive_access_bool is True:
            sdk.request_control()
            exclusive_access_bool = False
        elif exclusive_access_bool is False:
            sdk.release_control()
            exclusive_access_bool = True

    def stop_m_threads(self):
        print('-- stopping mouse threads')
        try:
            mon_threads[5].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[5]', e)

    def start_m_threads(self):
        global mon_threads
        if len(mouse_device) >= 1:
            print('-- starting mouse threads')
            if ping_test_thread_startup is True:
                mon_threads[5].start()

    def stop_kb_threads(self):
        global mon_threads
        print('-- stopping keyboard threads')
        try:
            mon_threads[0].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[0]', e)
        try:
            mon_threads[1].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[1]', e)
        try:
            mon_threads[2].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[2]', e)
        try:
            mon_threads[3].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[3]', e)
        try:
            mon_threads[4].stop()
        except Exception as e:
            print('-- exception stopping mon_threads[4]', e)
        print('-- stopping keyboard threads: done')

    def start_kb_threads(self):
        global hdd_startup_bool, cpu_startup_bool, dram_startup_bool, vram_startup_bool, network_adapter_startup_bool
        global mon_threads
        if len(key_board) >= 1:
            print('-- starting keyboard threads:', )
            if hdd_startup_bool is True:
                mon_threads[0].start()
            if cpu_startup_bool is True:
                mon_threads[1].start()
            if dram_startup_bool is True:
                mon_threads[2].start()
            if vram_startup_bool is True:
                mon_threads[3].start()
            if network_adapter_startup_bool is True:
                mon_threads[4].start()
        else:
            print('-- keyboard device not found: setting startup boolean values false')
            hdd_startup_bool = False
            cpu_startup_bool = False
            dram_startup_bool = False
            vram_startup_bool = False
            network_adapter_startup_bool = False

    def run(self):
        global sdk, key_board, mouse_device, mouse_device_selected
        global connected_bool, connected_bool_prev
        global mon_threads, conf_thread
        global hdd_startup_bool, cpu_startup_bool, dram_startup_bool, vram_startup_bool, exclusive_access_bool
        global configuration_read_complete, network_adapter_startup_bool
        global key_name, key_id
        global enum_compile_kb_bool, enum_compile_m_bool, prev_device, prev_kb
        time.sleep(3)
        key_name = []
        key_id = []
        prev_device = []
        while True:
            try:
                if configuration_read_complete is False:
                    print('-- waiting for configuration file to be read')
                    while configuration_read_complete is False:
                        time.sleep(3)
                connected = sdk.connect()

                # Not Connected
                if not connected:
                    connected_bool = False
                    prev_device = []
                    err = sdk.get_last_error()
                    if connected_bool != connected_bool_prev:
                        self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_false)
                        self.stop_all_threads()
                        connected_bool_prev = False
                    print('-- waiting for icue service to be started')
                    while not connected:
                        try:
                            connected = sdk.connect()
                        except Exception as e:
                            print('-- CompileDevicesClass:', e)
                        time.sleep(3)
                    connected_bool_prev = False
                # Connected
                if connected:
                    connected_bool = True
                    device = sdk.get_devices()
                    if str(device) != str(prev_device):
                        enum_compile_kb_bool = True
                        enum_compile_m_bool = True
                        print('-- device list changed: enumerating changes')
                        device_i = 0
                        key_board = []
                        mouse_device = []
                        for _ in device:
                            target_name = str(device[device_i])
                            self.device_index = device_i
                            self.device_str = target_name
                            self.enumerate_device()
                            device_i += 1

                        if len(key_board) < 1:
                            print('-- keyboard unplugged:', key_board)
                            self.stop_kb_threads()
                        elif len(key_board) >= 1:
                            print('-- found keyboard:', sdk.get_device_info(key_board[0]))
                            if str(sdk.get_device_info(key_board[0])) not in str(prev_device):
                                self.start_kb_threads()

                        if len(mouse_device) < 1:
                            print('-- mouse unplugged:', mouse_device)
                            self.stop_m_threads()
                        elif len(mouse_device) >= 1:
                            print('-- found mouse:', sdk.get_device_info(mouse_device[0]))
                            if str(sdk.get_device_info(mouse_device[0])) not in str(prev_device):
                                self.start_m_threads()

                        prev_device = device

                if connected_bool is False and connected_bool != connected_bool_prev:
                    self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_false)
                    self.stop_all_threads()
                    connected_bool_prev = False

                elif connected_bool is True and connected_bool != connected_bool_prev:
                    if exclusive_access_bool is True:
                        sdk.request_control()
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], {1: (255, 0, 0)})
                        exclusive_access_bool = False
                    elif exclusive_access_bool is False:
                        sdk.release_control()
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], {1: (255, 0, 0)})
                        exclusive_access_bool = True
                    connected_bool_prev = True

                if connected_bool is False:
                    self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_false)
                elif connected_bool is True:
                    self.lbl_con_stat.setStyleSheet(self.lbl_con_stat_true)
            except Exception as e:
                print('[NAME]: CompileDevicesClass [FUNCTION]: run [EXCEPTION]:', e)
            time.sleep(3)

    def stop(self):
        print('-- stopping: CompileDevicesClass')
        self.stop_kb_threads()
        self.stop_m_threads()
        self.terminate()


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

    def network_adapter_sanitize(self):
        pythoncom.CoInitialize()
        global network_adapter_startup_bool, network_adapter_time_on
        global network_adapter_color_off
        global net_snt_led, net_rcv_led, network_adapter_led_off_rcv_item, network_adapter_led_off_snt_item
        global network_adapter_name, network_adapter_name_bool, network_adapter_led_off_snt_item
        global network_adapter_led_off_rcv_item
        network_adapter_name_bool = False
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('network_adapter_name: '):
                    var = line.replace('network_adapter_name: ', '')
                    print('-- checking network_adapter_name:', var)
                    try:
                        obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
                        obj_swbem_services = obj_wmi_service.ConnectServer(".", "root\\cimv2")
                        col_items = obj_swbem_services.ExecQuery('SELECT * FROM Win32_PerfFormattedData_Tcpip_NetworkAdapter')
                        for objItem in col_items:
                            if objItem.Name != None:
                                if var == objItem.Name:
                                    print('Found:', objItem.Name)
                                    var = objItem.Name
                                    network_adapter_name_bool = True
                    except Exception as e:
                        print('-- Exception in Function: network_adapter_sanitize.', e)
                    if network_adapter_name_bool is True:
                        network_adapter_name = var
                    elif network_adapter_name_bool is False:
                        network_adapter_name = ""
                    print('-- setting network_adapter_name:', network_adapter_name)
                if line == 'network_adapter_startup: true':
                    network_adapter_startup_bool = True
                    print('network_adapter_startup:', network_adapter_startup_bool)
                if line == 'network_adapter_startup: false':
                    network_adapter_startup_bool = False
                    print('network_adapter_startup:', network_adapter_startup_bool)
                if line.startswith('network_adapter_color_off: '):
                    network_adapter_color_off = line.replace('network_adapter_color_off: ', '')
                    network_adapter_color_off = network_adapter_color_off.split(',')
                    self.sanitize_str = network_adapter_color_off
                    self.sanitize_rgb_values()
                    if self.sanitize_passed is True:
                        network_adapter_color_off[0] = int(network_adapter_color_off[0])
                        network_adapter_color_off[1] = int(network_adapter_color_off[1])
                        network_adapter_color_off[2] = int(network_adapter_color_off[2])
                    elif self.sanitize_passed is False:
                        network_adapter_color_off = [0, 0, 0]
                if line.startswith('network_adapter_time_on: '):
                    line = line.replace('network_adapter_time_on: ', '')
                    try:
                        line = float(float(line))
                        if line >= 0.1 and line <= 5:
                            network_adapter_time_on = line
                    except Exception as e:
                        network_adapter_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: network_adapter_time_on [EXCEPTION]:', e)
        print('network_adapter_color_off:', network_adapter_color_off)
        print('network_adapter_time_on', network_adapter_time_on)
        print('network_adapter_startup:', network_adapter_startup_bool)

    def startup_settings(self):
        global start_minimized_bool, run_startup_bool
        startup_loc = '/AppData/Roaming/Microsoft/Windows/Start Menu/Programs/Startup/iCUEDisplay.lnk'
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line == 'start_minimized: true':
                    start_minimized_bool = True
                    print('-- setting start_minimized_bool:', start_minimized_bool)
                elif line == 'start_minimized_bool: false':
                    start_minimized_bool = False
                    print('-- setting start_minimized_bool:', start_minimized_bool)
                if line == 'run_startup: true' and os.path.exists(os.path.join(os.path.expanduser('~') + startup_loc)):
                    run_startup_bool = True
                    print('-- setting run_startup:', run_startup_bool)
                elif line == 'run_startup: false' or not os.path.exists(os.path.join(os.path.expanduser('~') + startup_loc)):
                    run_startup_bool = False
                    print('-- setting run_startup:', run_startup_bool)

    def exclusive_access(self):
        global exclusive_access_bool, sdk, key_board, key_board_selected
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
            self.network_adapter_sanitize()
            allow_mon_threads_bool = True
            allow_display_application = True
            configuration_read_complete = True
        except Exception as e:
            print('[NAME]: ReadConfigurationClass [FUNCTION]: run [EXCEPTION]:', e)

    def stop(self):
        print('-- stopping: ReadConfigurationClass')
        self.terminate()


class NetworkMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.b_type = ()
        self.u_type = ()
        self.b_type_1 = ()
        self.u_type_1 = ()
        self.num_len_key = ()
        self.switch_num = ()
        self.switch_num_key = ()
        self.switch_num_1 = ()
        self.b_type_key = ()

    def run(self):
        pythoncom.CoInitialize()
        global key_board, allow_mon_threads_bool, network_adapter_time_on, network_adapter_display_rcv_bool
        global network_adapter_display_snt_bool
        print('-- thread started: NetworkMonClass(QThread).run(self)')
        while True:
            if len(key_board) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(network_adapter_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global network_adapter_display_rcv_bool, sdk, key_board, key_board_selected
        global network_adapter_display_snt_bool, network_adapter_led_off_snt_item
        global network_adapter_led_rcv_item_bytes, network_adapter_led_rcv_item_kb, network_adapter_led_rcv_item_mb
        global network_adapter_led_rcv_item_gb, network_adapter_led_rcv_item_tb
        global network_adapter_led_rcv_item_unit, network_adapter_led_snt_item_unit
        global network_adapter_led_snt_item_bytes, network_adapter_led_snt_item_kb, network_adapter_led_snt_item_mb
        global network_adapter_led_snt_item_gb, network_adapter_led_snt_item_tb
        self.get_stat()
        net_rcv_i = 0
        try:
            for _ in network_adapter_display_rcv_bool:
                if network_adapter_display_rcv_bool[net_rcv_i] is True:
                    if self.u_type == 0:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_unit[0])
                    elif self.u_type == 1:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_unit[1])
                    elif self.u_type == 2:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_unit[2])
                    elif self.u_type == 3:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_unit[3])
                    if self.b_type == 0:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_bytes[net_rcv_i])
                    elif self.b_type == 1:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_kb[net_rcv_i])
                    elif self.b_type == 2:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_mb[net_rcv_i])
                    elif self.b_type == 3:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_gb[net_rcv_i])
                    elif self.b_type == 4:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_tb[net_rcv_i])
                elif network_adapter_display_rcv_bool[net_rcv_i] is False:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_off_rcv_item[net_rcv_i])
                if network_adapter_display_snt_bool[net_rcv_i] is True:
                    if self.u_type_1 == 0:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[0])
                    elif self.u_type_1 == 1:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[1])
                    elif self.u_type_1 == 2:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[2])
                    elif self.u_type_1 == 3:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[3])
                    if self.b_type_1 == 0:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_bytes[net_rcv_i])
                    elif self.b_type_1 == 1:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_kb[net_rcv_i])
                    elif self.b_type_1 == 2:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_mb[net_rcv_i])
                    elif self.b_type_1 == 3:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_gb[net_rcv_i])
                    elif self.b_type_1 == 4:
                        sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_tb[net_rcv_i])
                elif network_adapter_display_snt_bool[net_rcv_i] is False:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_off_snt_item[net_rcv_i])
                if network_adapter_display_rcv_bool == [False, False, False, False, False, False, False, False, False]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_rcv_item_unit[4])
                if network_adapter_display_snt_bool == [False, False, False, False, False, False, False, False, False]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[4])
                net_rcv_i += 1
        except Exception as e:
            print('NetworkMonClass:', e)
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global network_adapter_display_rcv_bool, net_rcv_led, network_adapter_name, network_adapter_time_on
        global network_adapter_display_snt_bool
        network_adapter_exists_bool = False
        try:
            network_adapter_display_rcv_bool = [False, False, False, False, False, False, False, False, False]
            network_adapter_display_snt_bool = [False, False, False, False, False, False, False, False, False]
            rec_item = ''
            sen_item = ''
            obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            obj_swbem_services = obj_wmi_service.ConnectServer(".", "root\\cimv2")
            col_items = obj_swbem_services.ExecQuery('SELECT * FROM Win32_PerfFormattedData_Tcpip_NetworkAdapter')
            for objItem in col_items:
                if objItem.Name != None:
                    if network_adapter_name == objItem.Name:
                        rec_item = objItem.BytesReceivedPersec
                        sen_item = objItem.BytesSentPersec
                        network_adapter_exists_bool = True
            if network_adapter_exists_bool is True:
                self.b_type_key = 0
                rec_bytes = self.convert_bytes(float(rec_item))
                self.b_type_key = 1
                sen_bytes = self.convert_bytes(float(sen_item))
                rec_bytes_int = int(rec_bytes)
                sen_bytes_int = int(sen_bytes)
                self.num_len_key = 0
                self.num_len(rec_bytes_int)
                self.num_len_key = 1
                self.num_len(sen_bytes_int)
                self.switch_num_key = 0
                self.switch_num_function(rec_bytes_int)
                self.switch_num_key = 1
                self.switch_num_function(sen_bytes_int)

                # print('rec_bytes_int:', rec_bytes_int, 'u_type:', self.u_type)
                # print('sen_bytes_int:', sen_bytes_int, 'u_type_1:', self.u_type_1)

        except Exception as e:
            print('[NAME]: NetworkMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()

    def convert_bytes(self, num):
        i = 0
        for x in ['bytes', 'KB', 'MB', 'GB', 'TB']:
            if num < 1024.0:
                if self.b_type_key == 0:
                    self.b_type = i
                elif self.b_type_key == 1:
                    self.b_type_1 = i
                return num
            num /= 1024.0
            i += 1

    def switch_num_function(self, num):
        global network_adapter_display_rcv_bool, network_adapter_display_snt_bool
        n = str(num)
        n = n[0]
        for x in ['1', '2', '3', '4', '5', '6', '7', '8', '9']:
            if n == x:
                if self.switch_num_key == 0:
                    self.switch_num = int(n)
                    network_adapter_display_rcv_bool = [False, False, False, False, False, False, False, False, False]
                    i_rcv = 0
                    for _ in network_adapter_display_rcv_bool:
                        if i_rcv < int(self.switch_num):
                            network_adapter_display_rcv_bool[i_rcv] = True
                        i_rcv += 1
                elif self.switch_num_key == 1:
                    self.switch_num_1 = int(n)
                    network_adapter_display_snt_bool = [False, False, False, False, False, False, False, False, False]
                    i_snt = 0
                    for _ in network_adapter_display_snt_bool:
                        if i_snt < int(self.switch_num_1):
                            network_adapter_display_snt_bool[i_snt] = True
                        i_snt += 1

    def num_len(self, num):
        n = len(str(num))
        if n == 1:
            if self.num_len_key == 0:
                self.u_type = 0
            elif self.num_len_key == 1:
                self.u_type_1 = 0
        elif n == 2:
            if self.num_len_key == 0:
                self.u_type = 1
            elif self.num_len_key == 1:
                self.u_type_1 = 1
        elif n == 3:
            if self.num_len_key == 0:
                self.u_type = 2
            elif self.num_len_key == 1:
                self.u_type_1 = 2
        elif n >= 4:
            if self.num_len_key == 0:
                self.u_type = 3
            elif self.num_len_key == 1:
                self.u_type_1 = 3

    def stop(self):
        print('-- stopping: NetworkMonClass')
        global sdk, key_board, key_board_selected, network_adapter_led_off_rcv_item
        try:
            net_rcv_i = 0
            for _ in network_adapter_led_off_rcv_item:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_off_rcv_item[net_rcv_i])
                net_rcv_i += 1
            sdk.set_led_colors_flush_buffer()
        except Exception as e:
            # print('NetworkMonClass:', e)
            pass
        self.terminate()


class PingTestClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.ping_bool = False
        self.ping_bool_prev = None
        self.ping_fail_i = 0

    def run(self):
        global key_board, allow_mon_threads_bool
        print('-- thread started: HddMonClass(QThread).run(self)')
        while True:
            try:
                if len(key_board) >= 1 or len(mouse_device) >= 1 and allow_mon_threads_bool is True:
                    self.send_instruction()
                    time.sleep(5)
                else:
                    time.sleep(3)
            except Exception as e:
                print('-- exception in PingTestClass:', e)

    def send_instruction(self):
        global mouse_led, sdk, mouse_device, mouse_device_selected, m_key_id, ping_test_key_id, ping_test_f11, ping_test_f12
        global network_adapter_led_snt_item_bytes, network_adapter_led_snt_item_unit, mon_threads
        self.ping()
        if self.ping_fail_i == 1:
            self.ping()
        if self.ping_fail_i == 2:
            self.ping_bool = False
            self.ping_fail_i = 0
        if self.ping_bool is True and self.ping_bool != self.ping_bool_prev:
            print('-- sending instruction: ping True')
            mon_threads[4].start()
            if len(mouse_device) >= 1:
                if len(m_key_id) >= 3:
                    sdk.set_led_colors_buffer_by_device_index(mouse_device[mouse_device_selected], ({m_key_id[3]: (0, 255, 0)}))
            if len(key_board) >= 1:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[0])
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], ({ping_test_f11: (0, 255, 0)}))
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], ({ping_test_f12: (0, 255, 0)}))
            self.ping_bool_prev = True
        elif self.ping_bool is False and self.ping_bool != self.ping_bool_prev:
            print('-- sending instruction: ping False')
            mon_threads[4].stop()
            time.sleep(0.1)
            if len(mouse_device) >= 1:
                if len(m_key_id) >= 3:
                    sdk.set_led_colors_buffer_by_device_index(mouse_device[mouse_device_selected], ({m_key_id[3]: (255, 0, 0)}))
            if len(key_board) >= 1:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], network_adapter_led_snt_item_unit[0])
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], ({ping_test_f11: (255, 0, 0)}))
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], ({ping_test_f12: (255, 0, 0)}))
                for _ in network_adapter_led_snt_item_bytes:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], _)
            self.ping_bool_prev = False
        sdk.set_led_colors_flush_buffer()

    def ping(self):
        self.ping_bool = False
        cmd = 'ping -n 1 -l 1 8.8.8.8'  # Google
        if self.ping_fail_i == 1:
            cmd = 'ping -n 1 -l 1 9.9.9.9'  # Quad9. The free DNS service was co-developed by the Global Cyber Alliance, IBM, and Packet Clearing House.
        try:
            p = subprocess.Popen(cmd, stdout=subprocess.PIPE)
            stdout, stderror = p.communicate()
            output = stdout.decode('UTF-8')
            lines = output.split(os.linesep)
            for _ in lines:
                if 'Packets: Sent = 1, Received = 0, Lost = 1 (100% loss)' in _:
                    self.ping_fail_i += 1
                if 'Packets: Sent = 1, Received = 1, Lost = 0 (0% loss)' in _:
                    self.ping_bool = True
        except Exception as e:
            print('-- exception in PingTestClass.ping:', e)


class HddMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.hdd_display_key_bool_prev = []
        for _ in alpha_led_id:
            self.hdd_display_key_bool_prev.append(False)

    def run(self):
        pythoncom.CoInitialize()
        global key_board, allow_mon_threads_bool
        print('-- thread started: HddMonClass(QThread).run(self)')
        while True:
            if len(key_board) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(hdd_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global hdd_display_key_bool, sdk, key_board, key_board_selected, hdd_led_off_item
        global hdd_led_item
        self.get_stat()
        hdd_i = 0
        try:
            for _ in hdd_display_key_bool:
                if hdd_display_key_bool[hdd_i] is True and self.hdd_display_key_bool_prev[hdd_i] != hdd_display_key_bool[hdd_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], hdd_led_item[hdd_i])
                    self.hdd_display_key_bool_prev[hdd_i] = True
                    # print('-- setting drive letter:', hdd_led_item[hdd_i], 'True')
                elif hdd_display_key_bool[hdd_i] is False and self.hdd_display_key_bool_prev[hdd_i] != hdd_display_key_bool[hdd_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], hdd_led_off_item[hdd_i])
                    self.hdd_display_key_bool_prev[hdd_i] = False
                    # print('-- setting drive letter:', hdd_led_item[hdd_i], 'False')
                hdd_i += 1
        except Exception as e:
            print('HddMonClass:', e)
        sdk.set_led_colors_flush_buffer()

    def get_stat(self):
        global hdd_display_key_bool, alpha_str, hdd_display_key_bool
        try:
            hdd_display_key_bool = []
            for _ in alpha_led_id:
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
        global sdk, key_board, key_board_selected, hdd_led_off_item
        try:
            hdd_i = 0
            for _ in hdd_led_off_item:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], hdd_led_off_item[hdd_i])
                hdd_i += 1
            sdk.set_led_colors_flush_buffer()
        except Exception as e:
            # print('HddMonClass:', e)
            pass
        self.terminate()


class CpuMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.cpu_display_key_bool_tmp_0 = [True, False, False, False]
        self.cpu_display_key_bool_tmp_1 = [True, True, False, False]
        self.cpu_display_key_bool_tmp_2 = [True, True, True, False]
        self.cpu_display_key_bool_tmp_3 = [True, True, True, True]
        self.cpu_display_key_bool_prev = [False, False, False, False]
        self.switch_count = 0

    def run(self):
        global key_board
        print('-- thread started: CpuMonClass(QThread).run(self)')

        while True:
            if len(key_board) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(cpu_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global cpu_display_key_bool, sdk, key_board, key_board_selected
        global cpu_led_item, cpu_led_off_item
        self.get_stat()
        cpu_i = 0
        try:
            for _ in cpu_led_item:
                if cpu_display_key_bool[cpu_i] is True and self.cpu_display_key_bool_prev[cpu_i] != cpu_display_key_bool[cpu_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], cpu_led_item[cpu_i])
                    self.cpu_display_key_bool_prev[cpu_i] = True
                    print(self.switch_count, '-- setting cpu key:', cpu_led_item[cpu_i], 'True')
                    self.switch_count += 1
                elif cpu_display_key_bool[cpu_i] is False and self.cpu_display_key_bool_prev[cpu_i] != cpu_display_key_bool[cpu_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], cpu_led_off_item[cpu_i])
                    self.cpu_display_key_bool_prev[cpu_i] = False
                    print(self.switch_count, '-- setting cpu key:', cpu_led_item[cpu_i], 'False')
                    self.switch_count += 1
                cpu_i += 1
        except Exception as e:
            print('CpuMonClass:', e)
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
        global sdk, key_board, key_board_selected, cpu_led_off_item
        try:
            cpu_i = 0
            for _ in cpu_led_off_item:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], cpu_led_off_item[cpu_i])
                cpu_i += 1
            sdk.set_led_colors_flush_buffer()
        except Exception as e:
            # print('-- CpuMonClass:', e)
            pass
        self.terminate()


class DramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.dram_display_key_bool_tmp_0 = [True, False, False, False]
        self.dram_display_key_bool_tmp_1 = [True, True, False, False]
        self.dram_display_key_bool_tmp_2 = [True, True, True, False]
        self.dram_display_key_bool_tmp_3 = [True, True, True, True]
        self.dram_display_key_bool_prev = [False, False, False, False]
        self.switch_count = 0

    def run(self):
        global key_board
        print('-- thread started: DramMonClass(QThread).run(self)')
        while True:
            if len(key_board) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(dram_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global dram_display_key_bool, sdk, key_board, key_board_selected
        self.get_stat()
        dram_i = 0
        try:
            for _ in dram_led_item:
                if dram_display_key_bool[dram_i] is True and self.dram_display_key_bool_prev[dram_i] != dram_display_key_bool[dram_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], dram_led_item[dram_i])
                    self.dram_display_key_bool_prev[dram_i] = True
                    print(self.switch_count, '-- setting dram key:', dram_led_item[dram_i], 'True')
                    self.switch_count += 1

                elif dram_display_key_bool[dram_i] is False and self.dram_display_key_bool_prev[dram_i] != dram_display_key_bool[dram_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], dram_led_off_item[dram_i])
                    self.dram_display_key_bool_prev[dram_i] = False
                    print(self.switch_count, '-- setting dram key:', dram_led_off_item[dram_i], 'False')
                    self.switch_count += 1
                dram_i += 1
        except Exception as e:
            print('DramMonClass:', e)
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
        global sdk, key_board, key_board_selected, dram_led_off_item
        try:
            dram_i = 0
            for _ in dram_led_off_item:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], dram_led_off_item[dram_i])
                dram_i += 1
            sdk.set_led_colors_flush_buffer()
        except Exception as e:
            # print('-- DramMonClass:', e)
            pass
        self.terminate()


class VramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)
        self.vram_display_key_bool_tmp_0 = [True, False, False, False]
        self.vram_display_key_bool_tmp_1 = [True, True, False, False]
        self.vram_display_key_bool_tmp_2 = [True, True, True, False]
        self.vram_display_key_bool_tmp_3 = [True, True, True, True]
        self.vram_display_key_bool_prev = [False, False, False, False]
        self.switch_count = 0

    def run(self):
        global key_board
        print('-- thread started: VramMonClass(QThread).run(self)')
        while True:
            if len(key_board) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
                time.sleep(vram_led_time_on)
            else:
                time.sleep(1)

    def send_instruction(self):
        global vram_display_key_bool, sdk, key_board, key_board_selected
        self.get_stat()
        vram_i = 0
        try:
            for _ in vram_led_item:
                if vram_display_key_bool[vram_i] is True and self.vram_display_key_bool_prev[vram_i] != vram_display_key_bool[vram_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], vram_led_item[vram_i])
                    self.vram_display_key_bool_prev[vram_i] = True
                    print(self.switch_count, '-- setting vram key:', vram_led_item[vram_i], 'True')
                    self.switch_count += 1

                elif vram_display_key_bool[vram_i] is False and self.vram_display_key_bool_prev[vram_i] != vram_display_key_bool[vram_i]:
                    sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], vram_led_off_item[vram_i])
                    self.vram_display_key_bool_prev[vram_i] = False
                    print(self.switch_count, '-- setting vram key:', vram_led_item[vram_i], 'True')
                    self.switch_count += 1

                vram_i += 1
        except Exception as e:
            print('VramMonClass:', e)
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
        global sdk, key_board, key_board_selected, vram_led_off_item
        try:
            vram_i = 0
            for _ in vram_led_off_item:
                sdk.set_led_colors_buffer_by_device_index(key_board[key_board_selected], vram_led_off_item[vram_i])
                vram_i += 1
            sdk.set_led_colors_flush_buffer()
        except Exception as e:
            # print('-- VramMonClass:', e)
            pass
        self.terminate()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
