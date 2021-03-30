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
from win32api import GetMonitorInfo, MonitorFromPoint
from PyQt5.QtCore import Qt, QThread, QSize, QPoint, QCoreApplication, QObject, QTimer
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QDesktopWidget
from PyQt5.QtGui import QIcon, QCursor
from PyQt5 import QtCore
from cuesdk import CueSdk
import time
import GPUtil
import psutil
import pythoncom
import unicodedata

print('initializing:')
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


def NFD(text):
    return unicodedata.normalize('NFD', text)


def canonical_caseless(text):
    return NFD(NFD(text).casefold())


out_of_bounds = False
glo_obj = []
prev_obj_eve = []

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
hdd_stat = []
hdd_led_color = [255, 255, 255]
hdd_led_color_off = [0, 0, 0]
hdd_led_time_on = 0.05
hdd_initiation = False
allow_hdd_mon_thread_bool = True
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
print('hdd_display_key_bool:', len(hdd_display_key_bool))
print('hdd_led_item:', len(hdd_led_item))
print('hdd_display_key_bool:', len(hdd_display_key_bool))
print('alpha_led:', len(alpha_led))

cpu_stat = ()
cpu_led_color = [255, 255, 255]
cpu_led_time_on = 0.05
cpu_led_item = [({116: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 2
    ({113: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 5
    ({109: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 8
    ({103: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])})]   # /
cpu_initiation = False
allow_cpu_mon_thread_bool = True
cpu_display_key_bool = [False, False, False, False]
cpu_led_off_item = [({116: (0, 0, 0)}),
                ({113: (0, 0, 0)}),
                ({109: (0, 0, 0)}),
                ({103: (0, 0, 0)})]

dram_stat = ()
dram_led_color = [255, 255, 255]
dram_led_time_on = 0.05
dram_led_item = [({117: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 2
    ({114: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 5
    ({110: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 8
    ({104: (dram_led_color[0], dram_led_color[1], dram_led_color[2])})]   # /
dram_initiation = False
allow_dram_mon_thread_bool = True
dram_display_key_bool = [False, False, False, False]
dram_led_off_item = [({117: (0, 0, 0)}),
                ({114: (0, 0, 0)}),
                ({110: (0, 0, 0)}),
                ({104: (0, 0, 0)})]

gpu_num = ()
vram_stat = ()
vram_led_color = [255, 255, 255]
vram_led_time_on = 0.05
vram_led_item = [({118: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 3
    ({115: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 6
    ({111: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 9
    ({105: (vram_led_color[0], vram_led_color[1], vram_led_color[2])})]  # keypad_asterisk
vram_led_off_item = [({118: (0, 0, 0)}),
                ({115: (0, 0, 0)}),
                ({111: (0, 0, 0)}),
                ({105: (0, 0, 0)})]

exclusiv_access_bool_initialize = True
exclusiv_access_bool = False
vram_initiation = False
allow_vram_mon_thread_bool = True
vram_display_key_bool = [False, False, False, False]


class ObjEveFilter(QObject):
    def eventFilter(self, obj, event):
        global glo_obj, prev_obj_eve, out_of_bounds
        obj_eve = obj, event

        # Uncomment This Line To See All Object Events
        # print('-- ObjEveFilter(QObject).eventFilter(self, obj, event):', obj_eve)

        # Filtered Object Events
        if out_of_bounds is False:
            if str(obj_eve[1]).startswith('<PyQt5.QtGui.QEnterEvent') and obj_eve[0] == glo_obj[1]:
                self.unhighlightObject()
            elif str(obj_eve[1]).startswith('<PyQt5.QtGui.QEnterEvent') and obj_eve[0] == glo_obj[2]:
                self.unhighlightObject()
            elif str(obj_eve[1]).startswith('<PyQt5.QtGui.QHoverEvent') and obj_eve[0] == glo_obj[3]:
                self.unhighlightObject()
                glo_obj[3].setStyleSheet(
                        """QPushButton{background-color: rgb(255, 0, 0);
                           border:0px solid rgb(0, 0, 0);}"""
                )
            elif str(obj_eve[1]).startswith('<PyQt5.QtGui.QHoverEvent') and obj_eve[0] == glo_obj[4]:
                self.unhighlightObject()
                glo_obj[4].setStyleSheet(
                    """QPushButton{background-color: rgb(0, 0, 255);
                       border:0px solid rgb(0, 0, 0);}"""
                )
        return False

    def unhighlightObject(self):
        glo_obj[2].setStyleSheet(
            """QLabel {background-color: rgb(15, 15, 15);
           border:0px solid rgb(35, 35, 35);}"""
        )
        # print('-- ObjEveFilter(QObject).unhighlightObject(self):', glo_obj[2])
        glo_obj[3].setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        # print('-- ObjEveFilter(QObject).unhighlightObject(self):', glo_obj[3])
        glo_obj[4].setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        # print('-- ObjEveFilter(QObject).unhighlightObject(self):', glo_obj[4])


class App(QMainWindow):
    cursorMove = QtCore.pyqtSignal(object)

    def __init__(self):
        super(App, self).__init__()
        global glo_obj

        self.filter = ObjEveFilter()

        self.setWindowIcon(QIcon('./icon.png'))
        self.title = '{dev gui}'
        print('-- setting self.title as:', self.title)
        self.setWindowTitle(self.title)

        self.width = 605
        self.height = 110
        self.prev_width = ()
        self.prev_height = ()
        self.prev_pos_w = ()
        self.prev_pos_h = ()
        self.prev_pos = self.pos()
        pos_w = QDesktopWidget().availableGeometry().width()
        pos_h = QDesktopWidget().availableGeometry().height()
        pos_w = (pos_w / 2) - (self.width / 2)
        pos_h = (pos_h / 2) - (self.height / 2)
        print('-- setting window dimensions:', self.width, self.height)
        self.setGeometry(pos_w, pos_h, self.width, self.height)

        p = self.palette()
        p.setColor(self.backgroundRole(), Qt.black)
        self.setWindowFlags(Qt.Window | Qt.FramelessWindowHint)
        self.setPalette(p)

        self.cursorMove.connect(self.handleCursorMove)
        self.timer = QTimer(self)
        self.timer.setInterval(50)
        self.timer.timeout.connect(self.pollCursor)
        self.timer.start()
        self.cursor = None

        self.btn_title_logo = QPushButton(self)
        self.btn_title_logo.move(0, 0)
        self.btn_title_logo.resize(20, 20)
        # self.btn_title_logo.setIcon(QIcon("./image/img_logo_titlebar.png"))
        self.btn_title_logo.setIconSize(QSize(20, 20))
        self.btn_title_logo.setStyleSheet(
            """QPushButton{background-color: rgb(0, 0, 0);
               border:0px solid rgb(0, 0, 0);}"""
        )
        self.btn_title_logo.installEventFilter(self.filter)
        print('-- created btn_title_logo', self.btn_title_logo)
        glo_obj.append(self.btn_title_logo)

        self.lbl_title_bg = QLabel(self)
        self.lbl_title_bg.move(20, 0)
        self.lbl_title_bg.resize(self.width - 40, 20)
        self.lbl_title_bg.setStyleSheet(
            """QLabel {background-color: rgb(0, 0, 0);
           border:0px solid rgb(35, 35, 35);}"""
        )
        self.lbl_title_bg.installEventFilter(self.filter)
        print('-- created lbl_title_bg', self.lbl_title_bg)
        glo_obj.append(self.lbl_title_bg)

        self.lbl_main_bg = QLabel(self)
        self.lbl_main_bg.move(0, 20)
        self.lbl_main_bg.resize(self.width, self.height)
        self.lbl_main_bg.setStyleSheet(
            """QLabel {background-color: rgb(15, 15, 15);
           border:0px solid rgb(35, 35, 35);}"""
        )
        self.lbl_main_bg.installEventFilter(self.filter)
        print('-- created lbl_main_bg', self.lbl_main_bg)
        glo_obj.append(self.lbl_main_bg)

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
        self.btn_quit.installEventFilter(self.filter)
        print('-- created self.btn_quit', self.btn_quit)
        glo_obj.append(self.btn_quit)

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
        self.btn_minimize.installEventFilter(self.filter)
        print('-- created self.btn_minimize', self.btn_minimize)
        glo_obj.append(self.btn_minimize)

        self.initUI()

    def initUI(self):
        scaling_thread = ScalingClass(self.setGeometry, self.width, self.height, self.pos, self.frameGeometry,
                                      self.setFixedSize)
        scaling_thread.start()

        compile_devices_thread = CompileDevicesClass()
        compile_devices_thread.start()

        read_configuration_thread = ReadConfigurationClass()
        read_configuration_thread.start()

        hdd_mon_thread = HddMonClass()
        hdd_mon_thread.start()

        cpu_mon_thread = CpuMonClass()
        cpu_mon_thread.start()

        dram_mon_thread = DramMonClass()
        dram_mon_thread.start()

        vram_mon_thread = VramMonClass()
        vram_mon_thread.start()

        print('\ndisplaying application:')
        self.show()

    def center(self):
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def mousePressEvent(self, event):
        self.prev_pos = event.globalPos()

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
        if pos.x() > self.x() and pos.x() < (self.x() + self.width) and\
                pos.y() < (self.y() + self.height) and pos.y() > self.y() and self.isMinimized() is False:
            # print('-- App(QMainWindow).handleCursorMove(self, pos):', pos)
            out_of_bounds = False
        else:
            out_of_bounds = True
            glo_obj[2].setStyleSheet(
                """QLabel {background-color: rgb(15, 15, 15);
               border:0px solid rgb(35, 35, 35);}"""
            )
            glo_obj[3].setStyleSheet(
                """QPushButton{background-color: rgb(0, 0, 0);
                   border:0px solid rgb(0, 0, 0);}"""
            )
            glo_obj[4].setStyleSheet(
                """QPushButton{background-color: rgb(0, 0, 0);
                   border:0px solid rgb(0, 0, 0);}"""
            )


class CompileDevicesClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global sdk, k95_rgb_platinum, allow_vram_mon_thread_bool
        while True:
            connected = sdk.connect()
            if not connected:
                err = sdk.get_last_error()
                # print('[NAME]: CompileDevicesClass [FUNCTION]: run [MESSAGE]: Handshake failed: %s' % err)
                allow_vram_mon_thread_bool = False
            elif connected:
                device = sdk.get_devices()
                i = 0
                for _ in device:
                    target_name = str(device[i])
                    if 'K95 RGB PLATINUM' in target_name:
                        k95_rgb_platinum.append(i)
                        # print('[NAME]: CompileDevicesClass [FUNCTION]: run [MESSAGE]: Found Device:', i)
                        allow_vram_mon_thread_bool = True
                    i += 1
                else:
                    # print('[NAME]: CompileDevicesClass [FUNCTION]: run [MESSAGE]: Device Not Found')
                    time.sleep(1)
            time.sleep(3)


class ReadConfigurationClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def exclusiv_access(self):
        global exclusiv_access_bool, k95_rgb_platinum, k95_rgb_platinum_selected
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('exclusiv_access: '):
                    if line == 'exclusiv_access: true':
                        # print('exclusiv_access: true')
                        sdk.request_control()
                        exclusiv_access_bool = True
                        itm = {1: (255, 0, 0)}
                        sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], itm)
                    elif line == 'exclusiv_access: false':
                        # print('exclusiv_access: false')
                        sdk.release_control()
                        exclusiv_access_bool = False

    def hdd_sanitize(self):
        global hdd_led_color, hdd_led_time_on, hdd_led_item, hdd_led_color_off, hdd_led_off_item, exclusiv_access_bool
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('hdd_led_color: '):
                    hdd_led_color = line.replace('hdd_led_color: ', '')
                    hdd_led_color = hdd_led_color.split(',')
                    # print('hdd_led_color:', hdd_led_color)
                    if len(hdd_led_color) == 3:
                        # print('hdd_led_color: has 3 items')
                        if str(hdd_led_color[0]).isdigit() and str(hdd_led_color[1]).isdigit() and str(hdd_led_color[2]).isdigit():
                            # print('hdd_led_color 0-3: isdigits')
                            hdd_led_color[0] = int(hdd_led_color[0])
                            hdd_led_color[1] = int(hdd_led_color[1])
                            hdd_led_color[2] = int(hdd_led_color[2])
                            default_cpu_led_bool = [True, True, True]
                            if hdd_led_color[0] >= 0 and hdd_led_color[0] <= 255:
                                default_cpu_led_bool[0] = False
                            else:
                                default_cpu_led_bool[0] = True
                            if hdd_led_color[1] >= 0 and hdd_led_color[1] <= 255:
                                default_cpu_led_bool[1] = False
                            else:
                                default_cpu_led_bool[1] = True
                            if hdd_led_color[2] >= 0 and hdd_led_color[2] <= 255:
                                default_cpu_led_bool[2] = False
                            else:
                                default_cpu_led_bool[2] = True
                            if True not in default_cpu_led_bool:
                                # print('hdd_led_color: passed sanitization checks. using custom color')
                                i = 0
                                hdd_led_item = []
                                for _ in alpha_led:
                                    itm = {alpha_led[i]: hdd_led_color}
                                    hdd_led_item.append(itm)
                                    i += 1
                                # print(hdd_led_item)
                            elif True in default_cpu_led_bool:
                                # print('hdd_led_color: failed sanitization checks. using default color')
                                hdd_led_color = [255, 255, 255]
                if line.startswith('hdd_led_color_off: '):
                    hdd_led_color_off = line.replace('hdd_led_color_off: ', '')
                    hdd_led_color_off = hdd_led_color_off.split(',')
                    # print('hdd_led_color_off:', hdd_led_color_off)
                    if len(hdd_led_color_off) == 3:
                        # print('hdd_led_color_off: has 3 items')
                        if str(hdd_led_color_off[0]).isdigit() and str(hdd_led_color_off[1]).isdigit() and str(hdd_led_color_off[2]).isdigit():
                            # print('hdd_led_color_off 0-3: isdigits')
                            hdd_led_color_off[0] = int(hdd_led_color_off[0])
                            hdd_led_color_off[1] = int(hdd_led_color_off[1])
                            hdd_led_color_off[2] = int(hdd_led_color_off[2])
                            default_cpu_led_bool = [True, True, True]
                            if hdd_led_color_off[0] >= 0 and hdd_led_color_off[0] <= 255:
                                default_cpu_led_bool[0] = False
                            else:
                                default_cpu_led_bool[0] = True
                            if hdd_led_color_off[1] >= 0 and hdd_led_color_off[1] <= 255:
                                default_cpu_led_bool[1] = False
                            else:
                                default_cpu_led_bool[1] = True
                            if hdd_led_color_off[2] >= 0 and hdd_led_color_off[2] <= 255:
                                default_cpu_led_bool[2] = False
                            else:
                                default_cpu_led_bool[2] = True
                            if True not in default_cpu_led_bool:
                                # print('hdd_led_color_off: passed sanitization checks. using custom color')
                                i = 0
                                hdd_led_off_item = []
                                for _ in alpha_led:
                                    itm = {alpha_led[i]: hdd_led_color_off}
                                    hdd_led_off_item.append(itm)
                                    i += 1
                                # print(hdd_led_item)
                            elif True in default_cpu_led_bool:
                                # print('hdd_led_color_off: failed sanitization checks. using default color')
                                hdd_led_color_off = [0, 0, 0]
                if line.startswith('hdd_led_time_on: '):
                    line = line.replace('hdd_led_time_on: ', '')
                    # print('hdd_led_time_on:', line)
                    try:
                        # print('hdd_led_time_on: is decimal. using custom hdd_led_time_on value')
                        hdd_led_time_on = float(line)
                    except Exception as e:
                        # print('hdd_led_time_on: is not decimal. using default hdd_led_time_on value')
                        hdd_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: cpu_sanitize [EXCEPTION]:', e)

    def cpu_sanitize(self):
        global cpu_led_color, cpu_led_time_on, cpu_led_item
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('cpu_led_color: '):
                    cpu_led_color = line.replace('cpu_led_color: ', '')
                    cpu_led_color = cpu_led_color.split(',')
                    # print('cpu_led_color:', cpu_led_color)
                    if len(cpu_led_color) == 3:
                        # print('cpu_led_color: has 3 items')
                        if str(cpu_led_color[0]).isdigit() and str(cpu_led_color[1]).isdigit() and str(cpu_led_color[2]).isdigit():
                            # print('cpu_led_color 0-3: isdigits')
                            cpu_led_color[0] = int(cpu_led_color[0])
                            cpu_led_color[1] = int(cpu_led_color[1])
                            cpu_led_color[2] = int(cpu_led_color[2])
                            default_cpu_led_bool = [True, True, True]
                            if cpu_led_color[0] >= 0 and cpu_led_color[0] <= 255:
                                default_cpu_led_bool[0] = False
                            else:
                                default_cpu_led_bool[0] = True
                            if cpu_led_color[1] >= 0 and cpu_led_color[1] <= 255:
                                default_cpu_led_bool[1] = False
                            else:
                                default_cpu_led_bool[1] = True
                            if cpu_led_color[2] >= 0 and cpu_led_color[2] <= 255:
                                default_cpu_led_bool[2] = False
                            else:
                                default_cpu_led_bool[2] = True
                            if True not in default_cpu_led_bool:
                                # print('cpu_led_color: passed sanitization checks. using custom color')
                                cpu_led_item = [
                                    ({116: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 1
                                    ({113: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 4
                                    ({109: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])}),  # 7
                                    ({103: (cpu_led_color[0], cpu_led_color[1], cpu_led_color[2])})]  # num
                            elif True in default_cpu_led_bool:
                                # print('cpu_led_color: failed sanitization checks. using default color')
                                cpu_led_color = [255, 255, 255]
                if line.startswith('cpu_led_time_on: '):
                    line = line.replace('cpu_led_time_on: ', '')
                    # print('cpu_led_time_on:', line)
                    try:
                        # print('cpu_led_time_on: is decimal. using custom cpu_led_time_on value')
                        cpu_led_time_on = float(line)
                    except Exception as e:
                        # print('cpu_led_time_on: is not decimal. using default cpu_led_time_on value')
                        cpu_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: cpu_sanitize [EXCEPTION]:', e)

    def dram_sanitize(self):
        global dram_led_color, dram_led_time_on, dram_led_item
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('dram_led_color: '):
                    dram_led_color = line.replace('dram_led_color: ', '')
                    dram_led_color = dram_led_color.split(',')
                    # print('dram_led_color:', dram_led_color)
                    if len(dram_led_color) == 3:
                        # print('dram_led_color: has 3 items')
                        if str(dram_led_color[0]).isdigit() and str(dram_led_color[1]).isdigit() and str(dram_led_color[2]).isdigit():
                            # print('dram_led_color 0-3: isdigits')
                            dram_led_color[0] = int(dram_led_color[0])
                            dram_led_color[1] = int(dram_led_color[1])
                            dram_led_color[2] = int(dram_led_color[2])
                            default_dram_led_bool = [True, True, True]
                            if dram_led_color[0] >= 0 and dram_led_color[0] <= 255:
                                default_dram_led_bool[0] = False
                            else:
                                default_dram_led_bool[0] = True
                            if dram_led_color[1] >= 0 and dram_led_color[1] <= 255:
                                default_dram_led_bool[1] = False
                            else:
                                default_dram_led_bool[1] = True
                            if dram_led_color[2] >= 0 and dram_led_color[2] <= 255:
                                default_dram_led_bool[2] = False
                            else:
                                default_dram_led_bool[2] = True
                            if True not in default_dram_led_bool:
                                # print('dram_led_color: passed sanitization checks. using custom color')
                                dram_led_item = [
                                    ({117: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 2
                                    ({114: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 5
                                    ({110: (dram_led_color[0], dram_led_color[1], dram_led_color[2])}),  # 8
                                    ({104: (dram_led_color[0], dram_led_color[1], dram_led_color[2])})]  # /
                            elif True in default_dram_led_bool:
                                # print('dram_led_color: failed sanitization checks. using default color')
                                dram_led_color = [255, 255, 255]
                if line.startswith('dram_led_time_on: '):
                    line = line.replace('dram_led_time_on: ', '')
                    # print('dram_led_time_on:', line)
                    try:
                        # print('dram_led_time_on: is decimal. using custom dram_led_time_on value')
                        dram_led_time_on = float(line)
                    except Exception as e:
                        # print('dram_led_time_on: is not decimal. using default dram_led_time_on value')
                        dram_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: dram_sanitize [EXCEPTION]:', e)

    def vram_sanitize(self):
        global vram_led_color, vram_led_time_on, vram_led_item, gpu_num
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('vram_led_color: '):
                    vram_led_color = line.replace('vram_led_color: ', '')
                    vram_led_color = vram_led_color.split(',')
                    # print('vram_led_color:', vram_led_color)
                    if len(vram_led_color) == 3:
                        # print('vram_led_color: has 3 items')
                        if str(vram_led_color[0]).isdigit() and str(vram_led_color[1]).isdigit() and str(vram_led_color[2]).isdigit():
                            # print('vram_led_color 0-3: isdigits')
                            vram_led_color[0] = int(vram_led_color[0])
                            vram_led_color[1] = int(vram_led_color[1])
                            vram_led_color[2] = int(vram_led_color[2])
                            default_vram_led_bool = [True, True, True]
                            if vram_led_color[0] >= 0 and vram_led_color[0] <= 255:
                                default_vram_led_bool[0] = False
                            else:
                                default_vram_led_bool[0] = True
                            if vram_led_color[1] >= 0 and vram_led_color[1] <= 255:
                                default_vram_led_bool[1] = False
                            else:
                                default_vram_led_bool[1] = True
                            if vram_led_color[2] >= 0 and vram_led_color[2] <= 255:
                                default_vram_led_bool[2] = False
                            else:
                                default_vram_led_bool[2] = True
                            if True not in default_vram_led_bool:
                                # print('vram_led_color: passed sanitization checks. using custom color')
                                vram_led_item = [
                                    ({118: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 3
                                    ({115: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 6
                                    ({111: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 9
                                    ({105: (vram_led_color[0], vram_led_color[1], vram_led_color[2])})]  # *
                            elif True in default_vram_led_bool:
                                # print('vram_led_color: failed sanitization checks. using default color')
                                vram_led_color = [255, 255, 255]
                if line.startswith('vram_led_time_on: '):
                    line = line.replace('vram_led_time_on: ', '')
                    # print('vram_led_time_on:', line)
                    try:
                        # print('vram_led_time_on: is decimal. using custom vram_led_time_on value')
                        vram_led_time_on = float(line)
                    except Exception as e:
                        # print('vram_led_time_on: is not decimal. using default vram_led_time_on value')
                        vram_led_time_on = 0.5
                        print('[NAME]: ReadConfigurationClass [FUNCTION]: vram_sanitize [EXCEPTION]:', e)
                if line.startswith('gpu_num: '):
                    line = line.replace('gpu_num: ', '')
                    # print('gpu_num:', line)
                    if line.isdigit():
                        # print('gpu_num: is digit')
                        line = int(line)
                        gpus = GPUtil.getGPUs()
                        if len(gpus) >= line:
                            # print('gpu_num: exists. using custom gpu_num value')
                            gpu_num = line
                        else:
                            print('gpu_num: may exceed gpus currently active on the system. using default value')
                    else:
                        # print('gpu_num: is not digit. using default gpu_num value')
                        gpu_num = 0

    def run(self):
        global exclusiv_access_bool
        while True:
            try:
                self.exclusiv_access()
                self.hdd_sanitize()
                self.cpu_sanitize()
                self.dram_sanitize()
                self.vram_sanitize()
            except Exception as e:
                print('[NAME]: ReadConfigurationClass [FUNCTION]: run [EXCEPTION]:', e)
            time.sleep(3)


class HddMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        pythoncom.CoInitialize()
        global allow_hdd_mon_thread_bool, sdk, k95_rgb_platinum
        print('-- thread started: HddMonClass(QThread).run(self)')

        while True:
            if allow_hdd_mon_thread_bool is True:
                if len(k95_rgb_platinum) >= 1:
                    self.send_instruction()
            else:
                time.sleep(3)

    def send_instruction(self):
        global hdd_initiation, hdd_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected, hdd_led_off_item
        global hdd_led_item, alpha_led
        self.get_stat()

        hdd_i = 0
        for _ in hdd_display_key_bool:
            if hdd_display_key_bool[hdd_i] is True:
                # print('setting hdd_led_item:', hdd_led_item[hdd_i])
                kb_on_dict_0 = hdd_led_item[hdd_i]

                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            elif hdd_display_key_bool[hdd_i] is False:
                kb_on_dict_0 = hdd_led_off_item[hdd_i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            hdd_i += 1
        sdk.set_led_colors_flush_buffer()
        time.sleep(hdd_led_time_on)

    def get_stat(self):
        print()
        global hdd_stat, hdd_display_key_bool, alpha_str, hdd_display_key_bool
        try:
            hdd_display_key_bool = []
            for _ in alpha_led:
                hdd_display_key_bool.append(False)
            strComputer = "."
            objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            objSWbemServices = objWMIService.ConnectServer(strComputer, "root\\cimv2")
            colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_PerfFormattedData_PerfDisk_PhysicalDisk")
            for objItem in colItems:
                if objItem.DiskBytesPersec != None:
                    if '_Total' not in objItem.Name:
                        var = objItem.Name.split()
                        try:
                            disk_letter = var[1]
                            disk_letter = disk_letter.replace(':', '')
                            if len(disk_letter) == 1:
                                if int(objItem.DiskBytesPersec) > 0:
                                    i = 0
                                    for _ in alpha_str:
                                        if canonical_caseless(disk_letter) == canonical_caseless(alpha_str[i]):
                                            # print('pairing disk letter:', disk_letter, 'to', 'hdd_display_key_bool:', i)
                                            hdd_display_key_bool[i] = True
                                        i += 1
                        except:
                            pass
        except Exception as e:
            print('[NAME]: HddMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()


class CpuMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global allow_cpu_mon_thread_bool, sdk, k95_rgb_platinum
        print('-- thread started: CpuMonClass(QThread).run(self)')

        while True:
            if allow_cpu_mon_thread_bool is True:
                if len(k95_rgb_platinum) >= 1:
                    self.send_instruction()
            else:
                time.sleep(3)

    def send_instruction(self):
        global cpu_initiation, cpu_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if cpu_initiation is False:
            for _ in cpu_led_off_item:
                kb_on_dict_0 = _
                try:
                    sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
                except Exception as e:
                    print('[NAME]: CpuMonClass [FUNCTION]: send_instruction [EXCEPTION]:', e)
            cpu_initiation = True
        i = 0
        for _ in cpu_led_item:
            if cpu_display_key_bool[i] is True:
                kb_on_dict_0 = cpu_led_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            elif cpu_display_key_bool[i] is False:
                kb_on_dict_0 = cpu_led_off_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            i += 1
        sdk.set_led_colors_flush_buffer()
        time.sleep(cpu_led_time_on)

    def get_stat(self):
        global cpu_stat, cpu_display_key_bool
        try:
            cpu_stat = psutil.cpu_percent(0.1)
            if cpu_stat < 25:
                cpu_display_key_bool[0] = True
                cpu_display_key_bool[1] = False
                cpu_display_key_bool[2] = False
                cpu_display_key_bool[3] = False
            elif cpu_stat >= 25 and cpu_stat < 50:
                cpu_display_key_bool[0] = True
                cpu_display_key_bool[1] = True
                cpu_display_key_bool[2] = False
                cpu_display_key_bool[3] = False
            elif cpu_stat >= 50 and cpu_stat < 75:
                cpu_display_key_bool[0] = True
                cpu_display_key_bool[1] = True
                cpu_display_key_bool[2] = True
                cpu_display_key_bool[3] = False
            elif cpu_stat >= 75:
                cpu_display_key_bool[0] = True
                cpu_display_key_bool[1] = True
                cpu_display_key_bool[2] = True
                cpu_display_key_bool[3] = True
        except Exception as e:
            print('[NAME]: CpuMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()


class DramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global allow_dram_mon_thread_bool, sdk, k95_rgb_platinum
        print('-- thread started: DramMonClass(QThread).run(self)')

        while True:
            if allow_dram_mon_thread_bool is True:
                if len(k95_rgb_platinum) >= 1:
                    self.send_instruction()
            else:
                time.sleep(3)

    def send_instruction(self):
        global dram_initiation, dram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if dram_initiation is False:
            for _ in dram_led_off_item:
                kb_on_dict_0 = _
                try:
                    sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
                except Exception as e:
                    print('[NAME]: DramMonClass [FUNCTION]: send_instruction [EXCEPTION]:', e)
            dram_initiation = True
        i = 0
        for _ in dram_led_item:
            if dram_display_key_bool[i] is True:
                kb_on_dict_0 = dram_led_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            elif dram_display_key_bool[i] is False:
                kb_on_dict_0 = dram_led_off_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            i += 1
        sdk.set_led_colors_flush_buffer()
        time.sleep(dram_led_time_on)

    def get_stat(self):
        global dram_stat, dram_display_key_bool
        try:
            dram_stat = psutil.virtual_memory().percent

            if dram_stat < 25:
                dram_display_key_bool[0] = True
                dram_display_key_bool[1] = False
                dram_display_key_bool[2] = False
                dram_display_key_bool[3] = False
            elif dram_stat >= 25 and dram_stat < 50:
                dram_display_key_bool[0] = True
                dram_display_key_bool[1] = True
                dram_display_key_bool[2] = False
                dram_display_key_bool[3] = False
            elif dram_stat >= 50 and dram_stat < 75:
                dram_display_key_bool[0] = True
                dram_display_key_bool[1] = True
                dram_display_key_bool[2] = True
                dram_display_key_bool[3] = False
            elif dram_stat >= 75:
                dram_display_key_bool[0] = True
                dram_display_key_bool[1] = True
                dram_display_key_bool[2] = True
                dram_display_key_bool[3] = True
        except Exception as e:
            print('[NAME]: DramMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()


class VramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global allow_vram_mon_thread_bool, sdk, k95_rgb_platinum
        print('-- thread started: VramMonClass(QThread).run(self)')

        while True:
            if allow_vram_mon_thread_bool is True:
                if len(k95_rgb_platinum) >= 1:
                    self.send_instruction()
            else:
                time.sleep(3)

    def send_instruction(self):
        global vram_initiation, vram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if vram_initiation is False:
            for _ in vram_led_off_item:
                kb_on_dict_0 = _
                try:
                    sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
                except Exception as e:
                    print('[NAME]: VramMonClass [FUNCTION]: send_instruction [EXCEPTION]:', e)
            vram_initiation = True
        i = 0
        for _ in vram_led_item:
            if vram_display_key_bool[i] is True:
                kb_on_dict_0 = vram_led_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            elif vram_display_key_bool[i] is False:
                kb_on_dict_0 = vram_led_off_item[i]
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], kb_on_dict_0)
            i += 1
        sdk.set_led_colors_flush_buffer()
        time.sleep(vram_led_time_on)

    def get_stat(self):
        global vram_stat, vram_display_key_bool, gpu_num
        try:
            gpus = GPUtil.getGPUs()
            if len(gpus) >= 0:
                vram_stat = float(f"{gpus[0].load * 100}")
            vram_stat = int(vram_stat)
            if vram_stat < 25:
                vram_display_key_bool[0] = True
                vram_display_key_bool[1] = False
                vram_display_key_bool[2] = False
                vram_display_key_bool[3] = False
            elif vram_stat >= 25 and vram_stat < 50:
                vram_display_key_bool[0] = True
                vram_display_key_bool[1] = True
                vram_display_key_bool[2] = False
                vram_display_key_bool[3] = False
            elif vram_stat >= 50 and vram_stat < 75:
                vram_display_key_bool[0] = True
                vram_display_key_bool[1] = True
                vram_display_key_bool[2] = True
                vram_display_key_bool[3] = False
            elif vram_stat >= 75:
                vram_display_key_bool[0] = True
                vram_display_key_bool[1] = True
                vram_display_key_bool[2] = True
                vram_display_key_bool[3] = True
        except Exception as e:
            print('[NAME]: VramMonClass [FUNCTION]: get_stat [EXCEPTION]:', e)
            sdk.set_led_colors_flush_buffer()


class ScalingClass(QThread):
    def __init__(self, setGeometry, width, height, pos, frameGeometry, setFixedSize):
        QThread.__init__(self)
        self.setGeometry = setGeometry
        self.width = width
        self.height = height
        self.pos = pos
        self.frameGeometry = frameGeometry
        self.setFixedSize = setFixedSize
        self.pos_x = ()
        self.pos_y = ()

    def run(self):
        print('-- thread started: ScalingClass(QThread).run(self)')

        # Store Work Area Geometry For Comparison
        monitor_info = GetMonitorInfo(MonitorFromPoint((0, 0)))
        work_area = monitor_info.get("Work")
        scr_geo0 = work_area[3]

        while True:

            # Get Work Area Geometry Each Loop
            try:
                monitor_info = GetMonitorInfo(MonitorFromPoint((0, 0)))
                work_area = monitor_info.get("Work")
                scr_geo1 = work_area[3]
            except Exception as e:
                print('-- ScalingClass(QThread).run(self):', e)

            # Compare Current Work Area Geometry To Stored Work Area Geometry
            if scr_geo0 != scr_geo1:

                try:
                    print('-- ScalingClass(QThread).run(self) ~ refreshing geometry')
                    time.sleep(0.5)
                    self.setGeometry(0, 0, self.width, self.height)
                    time.sleep(3)
                except Exception as e:
                    print('-- ScalingClass(QThread).run(self):', e)
                    time.sleep(3)

            time.sleep(0.1)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
