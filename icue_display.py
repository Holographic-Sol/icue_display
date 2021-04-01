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


def NFD(text):
    return unicodedata.normalize('NFD', text)


def canonical_caseless(text):
    return NFD(NFD(text).casefold())


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

out_of_bounds = False
glo_obj = []
prev_obj_eve = []

mon_threads = []
conf_thread = []
allow_mon_threads_bool = False
connected_bool = None
connected_bool_prev = None
allow_configuration_read_bool = False

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
hdd_led_time_on = 0.05
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
cpu_led_time_on = 0.05
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
dram_led_time_on = 0.05
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
vram_led_time_on = 0.05
vram_led_item = [({118: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 3
    ({115: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 6
    ({111: (vram_led_color[0], vram_led_color[1], vram_led_color[2])}),  # 9
    ({105: (vram_led_color[0], vram_led_color[1], vram_led_color[2])})]  # keypad_asterisk
vram_led_off_item = [({118: (0, 0, 0)}),
                ({115: (0, 0, 0)}),
                ({111: (0, 0, 0)}),
                ({105: (0, 0, 0)})]

exclusive_access_bool = False
exclusive_access_bool_prev = None
vram_initiation = False
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

        self.cursorMove.connect(self.handleCursorMove)
        self.timer = QTimer(self)
        self.timer.setInterval(50)
        self.timer.timeout.connect(self.pollCursor)
        self.timer.start()
        self.cursor = None

        self.filter = ObjEveFilter()

        self.setWindowIcon(QIcon('./icon.png'))
        self.title = '{dev gui}'
        print('-- setting self.title as:', self.title)
        self.setWindowTitle(self.title)

        self.width = 605
        self.height = 110
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
        global mon_threads, conf_thread

        compile_devices_thread = CompileDevicesClass()
        compile_devices_thread.start()

        read_configuration_thread = ReadConfigurationClass()
        conf_thread.append(read_configuration_thread)

        hdd_mon_thread = HddMonClass()
        mon_threads.append(hdd_mon_thread)

        cpu_mon_thread = CpuMonClass()
        mon_threads.append(cpu_mon_thread)

        dram_mon_thread = DramMonClass()
        mon_threads.append(dram_mon_thread)

        vram_mon_thread = VramMonClass()
        mon_threads.append(vram_mon_thread)

        print('\ndisplaying application:')
        self.show()

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
        out_of_bounds = True
        if pos.x() > self.x():
            if pos.x() < (self.x() + self.width):
                if pos.y() < (self.y() + self.height):
                    if pos.y() > self.y():
                        if self.isMinimized() is False:
                            out_of_bounds = False
        if out_of_bounds is True:
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
        global sdk, k95_rgb_platinum
        global allow_configuration_read_bool
        global connected_bool, connected_bool_prev
        global exclusive_access_bool_prev
        global mon_threads, conf_thread

        while True:
            connected = sdk.connect()

            if not connected:
                err = sdk.get_last_error()
                # print('[NAME]: CompileDevicesClass [FUNCTION]: run [MESSAGE]: Handshake failed: %s' % err)
                connected_bool = False
                allow_configuration_read_bool = False

            elif connected:
                device = sdk.get_devices()
                i = 0
                for _ in device:
                    target_name = str(device[i])
                    if 'K95 RGB PLATINUM' in target_name:
                        k95_rgb_platinum.append(i)
                        # print('[NAME]: CompileDevicesClass [FUNCTION]: run [MESSAGE]: Found Device:', i)
                    i += 1

                if len(k95_rgb_platinum) >= 1:
                    connected_bool = True
                    allow_configuration_read_bool = True

            if connected_bool is False and connected_bool != connected_bool_prev:
                print('stopping threads: configuration read and instructions', )
                conf_thread[0].stop()
                mon_threads[0].stop()
                mon_threads[1].stop()
                mon_threads[2].stop()
                mon_threads[3].stop()
                connected_bool_prev = False
                exclusive_access_bool_prev = None
            elif connected_bool is True and connected_bool != connected_bool_prev:
                print('starting threads: configuration read and instructions', )
                conf_thread[0].start()
                mon_threads[0].start()
                mon_threads[1].start()
                mon_threads[2].start()
                mon_threads[3].start()
                connected_bool_prev = True
                exclusive_access_bool_prev = None
            time.sleep(3)


class ReadConfigurationClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def exclusive_access(self):
        global exclusive_access_bool, exclusive_access_bool_prev, k95_rgb_platinum, k95_rgb_platinum_selected
        with open('.\\config.dat', 'r') as fo:
            for line in fo:
                line = line.strip()
                if line.startswith('exclusive_access: '):
                    if line == 'exclusive_access: true':
                        # print('exclusive_access: true')
                        exclusive_access_bool = True
                    elif line == 'exclusive_access: false':
                        # print('exclusive_access: false')
                        exclusive_access_bool = False

        if exclusive_access_bool is True and exclusive_access_bool != exclusive_access_bool_prev:
            print('exclusive access request changed: requesting control')
            sdk.request_control()
            itm = {1: (255, 0, 0)}
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], itm)
            exclusive_access_bool_prev = True

        elif exclusive_access_bool is False and exclusive_access_bool != exclusive_access_bool_prev:
            print('exclusive access request changed: releasing control')
            sdk.release_control()
            itm = {1: (255, 0, 0)}
            sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], itm)
            exclusive_access_bool_prev = False

    def hdd_sanitize(self):
        global hdd_led_color, hdd_led_color_off, hdd_led_time_on, hdd_led_item, hdd_led_off_item
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
                            default_hdd_led_bool = [True, True, True]
                            if hdd_led_color[0] >= 0 and hdd_led_color[0] <= 255:
                                default_hdd_led_bool[0] = False
                            else:
                                default_hdd_led_bool[0] = True
                            if hdd_led_color[1] >= 0 and hdd_led_color[1] <= 255:
                                default_hdd_led_bool[1] = False
                            else:
                                default_hdd_led_bool[1] = True
                            if hdd_led_color[2] >= 0 and hdd_led_color[2] <= 255:
                                default_hdd_led_bool[2] = False
                            else:
                                default_hdd_led_bool[2] = True
                            if True not in default_hdd_led_bool:
                                # print('hdd_led_color: passed sanitization checks. using custom color')
                                i = 0
                                hdd_led_item = []
                                for _ in alpha_led:
                                    itm = {alpha_led[i]: hdd_led_color}
                                    hdd_led_item.append(itm)
                                    i += 1
                                # print(hdd_led_item)
                            elif True in default_hdd_led_bool:
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
                            default_hdd_led_bool = [True, True, True]
                            if hdd_led_color_off[0] >= 0 and hdd_led_color_off[0] <= 255:
                                default_hdd_led_bool[0] = False
                            else:
                                default_hdd_led_bool[0] = True
                            if hdd_led_color_off[1] >= 0 and hdd_led_color_off[1] <= 255:
                                default_hdd_led_bool[1] = False
                            else:
                                default_hdd_led_bool[1] = True
                            if hdd_led_color_off[2] >= 0 and hdd_led_color_off[2] <= 255:
                                default_hdd_led_bool[2] = False
                            else:
                                default_hdd_led_bool[2] = True
                            if True not in default_hdd_led_bool:
                                # print('hdd_led_color_off: passed sanitization checks. using custom color')
                                i = 0
                                hdd_led_off_item = []
                                for _ in alpha_led:
                                    itm = {alpha_led[i]: hdd_led_color_off}
                                    hdd_led_off_item.append(itm)
                                    i += 1
                                # print(hdd_led_item)
                            elif True in default_hdd_led_bool:
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
        global vram_led_color, vram_led_time_on, gpu_num, vram_led_item
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
        global allow_configuration_read_bool, allow_mon_threads_bool
        try:
            print('allow_configuration_read_bool', allow_configuration_read_bool)
            if allow_configuration_read_bool is True:
                self.exclusive_access()
                self.hdd_sanitize()
                self.cpu_sanitize()
                self.dram_sanitize()
                self.vram_sanitize()
                allow_configuration_read_bool = False
                allow_mon_threads_bool = True
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
        time.sleep(hdd_led_time_on)

    def get_stat(self):
        print()
        global hdd_display_key_bool, alpha_str, hdd_display_key_bool
        try:
            hdd_display_key_bool = []
            for _ in alpha_led:
                hdd_display_key_bool.append(False)
            str_computer = "."
            obj_wmi_service = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            obj_swbem_services = obj_wmi_service.ConnectServer(str_computer, "root\\cimv2")
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
        self.terminate()


class CpuMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: CpuMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
            else:
                time.sleep(1)

    def send_instruction(self):
        global cpu_initiation, cpu_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if cpu_initiation is False:
            for _ in cpu_led_off_item:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], _)
            cpu_initiation = True
        i = 0
        for _ in cpu_led_item:
            if cpu_display_key_bool[i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], cpu_led_item[i])
            elif cpu_display_key_bool[i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], cpu_led_off_item[i])
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

    def stop(self):
        print('-- stopping: CpuMonClass')
        self.terminate()


class DramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: DramMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
            else:
                time.sleep(1)

    def send_instruction(self):
        global dram_initiation, dram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if dram_initiation is False:
            for _ in dram_led_off_item:
                try:
                    sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], _)
                except Exception as e:
                    print('[NAME]: DramMonClass [FUNCTION]: send_instruction [EXCEPTION]:', e)
            dram_initiation = True
        i = 0
        for _ in dram_led_item:
            if dram_display_key_bool[i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], dram_led_item[i])
            elif dram_display_key_bool[i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], dram_led_off_item[i])
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

    def stop(self):
        print('-- stopping: DramMonClass')
        self.terminate()


class VramMonClass(QThread):
    def __init__(self):
        QThread.__init__(self)

    def run(self):
        global k95_rgb_platinum
        print('-- thread started: VramMonClass(QThread).run(self)')
        while True:
            if len(k95_rgb_platinum) >= 1 and allow_mon_threads_bool is True:
                self.send_instruction()
            else:
                time.sleep(1)

    def send_instruction(self):
        global vram_initiation, vram_display_key_bool, sdk, k95_rgb_platinum, k95_rgb_platinum_selected
        self.get_stat()
        if vram_initiation is False:
            for _ in vram_led_off_item:
                try:
                    sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], _)
                except Exception as e:
                    print('[NAME]: VramMonClass [FUNCTION]: send_instruction [EXCEPTION]:', e)
            vram_initiation = True
        i = 0
        for _ in vram_led_item:
            if vram_display_key_bool[i] is True:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], vram_led_item[i])
            elif vram_display_key_bool[i] is False:
                sdk.set_led_colors_buffer_by_device_index(k95_rgb_platinum[k95_rgb_platinum_selected], vram_led_off_item[i])
            i += 1
        sdk.set_led_colors_flush_buffer()
        time.sleep(vram_led_time_on)

    def get_stat(self):
        global vram_stat, vram_display_key_bool, gpu_num
        try:
            gpus = GPUtil.getGPUs()
            if len(gpus) >= 0:
                vram_stat = float(f"{gpus[gpu_num].load * 100}")
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

    def stop(self):
        print('-- stopping: VramMonClass')
        self.terminate()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
