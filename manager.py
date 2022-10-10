from stat import FILE_ATTRIBUTE_NO_SCRUB_DATA
from cjutils.utils import *
sys.path.insert(0, './win32')
from keyboard import AnyKey
from windows import WindowsHook
from windows import WindowInfo
from loop import EventLoop
import win32con
import win32com.client
import win32gui


class WindowManager:

    def __add_new_hwnd(self, hwnd):
        if hwnd == self.current_hwnd:
            return
        if hwnd in self.windows_list:
            self.windows_list.remove(hwnd)
            self.ws_idx = self.windows_list.index(self.current_hwnd)
        self.windows_list = self.windows_list[:self.ws_idx + 1]
        self.windows_list.append(hwnd)
        self.ws_idx += 1
        self.current_hwnd = hwnd

    def __update_window_info(self, hwnd, event):
        if not hwnd:
            return
        winfo = self.windows_hook.get_window_info(hwnd)
        if not winfo.file_name:
            err(winfo, 'file name is None')
        if not winfo.file_name or os.path.basename(winfo.file_name) in self.exclude_programs:
            return
        if not self.visible(winfo):
            return

        self.windows_info_dict[hwnd] = winfo
        if event == win32con.EVENT_SYSTEM_FOREGROUND and not self.__mod1_pressed:
            self.__add_new_hwnd(hwnd)

    def __onModKey1Down(self):
        self.__mod1_pressed = True
        return 0

    def __onModKey1Up(self):
        self.__mod1_pressed = False
        self.__flash_back()
        return 0

    def __onModKey2Down(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__mod2_pressed = True
        return win32con.HC_ACTION

    def __onModKey2Up(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__mod2_pressed = False
        self.__move_current_hwnd_front()
        return win32con.HC_ACTION

    def __onBackDown(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__back()
        if self.__mod2_pressed:
            self.__update_current_hwnd()
        return win32con.HC_SKIP

    def __onBackUp(self):
        return win32con.HC_SKIP

    def __onFrontDown(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__front()
        if self.__mod2_pressed:
            self.__update_current_hwnd()
        return win32con.HC_SKIP

    def __onFrontUp(self):
        return win32con.HC_ACTION

    def __init__(self) -> None:
        with open('config.json') as f:
            self.config = json.load(f)

        self.exclude_programs = self.config["exclude_programs"]
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.current_hwnd = win32gui.GetForegroundWindow()
        self.windows_list = [self.current_hwnd]
        self.ws_idx = 0
        self.windows_hook = WindowsHook(listenEvents=[
            (win32con.EVENT_SYSTEM_FOREGROUND, self.__update_window_info),
            (win32con.EVENT_SYSTEM_MOVESIZEEND, self.__update_window_info),
            (win32con.EVENT_OBJECT_LOCATIONCHANGE, self.__update_window_info)
        ])
        self.windows_info_dict = {
            self.current_hwnd: self.windows_hook.get_window_info(
                self.current_hwnd)
        }
        self.anykey = AnyKey()
        hotkey = self.config["hotkey"]
        self.anykey.register_hotkeys(
            # do not use lalt!!
            ((hotkey["mod_key"],), self.__onModKey1Down, self.__onModKey1Up, 0),
            ((hotkey["move_mod_key"],),
             self.__onModKey2Down, self.__onModKey2Up, 0),
            ((hotkey["back_key"],), self.__onBackDown, self.__onBackUp, 0),
            ((hotkey["front_key"],), self.__onFrontDown, self.__onFrontUp, 0)
        )
        self.current_state = ''
        self.__mod1_pressed = False
        self.__mod2_pressed = False
        self.__last_change_time = time.time()
        self.__change_interval = 0.1
        self.__event_loop = EventLoop()

    def visible(self, winfo: WindowInfo):
        rect = winfo.rectangle
        if rect[0] <= 0 and rect[1] <= 0 and rect[2] <= 0 and rect[3] <= 0:
            return False
        if (rect[2] - rect[0]) * (rect[3] - rect[1]) <= 0:
            return False
        return True

    def start(self):
        self.__event_loop.start()

    def move_foreground(self, hwnd):
        d = self.windows_info_dict[hwnd]
        # rect = d.rectangle
        win32gui.ShowWindow(hwnd, d.showCmd)
        # win32gui.SetWindowPos(
        #     hwnd, win32con.HWND_TOP, rect[0], rect[1], rect[2], rect[3], win32con.SWP_SHOWWINDOW)
        # https://stackoverflow.com/questions/14295337/win32gui-setactivewindow-error-the-specified-procedure-could-not-be-found
        self.shell.SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)
        info(self.windows_hook.get_window_info(hwnd))

    def __back(self):
        if self.ws_idx == 0:
            return
        if time.time() - self.__last_change_time < self.__change_interval:
            return
        self.__last_change_time = time.time()
        self.ws_idx -= 1
        hwnd = self.windows_list[self.ws_idx]
        if not win32gui.IsWindow(hwnd):
            self.windows_list.remove(hwnd)
            return
        self.move_foreground(hwnd)

    def __front(self):
        if self.ws_idx == len(self.windows_list) - 1:
            return
        if time.time() - self.__last_change_time < self.__change_interval:
            return
        self.__last_change_time = time.time()
        self.ws_idx += 1
        hwnd = self.windows_list[self.ws_idx]
        if not win32gui.IsWindow(hwnd):
            self.windows_list.remove(hwnd)
            self.ws_idx -= 1
            return
        self.move_foreground(hwnd)

    def __flash_back(self):
        if self.windows_list[self.ws_idx] == self.current_hwnd:
            return
        if not win32gui.IsWindow(self.current_hwnd):
            hwnd = win32gui.GetForegroundWindow()
            self.__add_new_hwnd(hwnd)
        self.ws_idx = self.windows_list.index(self.current_hwnd)
        self.move_foreground(self.current_hwnd)

    def __update_current_hwnd(self):
        self.current_hwnd = self.windows_list[self.ws_idx]

    def __move_current_hwnd_front(self):
        self.windows_list.remove(self.current_hwnd)
        self.windows_list.append(self.current_hwnd)
        self.ws_idx = len(self.windows_list) - 1

    def show_window(self, hwnd):
        d = self.windows_info_dict[hwnd]
        info(
            f'{hwnd: <10}{d.file_name[-20:]: <30}{d.title[:50]: <60}{d.rectangle}')


if __name__ == '__main__':
    wm = WindowManager()
    wm.start()
