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

    def __update_window_info(self, hwnd, event):
        if not hwnd:
            return
        winfo = self.windows_hook.get_window_info(hwnd)
        if os.path.basename(winfo.file_name) in self.exclude_programs:
            return
        if not self.visible(winfo):
            return

        info(winfo)
        self.windows_info_dict[hwnd] = winfo
        if event == win32con.EVENT_SYSTEM_FOREGROUND:
            if hwnd != self.current_hwnd and not self.__mod1_pressed:
                if hwnd in self.windows_list:
                    self.windows_list.remove(hwnd)
                    self.ws_idx = self.windows_list.index(self.current_hwnd)
                self.windows_list = self.windows_list[:self.ws_idx + 1]
                self.windows_list.append(hwnd)
                self.ws_idx += 1
                self.current_hwnd = hwnd

    def onModKey1Down(self):
        self.__mod1_pressed = True
        return 0

    def onModKey1Up(self):
        self.__mod1_pressed = False
        self.flash_back()
        return 0

    def onModKey2Down(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__mod2_pressed = True
        return win32con.HC_ACTION

    def onModKey2Up(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.__mod2_pressed = False
        return win32con.HC_ACTION

    def onBackDown(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.back()
        if self.__mod2_pressed:
            self.update_current_hwnd()
        return win32con.HC_SKIP

    def onBackUp(self):
        return win32con.HC_SKIP

    def onFrontDown(self):
        if not self.__mod1_pressed:
            return win32con.HC_ACTION
        self.front()
        if self.__mod2_pressed:
            self.update_current_hwnd()
        return win32con.HC_SKIP

    def onFrontUp(self):
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
            ((hotkey["mod_key"],), self.onModKey1Down, self.onModKey1Up, 0),
            ((hotkey["move_mod_key"],),
             self.onModKey2Down, self.onModKey2Up, 0),
            ((hotkey["back_key"],), self.onBackDown, self.onBackUp, 0),
            ((hotkey["front_key"],), self.onFrontDown, self.onFrontUp, 0)
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

    def stop(self):
        self.__event_loop.stop()

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

    def back(self):
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

    def front(self):
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

    def flash_back(self):
        if self.windows_list[self.ws_idx] == self.current_hwnd:
            return
        self.ws_idx = self.windows_list.index(self.current_hwnd)
        self.move_foreground(self.current_hwnd)

    def update_current_hwnd(self):
        self.current_hwnd = self.windows_list[self.ws_idx]

    def show_window(self, hwnd):
        d = self.windows_info_dict[hwnd]
        info(
            f'{hwnd: <10}{d.file_name[-20:]: <30}{d.title[:50]: <60}{d.rectangle}')


if __name__ == '__main__':
    wm = WindowManager()
    wm.start()
