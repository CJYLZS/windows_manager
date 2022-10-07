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
        if os.path.basename(winfo.file_name) not in self.programs:
            return

        info(winfo)
        if self.visible(winfo):
            self.windows_info_dict[hwnd] = winfo
        if event == win32con.EVENT_SYSTEM_FOREGROUND:
            if hwnd != self.now_hwnd and not self.__alt:
                if hwnd in self.windows_list:
                    self.windows_list.remove(hwnd)
                    self.ws_idx = self.windows_list.index(self.now_hwnd)
                self.windows_list = self.windows_list[:self.ws_idx + 1]
                self.windows_list.append(hwnd)
                self.ws_idx += 1
                self.now_hwnd = hwnd

    def onRALTDown(self):
        self.__alt = True
        return 0

    def onRALTUp(self):
        self.__alt = False
        self.flash_back()
        return 0

    def onBackDown(self):
        if not self.__alt:
            return win32con.HC_ACTION
        self.back(flash=True)
        return win32con.HC_SKIP

    def onBackUp(self):
        return win32con.HC_ACTION

    def onFrontDown(self):
        if not self.__alt:
            return win32con.HC_ACTION

        self.front(flash=True)

        return win32con.HC_SKIP

    def onFrontUp(self):
        return win32con.HC_ACTION

    def __init__(self) -> None:
        with open('config.json') as f:
            self.config = json.load(f)
        self.programs = self.config["programs"]
        self.shell = win32com.client.Dispatch("WScript.Shell")
        self.now_hwnd = win32gui.GetForegroundWindow()
        self.windows_list = [self.now_hwnd]
        self.ws_idx = 0
        self.windows_hook = WindowsHook(listenEvents=[
            (win32con.EVENT_SYSTEM_FOREGROUND, self.__update_window_info),
            (win32con.EVENT_SYSTEM_MOVESIZEEND, self.__update_window_info),
            (win32con.EVENT_OBJECT_LOCATIONCHANGE, self.__update_window_info)
        ])
        self.windows_info_dict = {
            self.now_hwnd: self.windows_hook.get_window_info(self.now_hwnd)
        }
        self.anykey = AnyKey()
        self.anykey.register_hotkeys(
            # do not use lalt!!
            (('ralt',), self.onRALTDown, self.onRALTUp, 0),
            (('h',), self.onBackDown, self.onBackUp, 0),
            (('l',), self.onFrontDown, self.onFrontUp, 0)
        )
        self.current_state = ''
        self.__alt = False
        self.__last_change_time = time.time()
        self.__change_interval = 0.2
        self.__event_loop = EventLoop()

    def visible(self, winfo: WindowInfo):
        rect = winfo.rectangle
        return not(rect[0] <= 0 and rect[1] <= 0 and rect[2] <= 0 and rect[3] <= 0)

    def stop(self):
        self.__event_loop.stop()

    def start(self):
        self.__event_loop.start()

    def move_foreground(self, hwnd):
        d = self.windows_info_dict[hwnd]
        rect = d.rectangle
        win32gui.ShowWindow(hwnd, d.showCmd)
        # win32gui.SetWindowPos(
        #     hwnd, win32con.HWND_TOP, rect[0], rect[1], rect[2], rect[3], win32con.SWP_SHOWWINDOW)
        # https://stackoverflow.com/questions/14295337/win32gui-setactivewindow-error-the-specified-procedure-could-not-be-found
        self.shell.SendKeys('%')
        win32gui.SetForegroundWindow(hwnd)
        info(self.windows_hook.get_window_info(hwnd))

    def back(self, flash=False):
        if self.ws_idx == 0:
            return
        if time.time() - self.__last_change_time < self.__change_interval:
            return
        self.__last_change_time = time.time()
        self.ws_idx -= 1
        hwnd = self.windows_list[self.ws_idx]
        self.move_foreground(hwnd)
        if not flash:
            self.now_hwnd = hwnd

    def front(self, flash=False):
        if self.ws_idx == len(self.windows_list) - 1:
            return
        if time.time() - self.__last_change_time < self.__change_interval:
            return
        self.__last_change_time = time.time()
        self.ws_idx += 1
        hwnd = self.windows_list[self.ws_idx]
        self.move_foreground(hwnd)
        if not flash:
            self.now_hwnd = hwnd

    def flash_back(self):
        if self.windows_list[self.ws_idx] == self.now_hwnd:
            return
        self.ws_idx = self.windows_list.index(self.now_hwnd)
        self.move_foreground(self.now_hwnd)

    def show_window(self, hwnd):
        d = self.windows_info_dict[hwnd]
        info(
            f'{hwnd: <10}{d.file_name[-20:]: <30}{d.title[:50]: <60}{d.rectangle}')


if __name__ == '__main__':
    time.sleep(5)
    wm = WindowManager()
    wm.start()
