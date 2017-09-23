# A sample context menu handler.
# Adds a 'Hello from Python' menu entry to .py files.  When clicked, a
# simple message box is displayed.
#
# To demostrate:
# * Execute this script to register the context menu.
# * Open Windows Explorer, and browse to a directory with a .py file.
# * Right-Click on a .py file - locate and click on 'Hello from Python' on
#   the context menu.

import pythoncom
from win32com.shell import shell, shellcon
import win32gui
import win32con

import win32file
import json
from os import path

FILE = "MizTagger.json"
APP = "MizTagger"

class ShellExtension:
    _reg_progid_ = "Python.ShellExtension.ContextMenu"
    _reg_desc_ = "Python Sample Shell Extension (context menu)"
    _reg_clsid_ = "{CED0336C-C9EE-4a7f-8D7F-C660393C381F}"
    _com_interfaces_ = [shell.IID_IShellExtInit, shell.IID_IContextMenu]
    _public_methods_ = shellcon.IContextMenu_Methods + shellcon.IShellExtInit_Methods

    def Initialize(self, folder, dataobj, hkey):
        print("Init", folder, dataobj, hkey)
        self.dataobj = dataobj
        self.data = {"tags": ["Unix", "Windows", "MIZip"], "maps": {}}
        # maps: {"uid": tagum}
        self.fnames = []
        self.uids = []

    def QueryContextMenu(self, hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags):
        print("QCM", hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)
        # Query the items clicked on
        format_etc = win32con.CF_HDROP, None, 1, -1, pythoncom.TYMED_HGLOBAL
        sm = self.dataobj.GetData(format_etc)
        num_files = shell.DragQueryFile(sm.data_handle, -1)

        filepath = path.join(
            path.dirname(shell.DragQueryFile(sm.data_handle, 0)),
            FILE)
        self.filepath = filepath

        if path.exists(filepath):
            with open(filepath, "r", encoding="utf8") as fp:
                self.data = json.load(fp)

        for i in range(num_files):
            fname = shell.DragQueryFile(sm.data_handle, i)
            self.fnames.append(fname)
            with open(fname, "rb") as fp:
                handle = win32file._get_osfhandle(fp.fileno())
                info = win32file.GetFileInformationByHandle(handle)
                uid = str((info[8]<<32)+info[9])
                self.uids.append(uid)
                if uid not in self.data["maps"]:
                    self.data["maps"][uid] = 0b0

        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1

        flag = win32con.MF_STRING|win32con.MF_BYPOSITION
        tagum = 0b0
        idCmd = idCmdFirst
        if num_files>1:
            pass
            # if more than one file selected, then add the tag to all of them.
            # what if want to remove the tag from all of them? rare case, right?
        else:
            tagum = self.data["maps"][self.uids[0]]
            # read tagum of this file from data/storage
        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        for i, item in enumerate(self.data["tags"]):
            tagunit = 0b1 << i
            f = flag
            if (tagum & tagunit) == tagunit:
                f |= win32con.MF_CHECKED
            win32gui.InsertMenu(submenu, subindex,
                                f,
                                idCmd, item)
            subindex += 1
            idCmd += 1
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_POPUP|win32con.MF_STRING|win32con.MF_BYPOSITION,
                            submenu, APP)
        indexMenu += 1

        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1
        return idCmd-idCmdFirst # Must return number of menu items we added.

    def InvokeCommand(self, ci):
        mask, hwnd, verb, params, dir, nShow, hotkey, hicon = ci
        tagunit = 0b1 << verb
        if len(self.fnames)>1:
            for uid in self.uids:
                self.data["maps"][uid] |= tagunit
        else:
            self.data["maps"][self.uids[0]] ^= tagunit
        with open(self.filepath, "w", encoding="utf8") as fp:
            json.dump(self.data, fp, indent=2)
        # should pay attention to the 'verb', may related to idCmd

    def GetCommandString(self, cmd, typ):
        # If GetCommandString returns the same string for all items then
        # the shell seems to ignore all but one.  This is even true in
        # Win7 etc where there is no status bar (and hence this string seems
        # ignored)
        return "Hello from Python (cmd=%d)!!" % (cmd,)

def DllRegisterServer():
    import winreg
    key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT,
                            "*\\shellex")
    subkey = winreg.CreateKey(key, "ContextMenuHandlers")
    subkey2 = winreg.CreateKey(subkey, "PythonSample")
    winreg.SetValueEx(subkey2, None, 0, winreg.REG_SZ, ShellExtension._reg_clsid_)
    print(ShellExtension._reg_desc_, "registration complete.")

def DllUnregisterServer():
    import winreg
    try:
        key = winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT,
                                "*\\shellex\\ContextMenuHandlers\\PythonSample")
    except WindowsError as details:
        import errno
        if details.errno != errno.ENOENT:
            raise
    print(ShellExtension._reg_desc_, "unregistration complete.")

if __name__=='__main__':
    from win32com.server import register
    register.UseCommandLine(ShellExtension,
                   finalize_register = DllRegisterServer,
                   finalize_unregister = DllUnregisterServer)