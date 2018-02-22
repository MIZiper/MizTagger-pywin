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

from os import path
import os

from shared import APP, TaggerFolder, TaggerFile

class ShellExtension:
    _reg_progid_ = "Python.ShellExtension.ContextMenu"
    _reg_desc_ = "Python Sample Shell Extension (context menu)"
    _reg_clsid_ = "{CED0336C-C9EE-4a7f-8D7F-C660393C381F}"
    _com_interfaces_ = [shell.IID_IShellExtInit, shell.IID_IContextMenu]
    _public_methods_ = shellcon.IContextMenu_Methods + shellcon.IShellExtInit_Methods

    def Initialize(self, folder, dataobj, hkey):
        print("Init", folder, dataobj, hkey)
        self.dataobj = dataobj

    def QueryContextMenu(self, hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags):
        print("QCM", hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)
        # Query the items clicked on
        format_etc = win32con.CF_HDROP, None, 1, -1, pythoncom.TYMED_HGLOBAL
        sm = self.dataobj.GetData(format_etc)
        num_files = shell.DragQueryFile(sm.data_handle, -1)

        firstFile = shell.DragQueryFile(sm.data_handle, 0)
        folder = path.dirname(firstFile)
        self.folderObject = TaggerFolder(folder)

        self.cmdUnitMap = []
        # fnames = [shell.DragQueryFile(sm.data_handle, i) for i in range(num_files)]
        # TaggerFile is not for multifiles, so just pause toggleTagunits
        self.fileObject = TaggerFile(self.folderObject, path.basename(firstFile))

        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1

        flag = win32con.MF_STRING|win32con.MF_BYPOSITION
        idCmd = idCmdFirst
        if num_files>1:
            pass
            # if more than one file selected, then add the tag to all of them.
            # what if want to remove the tag from all of them? rare case, right?
        else:
            win32gui.InsertMenu(hMenu, indexMenu, flag,
                                idCmd, self.fileObject.getTitle() or "- unspecified title -")
            indexMenu += 1
        idCmd += 1
        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        for tagClass in self.fileObject.getTag():
            win32gui.InsertMenu(submenu, subindex, flag | win32con.MF_DISABLED,
                                0, "- %s -" % tagClass["name"])
            subindex += 1
            for tagItem in tagClass["tags"]:
                self.cmdUnitMap.append(tagItem["tagunit"])
                f = flag
                if tagItem["checked"]:
                    f |= win32con.MF_CHECKED
                win32gui.InsertMenu(submenu, subindex,
                                    f,
                                    idCmd, tagItem["name"])
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

        from PyQt4 import QtGui
        app = QtGui.QApplication([])

        if verb==0:
            self.fileObject.displayFileDescription()
        else:
            self.fileObject.toggleTagunit(self.cmdUnitMap[verb-1])
        # should pay attention to the 'verb', may related to idCmd

    def GetCommandString(self, cmd, typ):
        # If GetCommandString returns the same string for all items then
        # the shell seems to ignore all but one.  This is even true in
        # Win7 etc where there is no status bar (and hence this string seems
        # ignored)
        return "Hello from Python (cmd=%d)!!" % (cmd,)

class ShellExtensionFolder:
    _reg_progid_ = "MizTagger.ShellExtension.FolderContextMenu"
    _reg_desc_ = "MizTagger context menu entries for folder"
    _reg_clsid_ = "{8921201f-9f10-4c0f-9018-fa15f98b5924}"
    _com_interfaces_ = [shell.IID_IShellExtInit, shell.IID_IContextMenu]
    _public_methods_ = shellcon.IContextMenu_Methods + shellcon.IShellExtInit_Methods

    def Initialize(self, folder, dataobj, hkey):
        fd = shell.SHGetPathFromIDList(folder).decode("utf8")
        self.folderObject = TaggerFolder(fd)

    def QueryContextMenu(self, hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags):
        idCmd = idCmdFirst
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1

        if self.folderObject.isResultFolder:
            itm = "Clean Result"
        else:
            itm = "Manage Tags"
        
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_STRING|win32con.MF_BYPOSITION,
                            idCmd, itm)
        idCmd += 1
        indexMenu += 1

        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        win32gui.InsertMenu(submenu, subindex, win32con.MF_STRING|win32con.MF_BYPOSITION, idCmd, "Complex Filter")
        subindex += 1
        idCmd += 1
        win32gui.InsertMenu(submenu, subindex, win32con.MF_STRING|win32con.MF_BYPOSITION, idCmd, "Refresh File Name")
        subindex += 1
        idCmd += 1
        win32gui.InsertMenu(submenu, subindex, win32con.MF_STRING|win32con.MF_BYPOSITION, idCmd, "Refresh UID")
        subindex += 1
        idCmd += 1
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_POPUP|win32con.MF_STRING|win32con.MF_BYPOSITION,
                            submenu, "Util")
        indexMenu += 1

        self.cmdUnitMap = []
        flag = win32con.MF_STRING|win32con.MF_BYPOSITION
        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        for tagClass in self.folderObject.getTag():
            win32gui.InsertMenu(submenu, subindex, flag | win32con.MF_DISABLED,
                                0, "- %s -" % tagClass["name"])
            subindex += 1
            for tagItem in tagClass["tags"]:
                self.cmdUnitMap.append(tagItem["tagunit"])
                f = flag
                if tagItem["checked"]:
                    f |= win32con.MF_CHECKED
                win32gui.InsertMenu(submenu, subindex,
                                    f,
                                    idCmd, tagItem["name"])
                subindex += 1
                idCmd += 1

        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_POPUP|win32con.MF_STRING|win32con.MF_BYPOSITION,
                            submenu, "Quick Filter")
        indexMenu += 1

        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1
        return idCmd-idCmdFirst # Must return number of menu items we added.

    def InvokeCommand(self, ci):
        mask, hwnd, verb, params, dir, nShow, hotkey, hicon = ci

        from PyQt4 import QtGui
        app = QtGui.QApplication([])

        if verb==0: # Manage Tags || Clean Result
            if self.folderObject.isResultFolder:
                resultFolder = self.folderObject.resultFolder
                for f in os.listdir(resultFolder):
                    os.unlink(path.join(resultFolder, f))
                os.rmdir(resultFolder)
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            else:
                self.folderObject.displayTagManager()
        elif verb in (2, 3):
            pass
        else:
            if verb==1: # Complex Filter
                if not self.folderObject.displayFilterManager():
                    return
            else:
                self.folderObject.toggleFilterTagunit(self.cmdUnitMap[verb-4])
            
            folder = self.folderObject.folder
            if not self.folderObject.isResultFolder:
                import tempfile, subprocess
                resultFolder = tempfile.mkdtemp(prefix="MizResult_", dir=folder)
                subprocess.Popen("explorer %s" % resultFolder)
            else:
                resultFolder = self.folderObject.resultFolder
                for f in os.listdir(resultFolder):
                    os.unlink(path.join(resultFolder, f))
            self.folderObject.saveLogic(resultFolder)
            for f in self.folderObject.getFilterResult():
                os.link(path.join(folder, f), path.join(resultFolder, f))
                # there's problem here when the file name was changed
                # but how to create hard link by index?
    
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

    
    key = winreg.CreateKey(winreg.HKEY_CLASSES_ROOT,
                            "Directory\\Background\\shellex")
    subkey = winreg.CreateKey(key, "ContextMenuHandlers")
    subkey2 = winreg.CreateKey(subkey, "PythonSample")
    winreg.SetValueEx(subkey2, None, 0, winreg.REG_SZ, ShellExtensionFolder._reg_clsid_)
    print(ShellExtensionFolder._reg_desc_, "registration complete.")

def DllUnregisterServer():
    import winreg
    try:
        key = winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT,
                                "*\\shellex\\ContextMenuHandlers\\PythonSample")
        key = winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT,
                                "Directory\\Background\\shellex\\ContextMenuHandlers\\PythonSample")
    except WindowsError as details:
        import errno
        if details.errno != errno.ENOENT:
            raise
    print(ShellExtension._reg_desc_, "unregistration complete.")
    print(ShellExtensionFolder._reg_desc_, "unregistration complete.")

if __name__=='__main__':
    from win32com.server import register
    register.UseCommandLine(ShellExtension, ShellExtensionFolder,
                   finalize_register = DllRegisterServer,
                   finalize_unregister = DllUnregisterServer)