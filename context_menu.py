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
import os
from PyQt4 import QtGui

FILE = "MizTagger.json"
RESULT = "MizTagger.rslt"
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
        self.data = {"tags": {}, "maps": {}}
        # maps {"uid": {"tagum", "title", "fname", "desc"}}
        self.uids = []

    def QueryContextMenu(self, hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags):
        print("QCM", hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags)
        # Query the items clicked on
        format_etc = win32con.CF_HDROP, None, 1, -1, pythoncom.TYMED_HGLOBAL
        sm = self.dataobj.GetData(format_etc)
        num_files = shell.DragQueryFile(sm.data_handle, -1)

        folder = path.dirname(shell.DragQueryFile(sm.data_handle, 0))
        rsltpath = path.join(folder, RESULT)
        if path.exists(rsltpath):
            self.folder = path.join(folder, '../')
            self.resultFolder = folder
        else:
            self.folder = folder
            self.resultFolder = None
        filepath = path.join(self.folder, FILE)
        self.filepath = filepath

        if path.exists(filepath):
            with open(filepath, "r", encoding="utf8") as fp:
                self.data = json.load(fp)
        self.cmdUnitMap = []

        for i in range(num_files):
            fname = shell.DragQueryFile(sm.data_handle, i)
            with open(fname, "rb") as fp:
                handle = win32file._get_osfhandle(fp.fileno())
                info = win32file.GetFileInformationByHandle(handle)
                uid = str((info[8]<<32)+info[9])
                self.uids.append(uid)
                if uid not in self.data["maps"]:
                    self.data["maps"][uid] = [0b0, path.basename(fname), "", ""]

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
            tagum = self.data["maps"][self.uids[0]][0]
            title = self.data["maps"][self.uids[0]][2] or "- unspecified title -"
            win32gui.InsertMenu(hMenu, indexMenu, flag, idCmd, title)
            indexMenu += 1
            # read tagum of this file from data/storage
        idCmd += 1
        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        for k, v in self.data["tags"].items():
            win32gui.InsertMenu(submenu, subindex, flag | win32con.MF_DISABLED, 0, "- %s -"%k)
            subindex += 1
            for kk, vv in v.items():
                tagunit = 0b1 << vv
                self.cmdUnitMap.append(vv)
                f = flag
                if (tagum & tagunit) == tagunit:
                    f |= win32con.MF_CHECKED
                win32gui.InsertMenu(submenu, subindex,
                                    f,
                                    idCmd, kk)
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
        import sys
        app = QtGui.QApplication(sys.argv)
        if verb==0:
            dlg = Window({"Title": self.data["maps"][self.uids[0]][2], "Description": self.data["maps"][self.uids[0]][3]})
            dlg.exec()
            if dlg.result():
                self.data["maps"][self.uids[0]][2] = dlg.title
                self.data["maps"][self.uids[0]][3] = dlg.desc
        else:
            tagunit = 0b1 << (self.cmdUnitMap[verb-1])
            if len(self.uids)>1:
                for uid in self.uids:
                    self.data["maps"][uid][0] |= tagunit
            else:
                self.data["maps"][self.uids[0]][0] ^= tagunit
        with open(self.filepath, "w", encoding="utf8") as fp:
            json.dump(self.data, fp, indent=2, sort_keys=True)
        sys.exit()
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
        rsltpath = path.join(fd, RESULT)
        if path.exists(rsltpath):
            self.folder = path.join(fd, '../')
            self.resultFolder = fd
        else:
            self.folder = fd
            self.resultFolder = None
        filepath = path.join(self.folder, FILE)
        self.filepath = filepath
        self.data = {"tags": {}, "maps": {}}
        if path.exists(filepath):
            with open(filepath, "r", encoding="utf8") as fp:
                self.data = json.load(fp)

    def QueryContextMenu(self, hMenu, indexMenu, idCmdFirst, idCmdLast, uFlags):
        idCmd = idCmdFirst
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_SEPARATOR|win32con.MF_BYPOSITION,
                            0, None)
        indexMenu += 1

        if self.resultFolder:
            itm = "Clean Result"
            with open(path.join(self.resultFolder, RESULT), "r", encoding="utf8") as fp:
                lg_all = json.load(fp)["all"]
        else:
            itm = "Manage Tags"
            lg_all = 0b0
        self.lg_all = lg_all
        
        win32gui.InsertMenu(hMenu, indexMenu,
                            win32con.MF_STRING|win32con.MF_BYPOSITION,
                            idCmd, itm)
        idCmd += 1
        indexMenu += 1

        self.cmdUnitMap = []
        flag = win32con.MF_STRING|win32con.MF_BYPOSITION
        submenu = win32gui.CreatePopupMenu()
        subindex = 0
        for k, v in self.data['tags'].items():
            win32gui.InsertMenu(submenu, subindex, flag | win32con.MF_DISABLED, 0, "- %s -"%k)
            subindex += 1
            for kk, vv in v.items():
                tagunit = 0b1 << vv
                self.cmdUnitMap.append(vv)
                f = flag
                if (lg_all & tagunit) == tagunit:
                    f |= win32con.MF_CHECKED
                win32gui.InsertMenu(submenu, subindex,
                                    f,
                                    idCmd, kk)
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
        if verb==0: # Manage Tags || Clean Result
            if self.resultFolder:
                for f in os.listdir(self.resultFolder):
                    os.unlink(path.join(self.resultFolder, f))
                os.rmdir(self.resultFolder)
                win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
            else:
                import sys
                app = QtGui.QApplication(sys.argv)
                dlg = TagManager(self.data["tags"])
                dlg.exec()
                
                with open(self.filepath, "w", encoding="utf8") as fp:
                    json.dump(self.data, fp, indent=2, sort_keys=True)
                sys.exit()
        else:
            tagunit = 0b1 << (self.cmdUnitMap[verb-1])
            lg_all = self.lg_all ^ tagunit
            
            if not self.resultFolder:
                import tempfile, subprocess
                self.resultFolder = tempfile.mkdtemp(prefix="MizResult_", dir=self.folder)
                subprocess.Popen("explorer %s"%self.resultFolder)
            else:
                for f in os.listdir(self.resultFolder):
                    os.unlink(path.join(self.resultFolder, f))
            rsltpath = path.join(self.resultFolder, RESULT)
            with open(rsltpath, 'w', encoding='utf8') as fp:
                json.dump({"all": lg_all}, fp, indent=2)
            if lg_all != 0b0:
                for uid, v in self.data['maps'].items():
                    if (v[0] & lg_all) == lg_all:
                        os.link(path.join(self.folder, v[1]), path.join(self.resultFolder, v[1]))
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

class Window(QtGui.QDialog):
    def __init__(self, config):
        super().__init__()
        self.setWindowTitle(APP)
        self.config = config
        self.createUI()

    def createUI(self):
        config = self.config
        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

        lblTitle = QtGui.QLabel("Title")
        txtName = QtGui.QLineEdit(config["Title"])
        lblDesc = QtGui.QLabel("Description")
        txtDesc = QtGui.QPlainTextEdit(config["Description"])

        layoutMain = QtGui.QVBoxLayout()
        layoutMain.addWidget(lblTitle)
        layoutMain.addWidget(txtName)
        layoutMain.addWidget(lblDesc)
        layoutMain.addWidget(txtDesc, stretch=1)
        layoutMain.addWidget(buttonBox)
        
        self.setLayout(layoutMain)
        self.txtName = txtName
        self.txtDesc = txtDesc

    def accept(self):
        self.title = self.txtName.text()
        self.desc = self.txtDesc.toPlainText()
        super().accept()

class TagManager(QtGui.QDialog):
    def __init__(self, config):
        super().__init__()
        self.setWindowTitle(APP)
        self.config = config
        s = 0
        for k, v in config.items():
            s += len(v)
        self.count = s
        self.createUI()
        self.switchClass(0)

    def createUI(self):
        config = self.config
        # buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok|QtGui.QDialogButtonBox.Cancel)
        # buttonBox.accepted.connect(self.accept)
        # buttonBox.rejected.connect(self.reject)

        lblClass = QtGui.QLabel("Tag Class")
        cmbClass = QtGui.QComboBox()
        lblTags = QtGui.QLabel("Tags")
        lstTags = QtGui.QListWidget()
        cmbClass.addItems(list(config.keys()))
        cmbClass.currentIndexChanged.connect(self.switchClass)
        txtInput = QtGui.QLineEdit()
        btnByClass = QtGui.QPushButton("Add Class")
        btnByTag = QtGui.QPushButton("Add Tag")
        btnByClass.clicked.connect(self.addClass)
        btnByTag.clicked.connect(self.addTag)
        hLayout = QtGui.QHBoxLayout()
        hLayout.addWidget(txtInput, stretch=1)
        hLayout.addWidget(btnByClass)
        hLayout.addWidget(btnByTag)

        layoutMain = QtGui.QVBoxLayout()
        layoutMain.addWidget(lblClass)
        layoutMain.addWidget(cmbClass)
        layoutMain.addWidget(lblTags)
        layoutMain.addWidget(lstTags, stretch=1)
        layoutMain.addLayout(hLayout)
        self.setLayout(layoutMain)

        self.lstTags = lstTags
        self.cmbClass = cmbClass
        self.txtInput = txtInput

    def switchClass(self, index):
        if len(self.config) > index:
            self.lstTags.clear()
            self.lstTags.addItems(list(self.config[self.cmbClass.itemText(index)].keys()))
    
    def addClass(self, b):
        c = self.txtInput.text()
        if c not in self.config:
            self.config[c] = {}
            self.cmbClass.addItem(c)
        self.txtInput.clear()
        self.txtInput.setFocus()

    def addTag(self, b):
        t = self.txtInput.text()
        if self.cmbClass.currentIndex() >= 0:
            self.count += 1
            self.config[self.cmbClass.currentText()][t] = self.count
            self.lstTags.addItem(t)
        self.txtInput.clear()
        self.txtInput.setFocus()

if __name__=='__main__':
    from win32com.server import register
    register.UseCommandLine(ShellExtension, ShellExtensionFolder,
                   finalize_register = DllRegisterServer,
                   finalize_unregister = DllUnregisterServer)