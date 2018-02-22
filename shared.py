from PyQt4 import QtGui, QtCore
import os, win32file, json
from os import path

FILE = "MizTagger.json"
RESULT = "MizTagger.rslt"
APP = "MizTagger"

# logic
# {"Contain One": 0b0, "Contain All": 0b0, "Not All": 0b0, "Not Any": 0b0}
# config
# {"[TagClass1]": {"[TagName1]": 1, "[TagName2]": 2}}

class TaggerFolder(object):
    def __init__(self, folder):
        self.data = {"tags": {}, "maps": {}}
        self.resetLogic()
        # maps {"uid": {"tagum", "title", "fname", "desc"}}
        rsltpath = path.join(folder, RESULT)
        if path.exists(rsltpath):
            self.folder = path.join(folder, '../')
            self.resultFolder = folder
            self.isResultFolder = True
            with open(rsltpath, "r", encoding="utf8") as fp:
                self.logic = json.load(fp)
        else:
            self.folder = folder
            self.resultFolder = None
            self.isResultFolder = False
        filepath = path.join(self.folder, FILE)
        self.filepath = filepath
        if path.exists(filepath):
            with open(filepath, "r", encoding="utf8") as fp:
                self.data = json.load(fp)

    def getTag(self):
        logic = self.logic["Contain All"]
        return [
            {
                "name": k,
                "tags": [
                    {
                        "name": kk,
                        "checked": (logic & (0b1<<vv)) == (0b1<<vv),
                        "tagunit": 0b1 << vv
                    } for kk, vv in v.items()
                ]
            } for k, v in self.data["tags"].items()
        ]

    def resetLogic(self):
        self.logic = {"Contain All": 0b0, "Contain One": 0b0, "Not Any": 0b0, "Not All": 0b0}

    def displayTagManager(self):
        dlg = TagManager(self.data["tags"])
        dlg.exec()
        self.saveData()

    def displayFilterManager(self):
        dlg = FilterManager(self.data["tags"], self.logic)
        dlg.exec()
        if dlg.result():
            self.logic = dlg.logic
        return dlg.result()

    def toggleFilterTagunit(self, tagunit):
        self.logic["Contain All"] ^= tagunit

    def getFilterResult(self):
        logic = self.logic
        r = []
        if sum(logic.values())!=0b0:
            for uid, v in self.data['maps'].items():
                if (logic["Contain All"]==0b0 or (v[0]&logic["Contain All"])==logic["Contain All"]) and \
                (logic["Contain One"]==0b0 or (v[0]&logic["Contain One"])!=0b0) and \
                (logic["Not All"]==0b0 or (v[0]&logic["Not All"])!=logic["Not All"]) and \
                (logic["Not Any"]==0b0 or (v[0]&logic["Not Any"])==0b0):
                    r.append(v[1])
        return r

    def saveLogic(self, resultFolder):
        rsltpath = path.join(resultFolder, RESULT)
        with open(rsltpath, 'w', encoding='utf8') as fp:
            json.dump(self.logic, fp, indent=2)

    def saveData(self):
        with open(self.filepath, "w", encoding="utf8") as fp:
            json.dump(self.data, fp, indent=2, sort_keys=True)

    def getInfoByUid(self, uid):
        return None if uid not in self.data["maps"] else self.data["maps"][uid]

    def saveInfoByUid(self, uid, info):
        self.data["maps"][uid] = info
        self.saveData()

class TaggerFile(object):
    def __init__(self, parent, fname):
        self.parent = parent
        
        with open(path.join(parent.folder, fname), "rb") as fp:
            handle = win32file._get_osfhandle(fp.fileno())
            info = win32file.GetFileInformationByHandle(handle)
            uid = str((info[8]<<32)+info[9])
            self.uid = uid

        self.info = self.parent.getInfoByUid(self.uid) or [0b0, fname, "", ""]
        
    def getTitle(self):
        return self.info[2]

    def displayFileDescription(self):
        dlg = FileDescription({"Title": self.info[2], "Description": self.info[3]})
        dlg.exec()
        if dlg.result():
            self.info[2] = dlg.title
            self.info[3] = dlg.desc
            self.parent.saveInfoByUid(self.uid, self.info)
            return dlg.title
        return None

    def toggleTagunit(self, tagunit):
        # if len(self.uids)>1:
        #     for uid in self.uids:
        #         self.parent.data["maps"][uid][0] |= tagunit
        # else:
        ##No direct access to parent data, self.parent.data["maps"][uid] may not exist
        self.info[0] ^= tagunit
        self.parent.saveInfoByUid(self.uid, self.info)

    def getTag(self):
        tagum = self.info[0]
        return [
            {
                "name": k,
                "tags": [
                    {
                        "name": kk,
                        "checked": (tagum & (0b1<<vv)) == (0b1<<vv),
                        "tagunit": 0b1 << vv
                    } for kk, vv in v.items()
                ]
            } for k, v in self.parent.data["tags"].items()
        ]

class FileDescription(QtGui.QDialog):
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

class FilterManager(QtGui.QDialog):
    def __init__(self, config, logic):
        super().__init__()
        self.setWindowTitle(APP)
        buttonBox = QtGui.QDialogButtonBox(QtGui.QDialogButtonBox.Ok | QtGui.QDialogButtonBox.Cancel)
        buttonBox.accepted.connect(self.accept)
        buttonBox.rejected.connect(self.reject)

        lblLogic = QtGui.QLabel("Logic")
        treeLogic = LogicWidget(config, logic)

        layoutMain = QtGui.QVBoxLayout()
        layoutMain.addWidget(lblLogic)
        layoutMain.addWidget(treeLogic, True)
        layoutMain.addWidget(buttonBox)
        self.setLayout(layoutMain)
        self.treeLogic = treeLogic
        self.logic = logic

    def accept(self):
        self.logic = self.treeLogic.result()
        super().accept()

class LogicWidget(QtGui.QTreeWidget):
    def __init__(self, config, logic, parent=None):
        super().__init__(parent)
        self.setHeaderHidden(True)
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.contextMenu)

        for lg, tagum in logic.items():
            item = QtGui.QTreeWidgetItem(self)
            item.setText(0, lg)
            item.setExpanded(True)
            for k, v in config.items():
                for kk, vv in v.items():
                    if (tagum & (0b1<<vv))==(0b1<<vv):
                        itm = QtGui.QTreeWidgetItem(item)
                        itm.setText(0, kk)
                        itm.setData(0, 33, vv)
        self.config = config

    def contextMenu(self, position):
        config = self.config
        idxes = self.selectedIndexes()
        if len(idxes)>0:
            level = 0
            idx = idxes[0]
            while idx.parent().isValid():
                idx = idx.parent()
                level += 1
            menu = QtGui.QMenu()
            if level==0:
                for k, v in config.items():
                    classMenu = menu.addMenu(k)
                    for kk, vv in v.items():
                        classMenu.addAction(kk).triggered.connect(self.triggerAdd((kk, vv)))
            elif level==1:
                menu.addAction("Delete").triggered.connect(self.triggerRemove)
            menu.exec(self.viewport().mapToGlobal(position))
    def triggerAdd(self, kv):
        itm = self.selectedItems()[0]
        kk, vv = kv
        def _():
            item = QtGui.QTreeWidgetItem(itm)
            item.setText(0, kk)
            item.setData(0, 33, vv)
        return _
    def triggerRemove(self):
        itm = self.selectedItems()[0]
        itm.parent().removeChild(itm)
    def result(self):
        r = {}
        for i in range(self.topLevelItemCount()):
            item = self.topLevelItem(i)
            t = 0b0
            for ii in range(item.childCount()):
                itm = item.child(ii)
                t |= 0b1<<itm.data(0, 33)
            r[item.text(0)] = t
        return r

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
        if not c: return
        if c not in self.config:
            self.config[c] = {}
            self.cmbClass.addItem(c)
        self.txtInput.clear()
        self.txtInput.setFocus()

    def addTag(self, b):
        t = self.txtInput.text()
        if not t: return
        if self.cmbClass.currentIndex() >= 0:
            self.count += 1
            self.config[self.cmbClass.currentText()][t] = self.count
            self.lstTags.addItem(t)
        self.txtInput.clear()
        self.txtInput.setFocus()