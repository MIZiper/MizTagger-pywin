from PyQt5 import QtWidgets as QtGui, QtCore, QtGui as QtGui5
from os import path
import os

from tagger.shared import APP, TaggerFolder, TaggerFile

class Window(QtGui.QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP)
        self.createUI()
        self.resize(720, 480)
        self.currentPath = "Browse Folder"
        self.setting = QtCore.QSettings(APP, "MIZip.net")
        self.folderObject = None

    def createUI(self):
        linkBrowse = QtGui.QPushButton()
        linkBrowse.clicked.connect(self.openFolder)
        listFiles = FileList()

        btnManageTags = QtGui.QPushButton("Manage Tags")
        btnFilter = QtGui.QPushButton("Filter")
        btnRefreshFname = QtGui.QPushButton("Refresh File Name")
        btnRefreshUid = QtGui.QPushButton("Refresh UID")
        btnCleanFilter = QtGui.QPushButton("Clean Result")

        btnManageTags.clicked.connect(self.manageTagsEvent)
        btnFilter.clicked.connect(self.complexFilterEvent)
        btnFilter.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        btnFilter.customContextMenuRequested.connect(self.quickFilterEvent)
        btnCleanFilter.clicked.connect(self.cleanFilterEvent)

        layoutMain = QtGui.QVBoxLayout()
        layoutAbove = QtGui.QHBoxLayout()
        layoutSide = QtGui.QVBoxLayout()
        layoutSide.addWidget(btnManageTags)
        layoutSide.addWidget(btnFilter)
        layoutSide.addWidget(btnCleanFilter)
        layoutSide.addStretch(1)
        layoutSide.addWidget(btnRefreshFname)
        layoutSide.addWidget(btnRefreshUid)
        layoutAbove.addWidget(listFiles, stretch=1)
        layoutAbove.addLayout(layoutSide)
        layoutMain.addLayout(layoutAbove, stretch=1)
        layoutMain.addWidget(linkBrowse)

        self.setLayout(layoutMain)
        self.linkBrowse = linkBrowse
        self.listFiles = listFiles
        self.btnFilter = btnFilter

    def openFolder(self):
        setting = self.setting
        directory = QtGui.QFileDialog.getExistingDirectory(self, APP, setting.value("RecentDir", QtCore.QDir.currentPath()))
        if directory:
            setting.setValue("RecentDir", directory)
            self.currentPath = directory
            self.resizeEvent(None)
            self.folderObject = TaggerFolder(directory)
            self.listFiles.showFiles(self.folderObject)
    
    def resizeEvent(self, event):
        directory = self.currentPath
        metrics = QtGui5.QFontMetrics(self.linkBrowse.font())
        elided = metrics.elidedText(directory, QtCore.Qt.ElideMiddle, self.linkBrowse.width()-64)
        self.linkBrowse.setText(elided)

    def manageTagsEvent(self):
        if self.folderObject:
            self.folderObject.displayTagManager()

    def quickFilterEvent(self, position):
        if self.folderObject is None: return
        menu = QtGui.QMenu()
        for tagClass in self.folderObject.getTag():
            classAction = QtGui.QAction("- %s -" % tagClass["name"], menu)
            classAction.setDisabled(True)
            menu.addAction(classAction)
            for tagItem in tagClass["tags"]:
                action = QtGui.QAction(tagItem["name"], menu)
                action.setCheckable(True)
                action.setChecked(tagItem["checked"])
                action.triggered.connect(self.triggerQuickFilter(tagItem["tagunit"]))
                menu.addAction(action)
        menu.exec(self.btnFilter.mapToGlobal(position))

    def triggerQuickFilter(self, tagunit):
        def _():
            self.folderObject.toggleFilterTagunit(tagunit)
            self.listFiles.showFiles(self.folderObject, self.folderObject.getFilterResult())
        return _

    def complexFilterEvent(self):
        if self.folderObject is None: return
        if self.folderObject.displayFilterManager():
            self.listFiles.showFiles(self.folderObject, self.folderObject.getFilterResult())

    def cleanFilterEvent(self):
        if self.folderObject is None: return
        self.listFiles.showFiles(self.folderObject)
        self.folderObject.resetLogic()

class FileList(QtGui.QTreeWidget):
    def __init__(self):
        super().__init__()
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.contextMenu)
        self.setHeaderLabels(('File Name', 'Title'))
        self.itemDoubleClicked.connect(self.openFile)
        self.folderObject = None
        self.setColumnWidth(0, 250)

    def contextMenu(self, position):
        idxes = self.selectedIndexes()
        if len(idxes)>0 and self.folderObject is not None:
            idx = idxes[0]
            menu = QtGui.QMenu()
            self.fileObject = TaggerFile(self.folderObject, self.selectedItems()[0].text(0))
            for tagClass in self.fileObject.getTag():
                # classMenu = menu.addMenu(tagClass["name"])
                classAction = QtGui.QAction("- %s -" % tagClass["name"], menu)
                classAction.setDisabled(True)
                menu.addAction(classAction)
                for tagItem in tagClass["tags"]:
                    action = QtGui.QAction(tagItem["name"], menu)
                    action.setCheckable(True)
                    action.setChecked(tagItem["checked"])
                    action.triggered.connect(self.triggerToggle(tagItem["tagunit"]))
                    menu.addAction(action)
            menu.exec(self.viewport().mapToGlobal(position))

    def triggerToggle(self, tagunit):
        def _():
            self.fileObject.toggleTagunit(tagunit)
        return _

    def openFile(self, itemClicked, idx):
        if idx==0:
            QtGui.QDesktopServices.openUrl(QtCore.QUrl(
                path.join(self.currentPath, itemClicked.text(0))
            ))
        elif idx==1:
            self.fileObject = TaggerFile(self.folderObject, itemClicked.text(0))
            title = self.fileObject.displayFileDescription()
            if title:
                itemClicked.setText(1, title)

    def showFiles(self, folderObject, fnames=None):
        self.clear()
        directory = folderObject.folder
        self.currentPath = directory
        self.folderObject = folderObject
        if fnames is None:
            fnames = [f for f in os.listdir(directory) if path.isfile(path.join(directory, f))]
        for f in fnames:
            # warning: symlink included
            item = QtGui.QTreeWidgetItem(self)
            fileObject = TaggerFile(folderObject, f)
            item.setText(0, f)
            item.setText(1, fileObject.getTitle())

def main():
    import sys
    app = QtGui.QApplication(sys.argv)
    win = Window()
    sys.exit(win.exec())

if __name__ == '__main__':
    main()