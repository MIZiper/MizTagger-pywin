from PyQt4 import QtGui, QtCore
from os import path
import os, json, win32file

from shared import APP, FILE, RESULT
from shared import FileDescription, FilterManager, TagManager

class Window(QtGui.QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP)
        self.createUI()
        self.resize(720, 480)
        self.currentPath = "Browse Folder"

    def createUI(self):
        linkBrowse = QtGui.QPushButton()
        linkBrowse.clicked.connect(self.openFolder)
        listFiles = FileList()

        btnManageTags = QtGui.QPushButton("Manage Tags")
        btnQuickFilter = QtGui.QPushButton("Quick Filter")
        btnComplexFilter = QtGui.QPushButton("Complex Filter")
        btnRefreshFname = QtGui.QPushButton("Refresh File Name")
        btnRefreshUid = QtGui.QPushButton("Refresh UID")

        layoutMain = QtGui.QVBoxLayout()
        layoutAbove = QtGui.QHBoxLayout()
        layoutSide = QtGui.QVBoxLayout()
        layoutSide.addWidget(btnManageTags)
        layoutSide.addWidget(btnQuickFilter)
        layoutSide.addWidget(btnComplexFilter)
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

    def openFolder(self):
        directory = QtGui.QFileDialog.getExistingDirectory(self, APP, QtCore.QDir.currentPath())
        if directory:
            self.currentPath = directory
            self.resizeEvent(None)
            files = [f for f in os.listdir(directory) if path.isfile(path.join(directory, f))]
            # warning: symlink included
            self.listFiles.showFiles(directory ,files)
    
    def resizeEvent(self, event):
        directory = self.currentPath
        metrics = QtGui.QFontMetrics(self.linkBrowse.font())
        elided = metrics.elidedText(directory, QtCore.Qt.ElideMiddle, self.linkBrowse.width()-64)
        self.linkBrowse.setText(elided)

class FileList(QtGui.QTreeWidget):
    def __init__(self):
        super().__init__()
        self.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.customContextMenuRequested.connect(self.contextMenu)
        self.setHeaderLabels(('File Name', 'Title'))
        self.itemDoubleClicked.connect(self.openFile)

    def createUI(self):
        pass

    def contextMenu(self, position):
        pass

    def openFile(self, itemClicked, idx):
        QtGui.QDesktopServices.openUrl(QtCore.QUrl(
            path.join(self.currentPath, itemClicked.text(0))
        ))

    def showFiles(self, directory, files):
        self.clear()
        self.currentPath = directory
        for f in files:
            item = QtGui.QTreeWidgetItem(self)
            item.setText(0, f)

def main():
    import sys
    app = QtGui.QApplication(sys.argv)
    win = Window()
    sys.exit(win.exec())

if __name__ == '__main__':
    main()