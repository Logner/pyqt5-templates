# PYQT5 template for easy spreadsheet manipulation

ctr = 0
Col_labels = ['Name','Date', 'Code', 'Quantity', 'Total Price', 'Comments']

from PyQt5 import QtCore, QtGui, QtWidgets
from datetime import datetime
import csv

#custom
def get_table_items():
    global ctr
    lst = ['item %d' % ctr,
            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "AA%s" % ctr,
            str(ctr * 2 + 1),
            str(ctr + 10),
            "comment %d" % ctr, ctr]
    return lst

class Ui_DockWidget(object):
    def setupUi(self, DockWidget):
        DockWidget.setObjectName("DockWidget")
        DockWidget.resize(548, 444)
        self.dockWidgetContents = QtWidgets.QWidget()
        self.dockWidgetContents.setObjectName("dockWidgetContents")
        self.Update = QtWidgets.QPushButton(self.dockWidgetContents)
        self.Update.setGeometry(QtCore.QRect(110, 390, 89, 25))
        self.Update.setObjectName("Update")
        self.Delete = QtWidgets.QPushButton(self.dockWidgetContents)
        self.Delete.setGeometry(QtCore.QRect(210, 390, 89, 25))
        self.Delete.setObjectName("Delete")
        self.Add = QtWidgets.QPushButton(self.dockWidgetContents)
        self.Add.setGeometry(QtCore.QRect(10, 390, 89, 25))
        self.Add.setObjectName("Add")
        self.tableWidget = QtWidgets.QTableWidget(self.dockWidgetContents)
        self.tableWidget.setGeometry(QtCore.QRect(0, 0, 541, 381))
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(len(Col_labels))
        self.tableWidget.setHorizontalHeaderLabels(Col_labels)
        self.tableWidget.setRowCount(100)
        DockWidget.setWidget(self.dockWidgetContents)

        self.retranslateUi(DockWidget)
        QtCore.QMetaObject.connectSlotsByName(DockWidget)

    #connectors
        self.Add.clicked.connect(self.AddRow)
        self.Update.clicked.connect(self.updateCSV)
        self.Delete.clicked.connect(self.delRow)
                
    def retranslateUi(self, DockWidget):
        _translate = QtCore.QCoreApplication.translate
        DockWidget.setWindowTitle(_translate("DockWidget", "DockWidget"))
        self.Update.setText(_translate("DockWidget", "Update"))
        self.Delete.setText(_translate("DockWidget", "Delete"))
        self.Add.setText(_translate("DockWidget", "Add"))

    #custom
    def AddRow(self):
        global ctr
        lst = get_table_items()
        row = ctr
        _ti = QtWidgets.QTableWidgetItem
        for i in range(0, len(Col_labels)):
            self.tableWidget.setItem(int(row), i, _ti(lst[i]))
        ctr += 1
        print('row %s added' % ctr)

    def updateCSV(self):
        wr = open('test.csv', 'w', newline='')
        w = csv.writer(wr, delimiter=',', quotechar="|", quoting=csv.QUOTE_MINIMAL)
        row = ctr
        lst = []
        w.writerow(Col_labels)

        for i in range(0, row):
            for j in range(0, len(Col_labels)):
                lst += [self.tableWidget.item(i, j).text()]
                if j == len(Col_labels)-1:
                    w.writerow(lst)
                    lst = []
                else:
                    pass
        
        wr.close()

    def delRow(self):
        global ctr
        _ti = QtWidgets.QTableWidgetItem
        curItem = self.tableWidget.currentItem()
        row = curItem.row()
        code_col = 2
        for i in range(0, len(Col_labels)):
            self.tableWidget.setItem(row, i, _ti(""))
        past = ctr
        ctr = row
        for i in range(row, past):       
            for j in range(0, len(Col_labels)):
                if i == past-1:
                    self.tableWidget.setItem(i, j, _ti(""))
                else:
                    self.tableWidget.setItem(i, j, _ti(self.tableWidget.item(i+1, j).text()))
        if self.tableWidget.item(row, code_col).text() != "":
            ctr = past-1
        else:
            pass

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    DockWidget = QtWidgets.QDockWidget()
    ui = Ui_DockWidget()
    ui.setupUi(DockWidget)
    DockWidget.show()
    sys.exit(app.exec_())


