from package.dialog import *
from package.list_speed import BrowserHandler



class MyWindow(QtWidgets.QMainWindow):

    def __init__(self, parent=None):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        # use button to invoke slot with another text and color
        self.ui.btnPage_1.clicked.connect(self.addAnotherTextAndColor)
        self.combobox_item = ["Получить скорости сетевых карт у ПК в сети", "Сравнить GUID обьектов в 2 контроллерах домена"]
        self.ui.comboBox.addItems(self.combobox_item)
        self.ui.comboBox.activated.connect(self.action_cmbBox)


    @QtCore.pyqtSlot(int)
    def editProgressBar(self, int):
        self.ui.progressBar.setValue(int)

    @QtCore.pyqtSlot(str)
    def addNewTextAndColor(self, string):
        self.ui.condition.setText(string)

    def addAnotherTextAndColor(self):
        self.ui.condition.setText("START!")
        self.ui.btnPage_1.setDisabled(True)
        # create thread
        self.thread = QtCore.QThread()
        # create object which will be moved to another thread
        # self.ip_addres = "192.168.1.4"
        # self.AD_PASSWORD = 'Vfybgekzwbz3@!'
        # self.AD_USER = 'optima-energy\ininsys'
        # self.AD_SEARCH_TREE = 'DC=optima-energy,DC=ru'
        ipaddress = self.ui.lineEditPage_1.text()
        aduser = self.ui.lineEdit_3Page_1.text()
        adpassord = self.ui.lineEdit_4Page_1.text()
        adsearchtree = self.ui.lineEdit_2Page_1.text()
        filename = self.ui.lineEdit_5Page_1.text()
        self.browserHandler = BrowserHandler(ip_addres=ipaddress, ad_user=aduser, ad_passord=adpassord,
                                             ad_search_tree=adsearchtree, file_name=filename)
        # move object to another thread
        self.browserHandler.moveToThread(self.thread)
        # after that, we can connect signals from this object to slot in GUI thread
        self.browserHandler.newTextAndColor.connect(self.addNewTextAndColor)
        self.browserHandler.progressBarValue.connect(self.editProgressBar)
        # connect started signal to run method of object in another thread
        self.thread.started.connect(self.browserHandler.run)
        self.thread.finished.connect(self.browserHandler_ends)
        # start thread
        self.thread.start()


    def action_cmbBox(self):
        res = self.ui.comboBox.currentText()

        if res == self.combobox_item[0]:
            self.ui.page_1.setVisible(True)
            self.ui.page_2.setVisible(False)
        if res == self.combobox_item[1]:
            self.ui.page_1.setVisible(False)
            self.ui.page_2.setVisible(True)

    def browserHandler_ends(self):
        self.ui.btnPage_1.setDisabled(False)


# class QtWidgets:
#     pass


if __name__ == '__main__':
    import sys
    app = QtWidgets.QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec())