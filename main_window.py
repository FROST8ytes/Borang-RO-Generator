# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'main_window.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.


from PyQt5 import QtCore, QtGui, QtWidgets
import sys


class MainWindow(object):
    def __init__(self, main_window):
        self.central_widget = None
        self.grid_layout_widget = None
        self.grid_layout = None
        self.button_choose_excel = None
        self.text_excel_path = None
        self.button_start_stop = None
        self.text_template_path = None
        self.button_choose_template = None
        self.logo = None
        self.generate_progress = None
        self.list_view_sheets = None
        self.menubar = None
        self.menu_menu = None
        self.menu_help = None
        self.statusbar = None
        self.action_reset = None
        self.action_import = None
        self.action_restart = None
        self.action_quit = None
        self.action_documentation = None
        self.action_credits = None
        self.action_save = None

        self.setup_ui(main_window)

    def setup_ui(self, main_window):
        main_window.setObjectName("Borang RO Generator")
        main_window.setFixedSize(800, 600)
        main_window.setLocale(QtCore.QLocale(QtCore.QLocale.Malay, QtCore.QLocale.Malaysia))
        self.central_widget = QtWidgets.QWidget(main_window)
        self.central_widget.setObjectName("centralwidget")
        self.grid_layout_widget = QtWidgets.QWidget(self.central_widget)
        self.grid_layout_widget.setGeometry(QtCore.QRect(10, 0, 771, 551))
        self.grid_layout_widget.setObjectName("gridLayoutWidget")
        self.grid_layout = QtWidgets.QGridLayout(self.grid_layout_widget)
        self.grid_layout.setContentsMargins(0, 0, 0, 0)
        self.grid_layout.setSpacing(10)
        self.grid_layout.setObjectName("grid_layout")

        self.button_choose_excel = QtWidgets.QPushButton(self.grid_layout_widget)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        size_policy.setHorizontalStretch(0)
        size_policy.setVerticalStretch(0)
        size_policy.setHeightForWidth(self.button_choose_excel.sizePolicy().hasHeightForWidth())
        self.button_choose_excel.setSizePolicy(size_policy)
        self.button_choose_excel.setObjectName("button_choose_excel")
        self.button_choose_excel.clicked \
            .connect(lambda: MainWindow.open_file_name_dialog("Microsoft Excel Worksheet (*.xlsx)"))
        self.grid_layout.addWidget(self.button_choose_excel, 2, 2, 2, 1)

        self.text_excel_path = QtWidgets.QPlainTextEdit(self.grid_layout_widget)
        self.text_excel_path.setObjectName("text_excel_path")
        self.text_excel_path.setReadOnly(True)
        self.grid_layout.addWidget(self.text_excel_path, 2, 0, 1, 2)

        self.button_start_stop = QtWidgets.QPushButton(self.grid_layout_widget)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        size_policy.setHorizontalStretch(0)
        size_policy.setVerticalStretch(0)
        size_policy.setHeightForWidth(self.button_start_stop.sizePolicy().hasHeightForWidth())
        self.button_start_stop.setSizePolicy(size_policy)
        self.button_start_stop.setObjectName("button_start_stop")
        self.grid_layout.addWidget(self.button_start_stop, 5, 2, 1, 1)

        self.text_template_path = QtWidgets.QPlainTextEdit(self.grid_layout_widget)
        self.text_template_path.setObjectName("text_template_path")
        self.text_template_path.setReadOnly(True)
        self.grid_layout.addWidget(self.text_template_path, 1, 0, 1, 2)

        self.button_choose_template = QtWidgets.QPushButton(self.grid_layout_widget)
        size_policy = QtWidgets.QSizePolicy(QtWidgets.QSizePolicy.Fixed, QtWidgets.QSizePolicy.Preferred)
        size_policy.setHorizontalStretch(0)
        size_policy.setVerticalStretch(0)
        size_policy.setHeightForWidth(self.button_choose_template.sizePolicy().hasHeightForWidth())
        self.button_choose_template.setSizePolicy(size_policy)
        self.button_choose_template.setLayoutDirection(QtCore.Qt.LeftToRight)
        self.button_choose_template.setObjectName("button_choose_template")
        self.button_choose_template.clicked \
            .connect(lambda: MainWindow.open_file_name_dialog("Microsoft Word Document (*.docx)"))
        self.grid_layout.addWidget(self.button_choose_template, 1, 2, 1, 1)

        self.logo = QtWidgets.QGraphicsView(self.grid_layout_widget)
        self.logo.setObjectName("logo")
        self.grid_layout.addWidget(self.logo, 0, 0, 1, 3)

        self.generate_progress = QtWidgets.QProgressBar(self.grid_layout_widget)
        self.generate_progress.setProperty("value", None)
        self.generate_progress.setAlignment(QtCore.Qt.AlignBottom | QtCore.Qt.AlignHCenter)
        self.generate_progress.setObjectName("generate_progress")
        self.grid_layout.addWidget(self.generate_progress, 5, 0, 1, 2)

        self.list_view_sheets = QtWidgets.QListView(self.grid_layout_widget)
        self.list_view_sheets.setProperty("showDropIndicator", False)
        self.list_view_sheets.setObjectName("view_sheets")
        self.grid_layout.addWidget(self.list_view_sheets, 3, 0, 1, 2)

        main_window.setCentralWidget(self.central_widget)
        self.menubar = QtWidgets.QMenuBar(main_window)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 792, 21))
        self.menubar.setObjectName("menubar")
        self.menu_menu = QtWidgets.QMenu(self.menubar)
        self.menu_menu.setObjectName("menu_menu")
        self.menu_help = QtWidgets.QMenu(self.menubar)
        self.menu_help.setObjectName("menu_help")
        main_window.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(main_window)
        self.statusbar.setObjectName("statusbar")
        main_window.setStatusBar(self.statusbar)
        self.action_reset = QtWidgets.QAction(main_window)
        self.action_reset.setObjectName("action_reset")
        self.action_import = QtWidgets.QAction(main_window)
        self.action_import.setObjectName("action_import")
        self.action_restart = QtWidgets.QAction(main_window)
        self.action_restart.setObjectName("action_restart")
        self.action_quit = QtWidgets.QAction(main_window)
        self.action_quit.setObjectName("action_quit")
        self.action_quit.triggered.connect(main_window.close)
        self.action_documentation = QtWidgets.QAction(main_window)
        self.action_documentation.setObjectName("action_documentation")
        self.action_credits = QtWidgets.QAction(main_window)
        self.action_credits.setObjectName("action_credits")
        self.action_save = QtWidgets.QAction(main_window)
        self.action_save.setObjectName("action_save")
        self.menu_menu.addAction(self.action_import)
        self.menu_menu.addAction(self.action_save)
        self.menu_menu.addAction(self.action_reset)
        self.menu_menu.addSeparator()
        self.menu_menu.addAction(self.action_restart)
        self.menu_menu.addAction(self.action_quit)
        self.menu_help.addAction(self.action_documentation)
        self.menu_help.addAction(self.action_credits)
        self.menubar.addAction(self.menu_menu.menuAction())
        self.menubar.addAction(self.menu_help.menuAction())

        self.retranslate_ui(main_window)
        QtCore.QMetaObject.connectSlotsByName(main_window)

    def retranslate_ui(self, main_window):
        _translate = QtCore.QCoreApplication.translate
        main_window.setWindowTitle(_translate("MainWindow", "Cipta Borang RO"))
        self.button_choose_excel.setText(_translate("MainWindow", "Pilih Excel"))
        self.button_start_stop.setText(_translate("MainWindow", "Cipta Borang"))
        self.button_choose_template.setText(_translate("MainWindow", "Pilih Templat"))
        self.menu_menu.setTitle(_translate("MainWindow", "Menu"))
        self.menu_help.setTitle(_translate("MainWindow", "Bantu"))
        self.action_reset.setText(_translate("MainWindow", "Mula Semula"))
        self.action_import.setText(_translate("MainWindow", "Import Pilihan"))
        self.action_restart.setText(_translate("MainWindow", "Lancar Semula"))
        self.action_quit.setText(_translate("MainWindow", "Keluar"))
        self.action_documentation.setText(_translate("MainWindow", "Dokumentasi"))
        self.action_credits.setText(_translate("MainWindow", "Kredit"))
        self.action_save.setText(_translate("MainWindow", "Simpan Pilihan"))

    @staticmethod
    def open_file_name_dialog(filetype):
        options = QtWidgets.QFileDialog.Options()
        options |= QtWidgets.QFileDialog.DontUseNativeDialog
        file_name, _ = QtWidgets.QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "",
                                                             filetype, options=options)
        if file_name:
            print(file_name)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    app_main_window = QtWidgets.QMainWindow()
    ui = MainWindow(app_main_window)
    app_main_window.show()
    sys.exit(app.exec_())
