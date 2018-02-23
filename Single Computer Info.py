import pythoncom
import os
import queue
import sys
import win32com.client, pywintypes
import subprocess
import datetime
import re
from shutil import copyfile
from winreg import *
import operator
import ctypes
import argparse
import socket

from PyQt5.QtWidgets import QApplication, QMainWindow, QWidget, QAction, QLabel, QBoxLayout, QVBoxLayout, QHBoxLayout, QLineEdit, QPlainTextEdit, QPushButton, QProgressBar, QTabWidget, QFileDialog, QMessageBox, QScrollArea, QStatusBar, QDialog, QTableWidget, QTableWidgetItem, QSplitter, QSizePolicy, QFrame, QGraphicsOpacityEffect, QLayout,QShortcut, QToolButton, QActionGroup,QMenu, QStyleFactory
from PyQt5.QtGui import QFont, QBrush, QColor, QMovie,QKeySequence, QClipboard, QCursor, QIcon, QPalette
from PyQt5.QtCore import *

from ComputerInfoSharedResources.CIStorage import Program, ThreadSafeCounter
from ComputerInfoSharedResources.CIWMI import ComputerInfo, WMIThread
from ComputerInfoSharedResources.CISettings import Settings
from ComputerInfoSharedResources.CIForms import AuthenticationForm, FileForm
from ComputerInfoSharedResources.CIColor import generate_color,generate_text,rgb_to_hex
from ComputerInfoSharedResources.CITime import format_date
from ComputerInfoSharedResources.CICustomWidgets import LinkLabel, CustomDialog, DialogTable, CustomDataLabel, CustomLineEdit
from ComputerInfoSharedResources.dynamic_forms.forms import DynamicForm
from ComputerInfoSharedResources.dynamic_forms.models import DynamicModel
from ComputerInfoSharedResources.CITime import format_date

global_blue = "#3366cc"
global_green = "#00cc00"
global_red = "#cc3333"

"""Workers to be moved to threads"""
class GenericWMIWorker(QObject):
    finished = pyqtSignal()

    def __init__(self,callback,parent=None,**kwargs):
        self.callback = callback
        self.kwargs = kwargs
        super().__init__(parent)

    @pyqtSlot()
    def startWork(self):
        pythoncom.CoInitialize()
        self.callback()
        pythoncom.CoUninitialize()
        self.finished.emit()

class GenericErroringWorker(GenericWMIWorker):
    finished = pyqtSignal(str)

    @pyqtSlot()
    def startWork(self):
        pythoncom.CoInitialize()
        retval = self.callback()
        pythoncom.CoUninitialize()
        self.finished.emit(str(retval))


"""
Main window for getting info from single computers
"""
class SinglePCApp(QMainWindow):
    def __init__(self,parent=None,debug=False,main_wind=None,main_path="",start_name=""):
        self.main_wind = main_wind
        self.main_path = main_path
        self.start_name = start_name
        self.manual_user = None
        self.manual_pass = None
        self.prog_counter = ThreadSafeCounter()
        self.settings = DynamicModel("comp_info.cfg",os.getenv("APPDATA") + '\\Single Computer Info\\comp_info.cfg')
        self.hosts_list = {}
        self.debug = debug
        self.link_color = "blue"
        self.recent_computers = []

        super().__init__(parent)

        self.get_recent_computers()
        self.createWidgets()

        if self.start_name:
            self.submit_btn.click()

    def get_recent_computers(self):
        if os.path.exists(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg'):
            with open(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg','r') as f:
                for line in f:
                    self.recent_computers.append(line.strip())
        else:
            open(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg','w').close()

    def write_recent_computers(self):
        temp_list = self.recent_computers
        if os.path.exists(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg'):
            with open(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg','r') as f:
                for line in f:
                    if not line.upper().strip() in (temp_comp.upper().strip() for temp_comp in temp_list) and line.strip():
                        temp_list.append(line.upper().strip())

        with open(os.getenv("APPDATA") + '\\Single Computer Info\\history.cfg','w') as f:
            for t in temp_list[:50]:
                f.write(t.upper().strip() + "\n")


    def createWidgets(self):
        """Creates initial widgets for main window and adds keyboard shortcuts"""
        self.setWindowTitle("Single Computer Info")
        self.mainmenu = self.menuBar()
        self.filemenu = self.mainmenu.addMenu('File')
        self.optionsmenu = self.mainmenu.addMenu('Options')
        self.helpmenu = self.mainmenu.addMenu('Help')

        self.exit_button = QAction('Exit',self)
        self.exit_button.setShortcut('Ctrl+Q')
        self.exit_button.triggered.connect(self.close)

        self.filemenu.addAction(self.exit_button)

        other_credentials = QAction("Use Other Credentials",self)
        other_credentials.triggered.connect(self.alt_user_popup)
        self.optionsmenu.addAction(other_credentials)
        self.optionsmenu.addSeparator()
        adv_options = QAction('Advanced Options',self)
        adv_options.triggered.connect(self.display_settings)
        self.optionsmenu.addAction(adv_options)

        about = QAction('About',self)
        about.triggered.connect(lambda:QMessageBox.information(self,"About", "Single Computer Info\nVersion 2.0"))
        self.helpmenu.addAction(about)

        self.containerWidget = QWidget()
        self.setCentralWidget(self.containerWidget)

        self.title_label = QLabel("Input Computer Name or IP Below",self.containerWidget)
        self.title_label.setAlignment(Qt.AlignCenter)
        font = self.title_label.font()
        font.setPointSize(16)
        self.title_label.setFont(font)
        self.subtitle = QLabel("(Connecting as " + str(self.manual_user) + ")",self.containerWidget)
        self.subtitle.setAlignment(Qt.AlignCenter)
        if not self.manual_user:
            self.subtitle.hide()
        self.inbox = CustomLineEdit(self.containerWidget,scroll_list=self.recent_computers)
        self.inbox.setPlaceholderText("Type Computer Name Here")
        self.inbox.setClearButtonEnabled(True)
        self.inbox.returnPressed.connect(self.get_computer_names)
        self.fill_input(self.start_name)

        self.inbox_focus = QShortcut(QKeySequence("Ctrl+L"),self,self.select_input)

        self.submit_btn = QPushButton("Get Info",self.containerWidget)
        self.submit_btn.clicked.connect(self.get_computer_names)

        self.clear_btn = QPushButton("Clear",self.containerWidget)

        self.clear_btn.clicked.connect(self.clear_outbox)
        self.vertbox = QScrollArea(self.containerWidget)
        self.vertbox.setWidgetResizable(True)
        self.vertbox.setStyleSheet("QScrollArea{Border:0;}")

        self.interior_widget = QWidget()
        self.interior_widget_layout = QVBoxLayout()
        self.interior_widget_layout.setDirection(QBoxLayout.BottomToTop)
        self.interior_widget.setLayout(self.interior_widget_layout)
        self.interior_widget_layout.setAlignment(Qt.AlignTop)
        self.interior_widget_layout.setContentsMargins(0,0,0,0)
        self.vertbox.setWidget(self.interior_widget)
        self.outboxes = []

        self.overlab = QLabel(" ",self.containerWidget)

        self.copybannerlabel = QStatusBar(self.containerWidget)
        self.copybannerlabel.showMessage("")
        self.copybannerlabel.setSizePolicy(QSizePolicy.Expanding,QSizePolicy.Fixed)

        self.layout = QVBoxLayout()
        self.marginlayout = QVBoxLayout()

        self.input_layout = QHBoxLayout()
        self.input_layout.addWidget(self.inbox)
        self.input_layout.addWidget(self.submit_btn)
        self.input_layout.addWidget(self.clear_btn)
        self.marginlayout.addWidget(self.title_label)
        self.marginlayout.addWidget(self.subtitle)
        self.marginlayout.addLayout(self.input_layout)
        self.marginlayout.addWidget(self.vertbox)
        self.layout.addLayout(self.marginlayout)
        self.layout.addWidget(self.copybannerlabel)
        self.marginlayout.setContentsMargins(10,0,10,0)
        self.layout.setContentsMargins(0,0,0,0)
        self.containerWidget.setLayout(self.layout)
        self.setMinimumHeight(400)
        self.setMinimumWidth(725)
        self.show()

        self.devices_shortcut = QShortcut(QKeySequence("Ctrl+1"),self)
        self.drives_shortcut = QShortcut(QKeySequence("Ctrl+2"),self)
        self.printers_shortcut = QShortcut(QKeySequence("Ctrl+3"),self)
        self.programs_shortcut = QShortcut(QKeySequence("Ctrl+4"),self)
        self.remotecmd_shortcut = QShortcut(QKeySequence("Ctrl+R"),self)


    def get_computer_names(self):
        self.submit_btn.setEnabled(False)
        computer_name = self.inbox.text().strip()
        if computer_name.upper().strip() in (recent_comp.upper().strip() for recent_comp in self.recent_computers):
            self.recent_computers.remove(computer_name.upper().strip())

        self.recent_computers = [computer_name.upper().strip()] + self.recent_computers
        self.write_recent_computers()
        self.inbox.update_list(self.recent_computers)
        if computer_name:
            if self.manual_user and self.manual_pass:
                c = ComputerInfo(input_name=computer_name,count=self.prog_counter.get(),manual_user=self.manual_user, manual_pass=self.manual_pass,debug=self.debug)
            else:
                c = ComputerInfo(input_name=computer_name,count=self.prog_counter.get(),debug=self.debug)
            self.prog_counter.increment()
            self.outboxes.append(OutputComputer(parent=self.interior_widget,comp_obj=c,clipboard_callback=self.to_clipboard,all_to_clipboard_callback=self.all_to_clipboard,settings=self.settings,main_path=self.main_path,link_color=self.link_color))
            self.interior_widget_layout.addWidget(self.outboxes[-1])
        else:
            QMessageBox.information(self,"No Computer name","Please specify a computer to connect to")

        try:
            self.devices_shortcut.disconnect()
        except:pass
        finally:
            self.devices_shortcut.activated.connect(lambda:self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().get_devices_action.activate(QAction.Trigger))

        try: self.drives_shortcut.disconnect()
        except:pass
        finally:
            self.drives_shortcut.activated.connect(lambda:self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().get_drives_action.activate(QAction.Trigger))

        try: self.printers_shortcut.disconnect()
        except:pass
        finally:
            self.printers_shortcut.activated.connect(lambda:self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().get_printers_action.activate(QAction.Trigger))

        try: self.programs_shortcut.disconnect()
        except:pass
        finally:
            self.programs_shortcut.activated.connect(lambda:self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().get_programs_action.activate(QAction.Trigger))

        try: self.remotecmd_shortcut.disconnect()
        except:pass
        finally:
            self.remotecmd_shortcut.activated.connect(self.active_run_cmd)

        self.submit_btn.setEnabled(True)

    def active_run_cmd(self):
        if not self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().comp_obj.local:
            self.interior_widget_layout.itemAt(self.interior_widget_layout.count()-1).widget().remote_cmd()

    def fill_input(self,name=""):
        input_name = socket.gethostname()
        if name:
            self.inbox.setText(name)
        else:
            self.inbox.setText(input_name)
        self.inbox.selectAll()
        self.inbox.setFocus()

    def select_input(self):
        self.inbox.selectAll()
        self.inbox.setFocus()

    """
    Creates window for editing settings.
    """
    def display_settings(self):
        top = QDialog(self)
        top.setWindowTitle("Settings")
        top.setSizeGripEnabled(True)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)

        settings_form = DynamicForm(top,title="",submit_callback=top.accept,submit_callback_kwargs={},dynamicmodel=self.settings)
        top_layout.addWidget(settings_form)
        top_layout.setAlignment(settings_form,Qt.AlignTop)
        top.show()
        top.activateWindow()

    def clear_outbox(self):
        while self.interior_widget_layout.count():
            child = self.interior_widget_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self.fill_input()

    def alt_user_popup(self):
        default_username = None
        if self.manual_user:
            default_username = self.manual_user
        else:
            try:
                default_username = self.settings.settings_dict['default username']
            except: pass

        def save_credentials(parent,username,password):
            self.update_credentials(username,password)
            parent.hide()

        top = QDialog(self)
        top.setWindowTitle("Authenticate as different user")
        top.setSizeGripEnabled(True)
        #top.setWindowModality(Qt.ApplicationModal)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        auth_form = AuthenticationForm(top,title="Admin Credentials",all_btns=True,save_callback=lambda:save_credentials(top,auth_form.username,auth_form.password))
        if default_username:
            auth_form.usernamefield.setText(default_username)
            auth_form.passwordfield.setFocus()
        top_layout.addWidget(auth_form)
        top_layout.setAlignment(auth_form,Qt.AlignTop)
        top.show()
        top.activateWindow()

    def update_credentials(self,username,password):
        self.manual_user = username
        self.manual_pass = password
        if username and password:
            self.subtitle.setText("(Connecting as " + str(username) + ")")
            self.subtitle.show()
        else:
            self.subtitle.hide()

    def to_clipboard(self,event=None,textdata=""):
        clip = QApplication.clipboard()
        clip.setText(textdata)
        self.copybannerlabel.show()

        self.copybannerlabel.setAutoFillBackground(True)
        self.copybannerlabel.setPalette(QPalette(QColor(global_blue)))
        self.copybannerlabel.setStyleSheet("color:white")
        self.copybannerlabel.showMessage("Copied '{0}' to clipboard".format(textdata))

    def all_to_clipboard(self,comp_obj=None):
        clip = QApplication.clipboard()
        clip_data = ""
        if comp_obj.input_name:clip_data += "Name: " + comp_obj.input_name.upper()
        if comp_obj.serial:clip_data += "\nSerial: " + comp_obj.serial.upper()
        if comp_obj.model:clip_data += "\nModel: " + comp_obj.model
        if comp_obj.user: clip_data += "\nUser: " + comp_obj.user
        if comp_obj.os: clip_data += "\nOS: " + comp_obj.os
        try:
            if comp_obj.ip_addresses[0]: clip_data += "\nIP: " + comp_obj.ip_addresses[0]
        except:pass
        clip.setText(clip_data)
        self.copybannerlabel.setAutoFillBackground(True)
        self.copybannerlabel.setPalette(QPalette(QColor(global_blue)))
        self.copybannerlabel.setStyleSheet("color:white")
        self.copybannerlabel.showMessage("Copied '{0}' specs to clipboard".format(comp_obj.input_name))
        self.copybannerlabel.show()

"""
Storage and display of a single computer within SinglePCApp
"""
class OutputComputer(QFrame):
    sig = pyqtSignal()
    def __init__(self,parent=None,comp_obj=None,destroy_callback=None,clipboard_callback=None,all_to_clipboard_callback=None,settings=None,extra_data=None,main_path="",link_color="blue"):
        #tk.Frame.__init__(self,master,bd=2, relief=GROOVE)
        super().__init__(parent)
        self.main_path = main_path
        self.parent = parent
        self.link_color = link_color
        self.loading_queue = ThreadSafeCounter()
        #self.setStyleSheet("border:1px solid black")
        self.setFrameShape(QFrame.StyledPanel)
        self.setLineWidth(1)
        self.settings = settings
        self.extra_data = extra_data
        #self.show = False
        self.comp_obj=comp_obj
        self.processentry = []

        self.destroy_callback = destroy_callback
        self.clipboard_callback = clipboard_callback
        self.all_to_clipboard_callback = all_to_clipboard_callback

        self.layout = QHBoxLayout()
        self.layout.setContentsMargins(0,0,0,0)
        self.setLayout(self.layout)
        self.setSizePolicy(QSizePolicy.MinimumExpanding,QSizePolicy.Maximum)
        self.placeholder = QLabel()
        self.placeholder.setMaximumHeight(50)
        self.placeholder.setAlignment(Qt.AlignCenter)

        if self.main_path:
            self.gif = QMovie(self.main_path + "\\loading2.gif")
        else:
            self.gif = QMovie("loading2.gif")
        self.placeholder.setMovie(self.gif)
        self.placeholder.setAttribute(Qt.WA_TranslucentBackground,True)
        self.gif.start()
        self.layout.addWidget(self.placeholder)
        self.show()

        self.install_status = LinkLabel("",self,None)

        self.outputworker = GenericWMIWorker(self.comp_obj.get_info)
        self.outputthread = QThread()
        self.outputworker.moveToThread(self.outputthread)
        self.outputthread.start()
        self.outputworker.finished.connect(self.dataWidgets)

        self.progworker = GenericWMIWorker(self.comp_obj.get_specific_program)
        self.progthread = QThread()
        self.progworker.moveToThread(self.progthread)
        self.progthread.start()
        self.progworker.finished.connect(self.program_window)

        self.patchworker = GenericWMIWorker(self.comp_obj.get_patches)
        self.patchthread = QThread()
        self.patchworker.moveToThread(self.patchthread)
        self.patchthread.start()
        self.patchworker.finished.connect(self.patch_window)

        self.printworker = GenericWMIWorker(self.comp_obj.get_printers)
        self.printthread = QThread()
        self.printworker.moveToThread(self.printthread)
        self.printthread.start()
        self.printworker.finished.connect(self.printer_window)

        self.deviceworker = GenericWMIWorker(self.comp_obj.get_devices)
        self.devicethread = QThread()
        self.deviceworker.moveToThread(self.devicethread)
        self.devicethread.start()
        self.deviceworker.finished.connect(self.devices_window)

        self.driveworker = GenericWMIWorker(self.comp_obj.get_disks)
        self.drivethread = QThread()
        self.driveworker.moveToThread(self.drivethread)
        self.drivethread.start()
        self.driveworker.finished.connect(self.drives_window)

        self.adminworker = GenericErroringWorker(self.add_admin)
        self.adminthread = QThread()
        self.adminworker.moveToThread(self.adminthread)
        self.adminthread.start()
        self.adminworker.finished.connect(self.admin_complete)

        self.installworker = GenericErroringWorker(self.comp_obj.manual_run_script)
        self.installthread = QThread()
        self.installworker.moveToThread(self.installthread)
        self.installthread.start()
        self.installworker.finished.connect(self.install_complete)

        self.damewareworker = GenericWMIWorker(self.comp_obj.start_service)
        self.damewarethread = QThread()
        self.damewareworker.moveToThread(self.damewarethread)
        self.damewarethread.start()
        self.damewareworker.finished.connect(self.dameware)


        self.sig.connect(self.outputworker.startWork)
        self.sig.emit()

    def all_to_clipboard(self):
        self.all_to_clipboard_callback(comp_obj=self.comp_obj)

    @pyqtSlot()
    def printer_window(self):
        def on_click_ip(item):
            if re.search('[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+',item.text()):
                try:
                    substr = re.search('[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+',item.text()).group(0)
                    os.startfile("http://"+substr)
                except Exception as e:
                    print(e)

        printer_list = self.comp_obj.printer_queue.get()
        printer_list.sort(key=operator.attrgetter("printer"),reverse=False)

        top = CustomDialog(self)
        top.setWindowTitle("Printers on %s" % self.comp_obj.input_name)

        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        table = DialogTable(top,['Printer','Port'])
        table.itemDoubleClicked.connect(on_click_ip)
        table.setRowCount(len(printer_list))
        top_layout.addWidget(table)

        for i,p in enumerate(printer_list):
            print_item = QTableWidgetItem(p.printer)
            print_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            table.setItem(i,0,print_item)
            port_item = QTableWidgetItem(p.port)
            port_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            if re.search('[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+',p.port):
                port_item.setForeground(QBrush(QColor(self.link_color)))
            table.setItem(i,1,port_item)
        table.resizeColumnsToContents()

        top.show()
        top.activateWindow()

        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

    def remote_cmd(self):
        res_path = ""
        if hasattr(sys, 'frozen'):
            res_path = os.path.realpath(sys.executable)
        else:
            res_path = os.path.realpath(sys.argv[0])

        os.system("start cmd /k \"" + os.path.dirname(res_path) + "\\psexec.exe\" \\\\" + self.comp_obj.input_name + " cmd")

    @pyqtSlot()
    def dameware(self):
        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

        for dame in self.settings.settings_dict['dameware']:
            try:
                subprocess.Popen([dame,"-h","-m:"+self.comp_obj.input_name,"-x"])
                break
            except: pass

    def gpedit(self):
        subprocess.Popen(["C:\windows\system32\gpedit.msc","/gpcomputer:",self.comp_obj.input_name],shell=True)

    def compmgmt(self):
        subprocess.Popen(["C:\windows\system32\compmgmt.msc", "/computer:\\\\"+self.comp_obj.input_name],shell=True)

    @pyqtSlot(str)
    def admin_complete(self,retval):
        if retval:
            self.msgboxsubmit.setEnabled(True)
            QMessageBox.critical(self,"Error", "Could not add group\n"+str(retval))

        else:
            QMessageBox.critical(self,"Complete!","'%s' added with no errors" % self.settings.settings_dict['group to add to admin'])
            self.msgboxwindow.hide()

    def add_admin(self):
        self.userName = self.localframe.usernamefield.text()
        self.passwd = self.localframe.passwordfield.text()
        self.domainUserName = self.domainframe.usernamefield.text()
        self.domainPasswd = self.domainframe.passwordfield.text()

        pythoncom.CoInitialize()
        NS = win32com.client.Dispatch('ADSNameSpaces')
        retval = "Unknown Error"
        if 'group to add to admin' in self.settings.settings_dict and 'domain' in self.settings.settings_dict:
            if self.userName and self.passwd:
                dso = NS.getobject('','WinNT:')
                try:
                    group = dso.OpenDSObject('WinNT://%s/%s/Administrators,group' % (self.settings.settings_dict['domain'], self.comp_obj.input_name) , self.comp_obj.input_name + "\\" + self.userName, self.passwd, 1)
                    if self.domainUserName and self.domainPasswd:
                        user = dso.OpenDSObject('WinNT://%s/%s,group' % (self.settings.settings_dict['domain'],self.settings.settings_dict['group to add to admin']), self.domainUserName, self.domainPasswd, 1)
                    else:
                        user = NS.getobject('','WinNT://%s/%s,group' % (self.settings.settings_dict['domain'],self.settings.settings_dict['group to add to admin']))
                    if not group.IsMember(user.ADsPath):
                        group.Add(user.ADsPath)
                        retval = None
                    else:
                        retval = "Already in the group"
                except pywintypes.com_error as e:
                    print(e)
                    retval = e
            else:
                retval = "Invalid Credentials"
        else:
            retval = "Invalid Group/Domain"
        pythoncom.CoUninitialize()
        return retval

    @pyqtSlot()
    def program_window(self):

        programs_list = self.comp_obj.programs_queue.get()
        programs_list.sort(key=operator.attrgetter("name"),reverse=False)
        top = CustomDialog(self)
        top.setWindowTitle("Installed Programs on %s" % self.comp_obj.input_name)

        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        searchbox = QLineEdit()

        searchbox.setPlaceholderText("Search Programs")
        searchbox.setClearButtonEnabled(True)

        table = DialogTable(top,['Name','Date','Version'])

        top_layout.addWidget(searchbox)
        top_layout.addWidget(table)
        #top_layout.setAlignment(table,Qt.AlignTop)

        def find_text(search_term,table,programs_list):
            table.clearContents()
            table.setRowCount(0)
            table.setRowCount(len(programs_list))
            i=0
            for p in programs_list:
                if not search_term or search_term.lower().strip() in p.name.lower().strip():
                    table.setRowCount(i+1)
                    name_item = QTableWidgetItem(p.name)
                    name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,0,name_item)

                    version_item = QTableWidgetItem(p.version)
                    version_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,1,version_item)

                    date_item = QTableWidgetItem(format_date(p.date))
                    date_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,2,date_item)
                    i+=1
        find_text("",table,programs_list)
        searchbox.textChanged.connect(lambda:find_text(searchbox.text(),table,programs_list))
        table.resizeColumnsToContents()
        top.show()
        top.activateWindow()
        searchbox.setFocus()
        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

    @pyqtSlot()
    def patch_window(self):
        patches_list = list(set(self.comp_obj.patches_queue.get()))
        patches_list.sort(key=operator.attrgetter("date"),reverse=False)
        top = CustomDialog(self)
        top.setWindowTitle("Installed Patches on %s" % self.comp_obj.input_name)

        top_layout= QVBoxLayout()
        top.setLayout(top_layout)
        searchbox = QLineEdit()

        searchbox.setPlaceholderText("Search Patches")
        searchbox.setClearButtonEnabled(True)

        table = DialogTable(top,['Description','KB','InstalledOn'])

        top_layout.addWidget(searchbox)
        top_layout.addWidget(table)

        def find_text(search_term,table,patches_list):
            table.clearContents()
            table.setRowCount(0)
            table.setRowCount(len(patches_list))
            i=0
            for p in patches_list:
                if not search_term or search_term.lower().strip() in p.kb.lower().strip():
                    table.setRowCount(i+1)
                    description_item = QTableWidgetItem(p.description)
                    description_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,0,description_item)

                    kb_item = QTableWidgetItem(p.kb)
                    kb_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,1,kb_item)

                    date_item = QTableWidgetItem(format_date(p.date))
                    date_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                    table.setItem(i,2,date_item)
                    i+=1
        find_text("",table,patches_list)
        searchbox.textChanged.connect(lambda:find_text(searchbox.text(),table,patches_list))
        table.resizeColumnsToContents()
        top.show()
        top.activateWindow()
        searchbox.setFocus()
        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

    @pyqtSlot()
    def devices_window(self):
        devices_list = list(set(self.comp_obj.devices_queue.get()))
        #devices_list.sort(key=operator.attrgetter("name"),reverse=False)

        top = CustomDialog(self)
        top.setWindowTitle("Devices on %s" % self.comp_obj.input_name)

        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        table = DialogTable(top,['Name'])
        table.setRowCount(len(devices_list))
        top_layout.addWidget(table)

        for i,p in enumerate(devices_list):
            dev_item = QTableWidgetItem(p)
            dev_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
            table.setItem(i,0,dev_item)
        top.show()
        top.activateWindow()
        table.resizeColumnsToContents()
        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

    def add_admin_popup(self):
        self.msgboxwindow = QDialog(self)
        self.msgboxwindow.setWindowTitle("Authenticate as different user")
        self.msgboxwindow.setSizeGripEnabled(True)
        self.domainframe = AuthenticationForm(self.msgboxwindow,title="Domain Credentials (optional)",all_btns=False)
        self.localframe = AuthenticationForm(self.msgboxwindow,title="Local Credentials",all_btns=False)
        self.msgboxlayout = QVBoxLayout()
        self.msgboxwindow.setLayout(self.msgboxlayout)
        self.msgboxlayout.addWidget(self.domainframe)
        self.msgboxlayout.addWidget(self.localframe)
        self.msgboxsubmit = QPushButton("Submit",self.msgboxwindow)
        self.msgboxlayout.addWidget(self.msgboxsubmit)
        self.msgboxsubmit.clicked.connect(lambda:self.msgboxsubmit.setEnabled(False))
        self.msgboxsubmit.clicked.connect(self.adminworker.startWork)
        self.msgboxwindow.show()

    def drives_window(self):
        def on_click_drive(item):
            try:
                pathstring = item.text()
                try:
                    os.startfile(pathstring)

                except:pass
            except Exception as e:
                print(e)

        drives_list = self.comp_obj.drives_queue.get()

        top = CustomDialog(self)

        top.setWindowTitle("Mapped Drives on %s" % self.comp_obj.input_name)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        table = QTableWidget(top)
        table.verticalHeader().setVisible(False)
        table.setColumnCount(3)
        table.setHorizontalHeaderLabels(['User','Name','Path'])
        table.setRowCount(len([d for k,u in self.comp_obj.users.items() for d in u.disks]))
        top_layout.addWidget(table)
        top.show()
        top.activateWindow()

        for key,value in self.comp_obj.users.items():
            for i,p in enumerate(value.disks):
                #p.date = datetime.strptime(p.date,"%Y%m%d")
                value_item = QTableWidgetItem(value.name)
                value_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                table.setItem(i,0,value_item)

                name_item = QTableWidgetItem(p.name)
                name_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                table.setItem(i,1,name_item)

                path_item = QTableWidgetItem(p.path)
                path_item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)
                table.setItem(i,2,path_item)
        table.itemDoubleClicked.connect(on_click_drive)
        table.resizeColumnsToContents()
        self.loading_queue.decrement()
        if self.loading_queue.get() < 1:
            self.small_loading.setVisible(False)
        print(self.loading_queue.get())

    def c_drive(self):
        os.startfile("\\\\" + self.comp_obj.input_name + "\\c$\\")

    def dell_url(self):
        if self.comp_obj.serial:
            os.startfile("http://www.dell.com/support/home/us/en/19/product-support/servicetag/" + self.comp_obj.serial + "/diagnose?s=BSD")

    @pyqtSlot(str)
    def install_complete(self,ret_message):
        ret_color = "black"
        if ret_message == "0":
            ret_message = "'%s' Success!" % os.path.splitext(os.path.basename(self.comp_obj.manual_install_path))[0]
            ret_color = global_green
        elif ret_message == "3010":
            ret_message = "'%s' Success! Reboot Pending" % os.path.splitext(os.path.basename(self.comp_obj.manual_install_path))[0]
            ret_color = "orange"
        else:
            ret_message = "'%s' Failure! %s" % (os.path.splitext(os.path.basename(self.comp_obj.manual_install_path))[0],ret_message)
            ret_color = global_red
        self.install_status.setText(ret_message)
        self.install_status.setStyleSheet("color:%s;" % ret_color)
        self.install_status.mouseReleaseEvent = lambda e: self.install_results()
        self.loading_queue.decrement()
        if self.loading_queue.get() <= 0:
            self.small_loading.setVisible(False)

    def install_results(self):
        top = QDialog(self)
        top.setWindowTitle("Install Output")
        top.setSizeGripEnabled(True)
        #top.setWindowModality(Qt.ApplicationModal)
        top_layout = QVBoxLayout()
        top.setLayout(top_layout)
        titlelabel = QLabel(self.comp_obj.input_name.upper())
        font = QFont()
        font.setPointSize(16)
        titlelabel.setFont(font)
        top_layout.addWidget(titlelabel)
        top_layout.setAlignment(titlelabel,Qt.AlignCenter)

        installout = QLabel()
        installout.setStyleSheet("background-color:black;color:white")
        installout.setContentsMargins(10,10,10,10)
        top_layout.addWidget(installout)
        top_layout.setAlignment(installout,Qt.AlignTop)
        if type(self.comp_obj.out1) is str and self.comp_obj.out1:
            try:
                installout.setText(self.comp_obj.out1)
            except:
                installout.setText("                    ")
        elif type(self.comp_obj.out2) is str and self.comp_obj.out2:
            try:
                installout.setText(self.comp_obj.out2)
            except:
                installout.setText("                    ")
        elif not type(self.comp_obj.out1) is str and self.comp_obj.out1:
            try:
                installout.setText(self.comp_obj.out1.decode('utf-8'))
            except:
                installout.setText("                    ")
        elif not type(self.comp_obj.out2) is str and self.comp_obj.out2:
            try:
                installout.setText(self.comp_obj.out2.decode('utf-8'))
            except:
                installout.setText("                    ")
        else:
            installout.setText("                    ")
        top.show()

    def install_software(self,top):
        top.hide()
        self.set_loading_queue()
        self.contents.addWidget(self.install_status)
        self.contents.setAlignment(self.install_status,Qt.AlignCenter)
        self.install_status.setText("Installing...")

    def get_vbs(self):
        vbs_top = CustomDialog(self)
        vbs_top.setWindowTitle("Install Software")
        fileform = FileForm(parent=vbs_top,extensionsallowed="VB Files (*.vbs)", title="Choose Script File")
        fileform.filelabel.textChanged.connect(lambda:self.comp_obj.set_manual_install_path(fileform.filename))
        vbs_layout = QVBoxLayout()
        vbs_top.setLayout(vbs_layout)
        vbs_layout.addWidget(fileform)
        install_btn = QPushButton("Install",vbs_top)
        vbs_layout.addWidget(install_btn)
        install_btn.clicked.connect(lambda:self.install_software(vbs_top))
        install_btn.clicked.connect(self.installworker.startWork)
        vbs_top.show()
        vbs_top.activateWindow()

    def check_mouse_btn(self,e,textdata="",leftcallback=None,callback_label=""):
        if e.button() == Qt.RightButton:
            if leftcallback:
                self.popup = QMenu(self)
                context_action = QAction(callback_label,self.popup)
                context_action.triggered.connect(leftcallback)
                self.popup.addAction(context_action)
                self.popup.popup(QCursor.pos())
        else:
            self.clipboard_callback(textdata=textdata)

    def set_loading_queue(self):
        self.small_loading.setVisible(True)
        self.loading_queue.increment()

    @pyqtSlot()
    def dataWidgets(self):
        self.setStyleSheet("font-size:10pt;font-family:sans-serif")
        self.placeholder.hide()

        self.color_label = QLabel(" ",self)

        self.color_label.setFixedWidth(10)

        default_r = self.palette().color(self.backgroundRole()).red()
        default_g = self.palette().color(self.backgroundRole()).green()
        default_b = self.palette().color(self.backgroundRole()).blue()

        if self.comp_obj.serial:
            self.color_label.setStyleSheet("background-color:%s" % rgb_to_hex(*generate_color(self.comp_obj.serial,default_r,default_g,default_b)))
        else:
            self.color_label.setStyleSheet("background-color:black")

        self.layout.addWidget(self.color_label)
        self.contents = QVBoxLayout()
        self.textcontents = QHBoxLayout()
        self.contents.addLayout(self.textcontents)
        self.layout.addLayout(self.contents)
        self.layout.setContentsMargins(0,0,0,0)

        self.col_one_layout = QVBoxLayout()
        self.col_one_layout.setContentsMargins(0,11,0,11)
        self.textcontents.addLayout(self.col_one_layout)

        self.name_layout = CustomDataLabel(self,"Name:", self.comp_obj.name.upper(), 50, self.c_drive, "Go to C: Drive", self.link_color)
        self.col_one_layout.addLayout(self.name_layout)

        if not self.comp_obj.status:
            self.serial_layout = CustomDataLabel(self,"Serial:", self.comp_obj.serial, 50, self.dell_url, "Find on Dell.com",self.link_color)
            self.col_one_layout.addLayout(self.serial_layout)

            self.model_layout = CustomDataLabel(self,"Model:",self.comp_obj.model, 50, None,None,self.link_color)
            self.col_one_layout.addLayout(self.model_layout)


            self.user_layout = CustomDataLabel(self,"User:", self.comp_obj.user, 50, None, None, self.link_color)
            self.col_one_layout.addLayout(self.user_layout)

            self.os_layout = CustomDataLabel(self,"OS:",self.comp_obj.os, 50, None, None, self.link_color)
            self.col_one_layout.addLayout(self.os_layout)

            self.col_two_layout = QVBoxLayout()
            self.col_two_layout.setContentsMargins(0,11,11,11)

            self.textcontents.addLayout(self.col_two_layout)

            temp_ip = ""
            try: temp_ip = self.comp_obj.ip_addresses[0]
            except: pass

            self.ip_layout = CustomDataLabel(self,"IP:", temp_ip, 75, None, None, self.link_color)
            self.col_two_layout.addLayout(self.ip_layout)

            self.resolution_layout = CustomDataLabel(self,"Resolution:", self.comp_obj.resolution, 75, None, None, None)
            self.col_two_layout.addLayout(self.resolution_layout)

            self.monitors_layout = CustomDataLabel(self,"Monitors:", str(self.comp_obj.monitors), 75, None, None, None)
            self.col_two_layout.addLayout(self.monitors_layout)


            self.cpu_layout = CustomDataLabel(self,"CPU:",self.comp_obj.cpu, 75, None, None, None)
            self.col_two_layout.addLayout(self.cpu_layout)

            self.memory_layout = CustomDataLabel(self, "Memory:",self.comp_obj.memory, 75, None, None, None)
            self.col_two_layout.addLayout(self.memory_layout)

            self.toolbtns_layout = QHBoxLayout()
            self.contents.addLayout(self.toolbtns_layout)
            self.utilbtn = QToolButton(self)
            self.utilgroup = QMenu()

            if not self.comp_obj.local:
                self.dameware_action = QAction("Dameware")
                self.dameware_action.triggered.connect(self.damewareworker.startWork)
                self.dameware_action.triggered.connect(self.set_loading_queue)
                self.utilgroup.addAction(self.dameware_action)

            self.gpedit_action = QAction("GPEdit")
            self.gpedit_action.triggered.connect(self.gpedit)
            self.utilgroup.addAction(self.gpedit_action)

            self.compmgmt_action = QAction("Computer Management")
            self.compmgmt_action.triggered.connect(self.compmgmt)
            self.utilgroup.addAction(self.compmgmt_action)
            self.utilgroup.addSeparator()

            if not self.comp_obj.local:
                self.install_action = QAction("Install Software")
                self.install_action.triggered.connect(self.get_vbs)
                self.utilgroup.addAction(self.install_action)


                if 'group to add to admin' in self.settings.settings_dict:
                    if self.settings.settings_dict["group to add to admin"]:
                        self.addadmin_action = QAction("Add '%s'" % self.settings.settings_dict['group to add to admin'])
                        self.addadmin_action.triggered.connect(self.add_admin_popup)
                        self.utilgroup.addAction(self.addadmin_action)

                self.remotecmd_action = QAction("Remote Command Line")
                self.remotecmd_action.triggered.connect(self.remote_cmd)
                self.utilgroup.addAction(self.remotecmd_action)

                self.default_opener = QAction("Remote Tools")
            else:
                self.default_opener = QAction("Tools")


            self.utilbtn.setMenu(self.utilgroup)

            self.default_opener.triggered.connect(self.utilbtn.showMenu)
            self.utilbtn.setDefaultAction(self.default_opener)

            self.utilbtn.setPopupMode(QToolButton.MenuButtonPopup)


            self.moreinfobtn = QToolButton(self)
            self.moreinfogroup = QMenu()

            if not self.comp_obj.manual_user:
                self.get_devices_action = QAction("Get Devices")
                self.get_devices_action.triggered.connect(self.deviceworker.startWork)
                self.get_devices_action.triggered.connect(self.set_loading_queue)
                self.moreinfogroup.addAction(self.get_devices_action)

                self.get_drives_action = QAction("Get Drives")
                self.get_drives_action.triggered.connect(self.driveworker.startWork)
                self.get_drives_action.triggered.connect(self.set_loading_queue)
                self.moreinfogroup.addAction(self.get_drives_action)

            self.get_printers_action = QAction("Get Printers")
            self.get_printers_action.triggered.connect(self.printworker.startWork)
            self.get_printers_action.triggered.connect(self.set_loading_queue)
            self.moreinfogroup.addAction(self.get_printers_action)

            self.get_programs_action = QAction("Get Programs")
            self.get_programs_action.triggered.connect(self.progworker.startWork)
            self.get_programs_action.triggered.connect(self.set_loading_queue)
            self.moreinfogroup.addAction(self.get_programs_action)

            self.get_patches_action = QAction("Get Patches")
            self.get_patches_action.triggered.connect(self.patchworker.startWork)
            self.get_patches_action.triggered.connect(self.set_loading_queue)
            self.moreinfogroup.addAction(self.get_patches_action)

            self.moreinfobtn.setMenu(self.moreinfogroup)
            self.default_info_opener = QAction("Get More Info")
            self.default_info_opener.triggered.connect(self.moreinfobtn.showMenu)
            self.moreinfobtn.setDefaultAction(self.default_info_opener)
            self.moreinfobtn.setPopupMode(QToolButton.MenuButtonPopup)

            self.small_loading = QLabel()
            self.small_loading.setContentsMargins(0,0,0,0)
            size_pol = QSizePolicy()
            size_pol.setRetainSizeWhenHidden(True)
            self.small_loading.setSizePolicy(size_pol)
            self.small_loading.setScaledContents(True)
            self.small_loading.setMaximumHeight(self.utilbtn.height())

            self.small_loading.setAlignment(Qt.AlignCenter)
            if self.main_path:
                self.gif2 = QMovie(self.main_path + "\\loading2.gif")
            else:
                self.gif2 = QMovie("loading2.gif")
            self.small_loading.setMovie(self.gif2)
            testsize = QSize()
            testsize.setHeight(self.utilbtn.height())
            testsize.setWidth(self.utilbtn.height())
            self.gif2.setScaledSize(testsize)
            self.small_loading.setAttribute(Qt.WA_TranslucentBackground,True)
            self.gif2.start()
            self.small_loading.setVisible(False)

            self.copy_all_btn = QPushButton("Copy Specs")
            self.copy_all_btn.clicked.connect(self.all_to_clipboard)

            self.toolbtns_layout.addStretch()
            self.toolbtns_layout.addWidget(self.utilbtn)
            self.toolbtns_layout.addWidget(self.moreinfobtn)
            self.toolbtns_layout.addWidget(self.copy_all_btn)
            self.toolbtns_layout.addWidget(self.small_loading)
            self.toolbtns_layout.addStretch()

        else:
            self.status_layout = QHBoxLayout()
            self.statuslabel = QLabel("Error:",self)
            self.statuslabel.setFixedWidth(50)
            self.status = QLabel(self.comp_obj.status,self)
            self.status_layout.addWidget(self.statuslabel)
            self.status_layout.addWidget(self.status)
            self.col_one_layout.addLayout(self.status_layout)

        self.delete_frame_btn = QPushButton("X",self)
        self.delete_frame_btn.setMaximumWidth(self.delete_frame_btn.height())
        self.delete_frame_btn.clicked.connect(self.deleteLater)
        self.layout.addWidget(self.delete_frame_btn)
        self.layout.setAlignment(self.delete_frame_btn,Qt.AlignTop)

if __name__ == "__main__":
    ico_path = ""
    try:
        if hasattr(sys,'frozen'):
            ico_path = sys.executable
        else:
            ico_path = sys.argv[0]

        if not os.path.exists(os.getenv("APPDATA") + '\\Single Computer Info'):
            os.makedirs(os.getenv("APPDATA") + '\\Single Computer Info')
        if not os.path.exists(os.getenv("APPDATA") + '\\Single Computer Info\\comp_info.cfg'):
            copyfile(os.path.dirname(ico_path) + '\\comp_info.cfg',os.getenv("APPDATA") + '\\Single Computer Info\\comp_info.cfg')
    except Exception as e: print(e)
    wind = QApplication(sys.argv)
    wind.setWindowIcon(QIcon('single_logo.ico'))

    if not ctypes.windll.UxTheme.IsThemeActive():
        wind.setStyle('Fusion')

    parser = argparse.ArgumentParser(description="Find information on computer based on name or IP")
    parser.add_argument('filename', nargs="?")
    parser.add_argument('-debug', default=False, action='store_true')
    args = parser.parse_args()
    try:
        uriname = args.filename.split(":")[1]
    except:
        uriname = args.filename
    app = SinglePCApp(debug=args.debug,main_wind=wind,main_path=os.path.dirname(ico_path),start_name=uriname)
    sys.exit(wind.exec_())
