# -*- coding: utf-8 -*-
"""
/***************************************************************************
 SPREADSHEET
                                 A QGIS plugin
 This plugin will  scan all spreadsheet in specified folder for a particular circuit 
                              -------------------
        begin                : 2017-05-05
        git sha              : $Format:%H$
        copyright            : (C) 2017 by Rau.D
        email                : raul.dasgupta@vocus.com.au
 ***************************************************************************/

/***************************************************************************
 *                                                                         *
 *   This program is free software; you can redistribute it and/or modify  *
 *   it under the terms of the GNU General Public License as published by  *
 *   the Free Software Foundation; either version 2 of the License, or     *
 *   (at your option) any later version.                                   *
 *                                                                         *
 ***************************************************************************/
"""
from PyQt4.QtCore import QSettings, QTranslator, qVersion, QCoreApplication
from PyQt4.QtGui import QAction, QIcon, QFileDialog, QApplication
from PyQt4 import QtCore, QtGui
from PyQt4.QtGui import QTextEdit
# from spreadsheet_dialog import SPREADSHEETDialog
# Initialize Qt resources from file resources.py
import resources
# Import the code for the dialog
from spreadsheet_dialog import SPREADSHEETDialog
import os.path
import os
import sys
import xlrd
import re
# import Spreadsheet_timestamp_new
import glob


# class Port:
#     def __init__(self,view):
#         self.view = view
#
#     def write(self,*args):
#         self.view.append(*args)


# # initialize output slots
# class EmittingStream(QtCore.QObject):
#
#     textWritten = QtCore.pyqtSignal(str)
#
#     def write(self, text):
#         self.textWritten.emit(str(text))

# class loading_bar(QtGui.QDialog):
#     def __init__(self, parent=None):
#
#         QtGui.QDialog.__init__(self, parent, Qt.WindowStaysOnTopHint)
#
#         self.progressBar = QtGui.QProgressBar(self)
#         self.progressBar.setGeometry(QtCore.QRect(10, 30, 381, 10))
#         self.progressBar.setMinimumSize(QtCore.QSize(141, 10))
#         self.progressBar.setMaximumSize(QtCore.QSize(381, 10))
#         self.progressBar.setRange(0, 100)
#         self.progressBar.setProperty("value", 0)
#         self.progressBar.setObjectName("progressBar")
#         self.label = QtGui.QLabel(self)
#         self.label.setGeometry(QtCore.QRect(10, 10, 141, 16))
#         self.label.setObjectName("label")
#
#     def show_loading_bar(self, text):
#
#         self.setObjectName("loading_bar_background")
#         self.setEnabled(True)
#         sizePolicy = QtGui.QSizePolicy(QtGui.QSizePolicy.Preferred, QtGui.QSizePolicy.Preferred)
#         sizePolicy.setHorizontalStretch(0)
#         sizePolicy.setVerticalStretch(0)
#         sizePolicy.setHeightForWidth(self.sizePolicy().hasHeightForWidth())
#         self.setSizePolicy(sizePolicy)
#         self.setMinimumSize(QtCore.QSize(160, 30))
#         self.setMaximumSize(QtCore.QSize(400, 50))
#         self.setWindowOpacity(1.0)
#         self.setAutoFillBackground(False)
#         self.label.setText(text)
#         self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
#
#         self.show()


class SPREADSHEET:
    """QGIS Plugin Implementation."""

    def __init__(self, iface):
        """Constructor.

        :param iface: An interface instance that will be passed to this class
            which provides the hook by which you can manipulate the QGIS
            application at run time.
        :type iface: QgsInterface
        """
        # Save reference to the QGIS interface
        self.iface = iface
        # initialize plugin directory
        self.plugin_dir = os.path.dirname(__file__)
        # initialize locale
        locale = QSettings().value('locale/userLocale')[0:2]
        locale_path = os.path.join(
            self.plugin_dir,
            'i18n',
            'SPREADSHEET_{}.qm'.format(locale))

        if os.path.exists(locale_path):
            self.translator = QTranslator()
            self.translator.load(locale_path)

            if qVersion() > '4.3.3':
                QCoreApplication.installTranslator(self.translator)

        self.dlg = SPREADSHEETDialog()
        # Declare instance attributes
        self.actions = []
        self.menu = self.tr(u'&Spreadsheet Scanner')
        # TODO: We are going to let the user set this up in a future iteration
        self.toolbar = self.iface.addToolBar(u'SPREADSHEET')
        self.toolbar.setObjectName(u'SPREADSHEET')
        # self.dlg.Console = self.dlg.scroll.widget()
        self.progress = 0
        self.steps = 0
        self.dlg.progressBar.setValue(0)
        self.dlg.progressBar.setVisible(False)
        # self.dlg.progressBar.setTextVisible(True)
        # self.dlg.progressBar.setFormat("Awaiting Input")
        self.dlg.input_directory.clear()
        self.dlg.input_button.clicked.connect(self.select_directory)
        self.dlg.search.clicked.connect(self.intial_display)
        # self.dlg.search.pressed.connect(self.intial_display)
        self.dlg.search.clicked.connect(self.circuit_finder)


        # # Install the custom output stream
        # sys.stdout = EmittingStream(textWritten=self.normalOutputWritten)

    # noinspection PyMethodMayBeStatic
    def tr(self, message):
        """Get the translation for a string using Qt translation API.

        We implement this ourselves since we do not inherit QObject.

        :param message: String for translation.
        :type message: str, QString

        :returns: Translated version of message.
        :rtype: QString
        """
        # noinspection PyTypeChecker,PyArgumentList,PyCallByClass
        return QCoreApplication.translate('SPREADSHEET', message)


    def add_action(
        self,
        icon_path,
        text,
        callback,
        enabled_flag=True,
        add_to_menu=True,
        add_to_toolbar=True,
        status_tip=None,
        whats_this=None,
        parent=None):
        """Add a toolbar icon to the toolbar.

        :param icon_path: Path to the icon for this action. Can be a resource
            path (e.g. ':/plugins/foo/bar.png') or a normal file system path.
        :type icon_path: str

        :param text: Text that should be shown in menu items for this action.
        :type text: str

        :param callback: Function to be called when the action is triggered.
        :type callback: function

        :param enabled_flag: A flag indicating if the action should be enabled
            by default. Defaults to True.
        :type enabled_flag: bool

        :param add_to_menu: Flag indicating whether the action should also
            be added to the menu. Defaults to True.
        :type add_to_menu: bool

        :param add_to_toolbar: Flag indicating whether the action should also
            be added to the toolbar. Defaults to True.
        :type add_to_toolbar: bool

        :param status_tip: Optional text to show in a popup when mouse pointer
            hovers over the action.
        :type status_tip: str

        :param parent: Parent widget for the new action. Defaults None.
        :type parent: QWidget

        :param whats_this: Optional text to show in the status bar when the
            mouse pointer hovers over the action.

        :returns: The action that was created. Note that the action is also
            added to self.actions list.
        :rtype: QAction
        """

##        # Create the dialog (after translation) and keep reference
##        self.dlg = SPREADSHEETDialog()

        icon = QIcon(icon_path)
        action = QAction(icon, text, parent)
        action.triggered.connect(callback)
        action.setEnabled(enabled_flag)

        if status_tip is not None:
            action.setStatusTip(status_tip)

        if whats_this is not None:
            action.setWhatsThis(whats_this)

        if add_to_toolbar:
            self.toolbar.addAction(action)

        if add_to_menu:
            self.iface.addPluginToMenu(
                self.menu,
                action)

        self.actions.append(action)

        return action

    def initGui(self):
        """Create the menu entries and toolbar icons inside the QGIS GUI."""

        icon_path = ':/plugins/SPREADSHEET/icon.png'
        self.add_action(
            icon_path,
            text=self.tr(u'Spreadsheet Scanner'),
            callback=self.run,
            parent=self.iface.mainWindow())

    def unload(self):
        """Removes the plugin menu item and icon from QGIS GUI."""
        for action in self.actions:
            self.iface.removePluginMenu(
                self.tr(u'&Spreadsheet Scanner'),
                action)
            self.iface.removeToolBarIcon(action)
        # remove the toolbar
        del self.toolbar

    def select_directory(self):
        folder = QFileDialog.getExistingDirectory(self.dlg, "Select directory to scan")
        self.dlg.input_directory.setText(folder)

    def recent_spreadsheet(self, files_list):
        folder = str(self.dlg.input_directory.text()) + "/"
        recent_file_list = list()
        error_file_list = list()
        archived_file_list = list()
        # self.dlg.Console.append ("THINKING...............\n")
        for files in files_list:
            try:
                # Create a list of files having a similar name
                file_list = list()
                base = os.path.basename(files)
                f = os.path.splitext(base)[0]
                f_split = f.rsplit(')', 1)[0]
                # print f
                for name in glob.glob(folder + '\\' + f_split + '*.*'):
                    file_list.append(name)

                # From the list select the file that has the most recent modified time.
                latest_file = max(file_list, key=os.path.getmtime)

                # Append the most recent file to the recent files list.
                if latest_file not in recent_file_list:
                    recent_file_list.append(latest_file)
                    # print 'appended'
                for archived_files in file_list:
                    if archived_files != latest_file:
                        if archived_files not in archived_file_list:
                            archived_file_list.append(archived_files)

            except Exception as error:
                error_list = list()
                # print error
                if files not in error_file_list:
                    error_list.append(files)
                    error_list.append(error)
                    error_file_list.append(error_list)

                if files not in recent_file_list:
                    recent_file_list.append(files)
        self.steps += self.progress
        self.dlg.progressBar.setValue(self.steps)
        # print "something......."
        # print len(recent_file_list),len(archived_file_list)
        return recent_file_list, archived_file_list

    # Function scans the specified folder for all spreadsheet with a particular address in the name
    def spreadsheet_scanner(self):

        address= str(self.dlg.address.text())
        # address = raw_input("Enter the destination address:")
        # address='l3 151 royal st'
        file = list()
        extensions = ('.xls', '.xlsx')
        folder = str(self.dlg.input_directory.text()) + "/"
        if self.dlg.checkBox.isChecked():
            f = [os.path.join(root, name)
                         for root, dirs, files in os.walk(folder)
                         for name in files
                         if name.endswith((".xls", ".xlsx"))]
            for i in f:
                if os.path.isfile(os.path.join(folder, i)) and address.lower() in i.lower():
                    if any(i.lower().endswith(ext) for ext in extensions):
                        file.append(i)
        for i in os.listdir(folder):
            if os.path.isfile(os.path.join(folder, i)) and address.lower() in i.lower():
                if any(i.lower().endswith(ext) for ext in extensions):
                    file.append(i)
        if len(file) != 0:
            self.progress = 100/len(file)+2
            # self.dlg.Console.append(str(self.progress))
        else:
            self.progress = 100
            # self.dlg.Console.append(str(self.progress))
        self.steps += self.progress
        self.dlg.progressBar.setValue(self.steps)
        recent_files, archived_files = self.recent_spreadsheet(file)

        return recent_files, archived_files

    # Scans the list of possible spreadsheet for the specified circuit id and prints the results
    def circuit_finder(self):
        # self.dlg.progressBar.setFormat("Scanning.....")
        circuit_id = str(self.dlg.circuit.text())
        # circuit_id = raw_input("Enter the circuit id:")
        id_pattern = re.compile(circuit_id.lower())
        master_list = list()
        counter = 0
        # Courtesy Vocus import Code
        rec_files, archived_files = self.spreadsheet_scanner()
        master_list.append(rec_files)
        master_list.append(archived_files)
        old_sheet_pattern = re.compile(r'old', flags=re.I | re.X)
        folder = str(self.dlg.input_directory.text()) + "/"
        if len(rec_files) != 0 or len(archived_files) != 0:

            for f_list in master_list:
                # Start reading all the potential spreadsheets
                if len(f_list) != 0:
                    try:
                        # print "Still Thinking........,btw number of matched files is ", len(files)
                        match = list()
                        # count = 0
                        for i in f_list:
                            self.steps += self.progress
                            self.dlg.progressBar.setValue(self.steps)
                            # print "#", count
                            # if os.stat(i) and os.path.isfile(i) is True:
                            try:
                                # print "something"
                                file_name = os.path.basename(i)
                                # print "something"
                                spreadsheet = xlrd.open_workbook(os.path.join(folder, i))
                                sheet_names = spreadsheet.sheet_names()
                                # print"something"
                                for sheet in sheet_names:
                                    if old_sheet_pattern.findall(sheet):
                                        print i, ":", sheet, " references old cable"
                                    else:

                                        ws = spreadsheet.sheet_by_name(sheet)

                                        merged__cells = ws.merged_cells

                                        data = list()
                                        for row in range(0, ws.nrows):
                                            formatted_row = list()

                                            for col in range(0, ws.ncols):
                                                try:
                                                    formatted_row.append(str(ws.cell_value(row, col)))

                                                except (Exception, UnicodeEncodeError):
                                                    uni = ws.cell_value(row, col)
                                                    string = uni.encode('utf-8')
                                                    formatted_row.append(string)

                                            if any(formatted_row):
                                                data.append(formatted_row)

                                        # Get the index of the point where the audit sheet begins
                                        # for l in data:
                                        #     dat=list(l)
                                        #     if dat.count('CABLE')==2 and dat.count('CORE')==2:
                                        #           audit_idx=data.index(dat)

                                        # Scans the information for the indication of the circuit id and prints it .
                                        # print ("Spreadsheet, Cable, Core, Colour, To, Cable, Core, Colour, Comments")
                                        sheet_match = list()
                                        for dat in data:
                                            for d in dat:
                                                cir = id_pattern.search(d.lower())
                                                if cir is not None:
                                                    # print i,":",dat
                                                    sheet_match.append(file_name)
                                                    sheet_match.append(dat)
                                                    if sheet_match not in match:
                                                        match.append(sheet_match)
                                                        # else:
                                                        #     print 'no match found'
                                spreadsheet.release_resources()
                                del spreadsheet
                                # count = count+1
                            except(Exception, xlrd.XLRDError) as file_error:
                                print i, ':', file_error
                                # count += 1
                    # except (Exception, xlrd.XLRDError)as error:
                    #     pass
                        if len(match) != 0:
                            if counter < 1:
                                # self.dlg.Console.append("Something"/n)
                                # sys.stdout = Port(self.dlg.Console("Something"))
                                self.dlg.Console.append ('**************Recent Files****************************************\n')
                                self.dlg.Console.append ("Spreadsheet, Cable, Core, Colour, To, Cable, Core, Colour, Comments\n")
                                for i in match:
                                    # CHECKS IF THE RESULT IS REPEATING ITSELF,
                                    # IF TRUE THEN DISPLAY RESULT AS IS, OTHERWISE DISPLAY THE FIRST TWO ELEMENTS
                                    self.dlg.Console.append((str(i)+'\n'))


                                    # else:
                                    #     for obj in range(len(i)):
                                    #         self.dlg.Console.append(obj[0])
                                    #         self.dlg.Console.append(obj[1])
                                    #         self.dlg.Console.append('\n')
                            else:
                                self.dlg.Console.append('**************Archived****************************************\n')
                                self.dlg.Console.append ("Spreadsheet, Cable, Core, Colour, To, Cable, Core, Colour, Comments\n")
                                for i in match:
                                    self.dlg.Console.append((str(i) + '\n'))
                        else:
                            if counter < 1:

                                self.dlg.Console.append ('***************Recent Files**********************************\n')
                                self.dlg.Console.append ("No Match Found\n")
                            else:
                                self.dlg.Console.append ('**************Archived***************************************\n')
                                self.dlg.Console.append('no match found\n')
                    except (Exception, xlrd.XLRDError)as error:
                        error_msg = "Error:"+str(error)+"\n"
                        self.dlg.Console.append(error_msg)
                        # self.dlg.Console.append("\n")
                        # self.dlg.Console.append spreadsheet
                        if 'match' in locals():
                            if len(match) != 0:
                                self.dlg.Console.append ("Spreadsheet, Cable, Core, Colour, To, Cable, Core, Colour, Comments\n")
                                for i in match:
                                    self.dlg.Console.append((str(i) + '\n'))

                counter += 1
        if len(rec_files) == 0 and len(archived_files) == 0:
            self.dlg.Console.append("No spreadsheet found with matching address.\n")
        self.dlg.Console.append("\n\n**************Scan Complete******************\n\n")
        self.dlg.progressBar.setVisible(False)
        # self.steps += self.progress
        # self.dlg.progressBar.setValue(self.steps)
        # self.dlg.progressBar.setFormat("Scan Completed")


    # # append text to the QTextEdit
    # def normalOutputWritten(self, text):
    #     """Append text to the QTextEdit."""
    #     # maybe QTextEdit.append() will work as well, but this is how I do it:
    #     # cur=self.dlg.Console.QTextEdit()
    #     cursor = self.dlg.Console.textCursor()
    #     font = QtGui.QFont("Monaco", 11, True)
    #     # cursor.movePosition(self.dlg.QtGui.QTextCursor.End)
    #     cursor.movePosition(cursor.End)
    #     # cursor.insertText(text)
    #     self.dlg.Console.setTextCursor(cursor)
    #     self.dlg.Console.ensureCursorVisible()
    #     self.dlg.Console.setFont(font)
    def intial_display(self):
        self.dlg.progressBar.reset()
        self.dlg.progressBar.setVisible(True)
        self.progress = 0
        self.steps = 0
        font = QtGui.QFont("Monaco", 11, True)
        self.dlg.Console.setFont(font)
        self.dlg.Console.clear()
        self.dlg.Console.append("--------------------------------------------------------------------------\n")
        # self.dlg.Console.append("\n")
        if self.dlg.checkBox.isChecked():
            self.dlg.Console.append("Scanning all folders in "+ str(self.dlg.input_directory.text())+" for "
                                    +(str(self.dlg.address.text()+"\n")).upper())
        else:
            self.dlg.Console.append(('Now Scanning spreadsheets in  '+str(self.dlg.input_directory.text())+' for '
                                 + (str(self.dlg.address.text()+"\n")).upper()))
        self.dlg.Console.append(("Searching spreadsheets for " + str(self.dlg.circuit.text()+'\n')))
        self.dlg.Console.append("----------------------------------------------------------------------------\n")

    def run(self):
        """Run method that performs all the real work"""
        # show the dialog
        self.dlg.show()
        # Run the dialog event loop
        # result = self.dlg.exec_()
        # # See if OK was pressed
        # if result:
        #     # Do something useful here - delete the line containing pass and
        #     # substitute with your code.
        #     pass
