#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os
import pandas as pd
from PyQt5 import QtCore, QtWidgets, uic
from PyQt5.QtWidgets import QMainWindow, QAction, QMessageBox, QFileDialog, QCheckBox, QGroupBox, QTableWidgetItem, QTableWidget, QLineEdit
from PyQt5.QtCore import QSettings, QTimer
import inspect
from datetime import datetime, timedelta

import matplotlib
matplotlib.use("Qt5Agg")
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt5agg import NavigationToolbar2QT as NavigationToolbar
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.figure import Figure
import matplotlib.pyplot as plt

#============================================================================================================
class Timerwindow (QMainWindow):
    def __init__(self):
        super(Timerwindow, self).__init__()
        uic.loadUi('tracker2.ui', self)
        self.setupui()
        self.set_menu() 
        self.location()

    def setupui(self):
        self.setStyleSheet("""
            QPushButton:hover {background-color: LightCyan;
            }
            QLineEdit:hover {background-color: AliceBlue;
            }
            QTabBar::tab:selected {background: Lightgrey;
            }
            QTabWidget>QWidget>QWidget{background: Lightgrey;
            }
                           """)

        # self.tabWidget.setStyleSheet('QTabBar::tab {background-color: red;}')
        self.tabWidget.setCurrentIndex(0)        

        #Config fig 1
        self.figure1 = Figure()
        self.canvas1 = FigureCanvas(self.figure1)
        self.plotlayout.addWidget(self.canvas1)        

        #Config fig 2
        self.figure2 = Figure()
        self.canvas2 = FigureCanvas(self.figure2)
        self.plotlayout_2.addWidget(self.canvas2)        

        #FigureCanvas.setSizePolicy(self, QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.toolbar = NavigationToolbar(self.canvas1, self)  
        self.plotlayout.addWidget(self.toolbar)

        self.toolbar2 = NavigationToolbar(self.canvas2, self)  
        self.plotlayout_2.addWidget(self.toolbar2)        

        #============================================================================================================       
        self.pushButtonStart.released.connect(self.timer_start)
        self.pushButtonStop.released.connect(self.timer_stop)

        self.pushButtonRemove.released.connect(self.Remove)
        self.pushButtonExport.released.connect(self.saveCSV)        

        self.pushButtonFilter.released.connect(self.filt)

        #============================================================================================================       
        self.tableWidget.setStyleSheet("border: 1px solid grey; border-radius: 3px; font: EADS Sans; font-size: 10;")
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setHorizontalHeaderLabels(('Date', 'Project', 'Start', 'Finish', 'Delta'))    
        self.tableWidget.cellChanged.connect(self.row)        

        self.lineEditProject.returnPressed.connect(self.timer_start)
        self.combo()        

        # self.comboBox.currentIndexChanged(self.filt)
        # self.comboBox_2.currentIndexChanged(self.filt)

        #============================================================================================================
    def set_menu (self):
        main_menu = self.menuBar()
        main_menu.setNativeMenuBar(False) #needed for mac       

        file_menu = main_menu.addMenu('File')

        new_session = QAction("New", self)
        new_session.setShortcut("Ctrl+N")
        new_session.setStatusTip("Click to start new session")
        file_menu.addAction(new_session)
        new_session.triggered.connect(self.newSession)     

        save_csv = QAction("Save project list as CSV", self)
        save_csv.setShortcut("Ctrl+C")
        save_csv.setStatusTip("Click to save CSV")
        file_menu.addAction(save_csv)
        save_csv.triggered.connect(self.saveCSV) 

        printer = QAction("Print", self)
        printer.setShortcut("Ctrl+P")
        printer.setStatusTip("Click to print Plot")
        file_menu.addAction(printer)
        printer.triggered.connect(self.Print)

        quiter = QAction("Quit", self)
        quiter.setShortcut("Ctrl+Q")
        quiter.setStatusTip("Click to Exit")
        file_menu.addAction(quiter)
        quiter.triggered.connect(self.closeSession)

        session_menu = main_menu.addMenu('Session')

        save_session = QAction("Save session", self)
        save_session.setShortcut("Ctrl+S")
        save_session.setStatusTip("Click to save session")
        session_menu.addAction(save_session)
        save_session.triggered.connect(self.saveSession)         

        restore_session = QAction("Restore session", self)
        restore_session.setShortcut("Ctrl+R")
        restore_session.setStatusTip("Click to restore session")
        session_menu.addAction(restore_session)
        restore_session.triggered.connect(self.restoreSession)        

        #============================================================================================================        
    def location(self):
        self.location = (os.path.join(os.path.join(os.path.expanduser('~')), 'Documents', 'Tracker'))
        if os.path.exists(self.location):
            print('Location exists')
        else:
            os.makedirs(self.location)        

        msg = QMessageBox.question(self, 'restore session',
                                    'Do you want to restore a previous session?',
                                    QMessageBox.Yes | QMessageBox.No)
        if msg == QMessageBox.Yes:
            self.restoreSession()
        if msg == QMessageBox.No:
            pass
    
        #============================================================================================================               
    def row(self):
        # numRow = self.tableWidget.rowCount()
        # self.tableWidget.insertRow(numRow)
        self.tableWidget.resizeColumnsToContents()
        self.load()
        # pass

        #============================================================================================================               
    def combo(self, poscol=0):
        num_of_rows = self.tableWidget.rowCount()
        # print(num_of_rows)
        b=[]
        for row in range(num_of_rows):
            df_list2= []
            table_item = self.tableWidget.item(row,poscol)
            df_list2.append('' if table_item is None else str(table_item.text()))
            [b.append(x) for x in df_list2 if x not in b]
            # print(b)
            self.comboBox.clear()
            self.comboBox_2.clear()
            # self.comboBox.addItems(b)
            for word in b:
                self.comboBox.addItem(word)
                self.comboBox_2.addItem(word)        

        # print(len(b))
        if len(b) >0:
            self.comboBox.setCurrentIndex(0)
            self.comboBox_2.setCurrentIndex(len(b)-1)
        else:
            pass

        #============================================================================================================               
    def timer_start(self):
        self.today = datetime.now().strftime('%d %b %y')
        self.start_time = datetime.now().time().strftime('%H:%M:%S')
        self.pushButtonStart.setDisabled(True)
        self.pushButtonStop.setDisabled(False)
        self.lineEditProject.setDisabled(True)
    
    def timer_stop(self):
        self.end_time = datetime.now().time().strftime('%H:%M:%S')
        self.pushButtonStart.setDisabled(False)
        self.pushButtonStop.setDisabled(True)
        self.lineEditProject.setDisabled(False)
    
        total_time = (datetime.strptime(self.end_time,'%H:%M:%S') - datetime.strptime(self.start_time,'%H:%M:%S'))
        # return total_time
        #get file name and add it to the table view
        project = self.lineEditProject.text()

        #Create empty row at bottom of table
        numRows = self.tableWidget.rowCount()
    
        self.tableWidget.insertRow(numRows)

        #Add text to row
        self.tableWidget.setItem(numRows, 0, QTableWidgetItem(self.today))
        self.tableWidget.setItem(numRows, 1, QTableWidgetItem(project))
        self.tableWidget.setItem(numRows, 2, QTableWidgetItem(self.start_time))
        self.tableWidget.setItem(numRows, 3, QTableWidgetItem(self.end_time))
        self.tableWidget.setItem(numRows, 4, QTableWidgetItem(str(total_time)))

        self.lineEditProject.clear()

        #============================================================================================================
    def load (self):
        num_of_rows = self.tableWidget.rowCount()
        num_of_col = self.tableWidget.columnCount()
        headers = [str(self.tableWidget.horizontalHeaderItem(i).text())for i in range(num_of_col)]

        df_list = []
        for row in range(num_of_rows):
            df_list2= []
            for col in range(num_of_col):
                table_item = self.tableWidget.item(row, col)
                df_list2.append('' if table_item is None else str(table_item.text()))
            df_list.append(df_list2)

        df = pd.DataFrame(df_list, columns=headers)

        df[['Start', 'Finish']]=df[['Start', 'Finish']].apply(pd.to_timedelta)
        df['Date']=df['Date'].astype('Datetime64').dt.strftime("%d %b %y")
        df['Delta']=df.apply(lambda row: round((row.Finish - row.Start).seconds/3600,2), axis=1)
        # print(df)        

        self.plotBar(df) 
        self.plotPie(df)

        self.combo()

        #============================================================================================================
    def filt (self):
        num_of_rows = self.tableWidget.rowCount()
        num_of_col = self.tableWidget.columnCount()
        headers = [str(self.tableWidget.horizontalHeaderItem(i).text())for i in range(num_of_col)]

        df_list = []
        for row in range(num_of_rows):
            df_list2= []
            for col in range(num_of_col):
                table_item = self.tableWidget.item(row, col)
                df_list2.append('' if table_item is None else str(table_item.text()))
            df_list.append(df_list2)

        df = pd.DataFrame(df_list, columns=headers)
        
        df[['Start', 'Finish']]=df[['Start', 'Finish']].apply(pd.to_timedelta)
        df['Date']=df['Date'].astype('Datetime64').dt.strftime("%d %b %y")
        df['Delta']=df.apply(lambda row: round((row.Finish - row.Start).seconds/3600,2), axis=1)
        # print(df)
        fr = self.comboBox.currentText().strip()
        unt = self.comboBox_2.currentText().strip()

        if fr != '':
            start = datetime.strptime(str(fr), '%d %b %y').date()
        if unt != '':
            finish = datetime.strptime(str(unt), '%d %b %y').date()

            # print('Start = {}'.format(start))
            # print('Finish = {}'.format(finish))

            filt = []
            delta = [start + timedelta(days=x) for x in range(0, (finish-start).days+1)]
            for day in delta:
                filt.append(day.strftime('%d %b %y'))
            print('date range filtered = {}'.format(filt))
            df_fil = df.loc[df['Date'].isin(filt)]
            # print(df_fil)
            ax1 = self.figure1.add_subplot(111)
            ax1.clear()

            df_bar = df_fil.groupby(['Date','Project'])['Delta'].sum().unstack()
            df_bar.plot(kind='bar',stacked=True, ax=ax1)    

            for tick in ax1.get_xticklabels():
                tick.set_rotation(90)
        
            leg= ax1.legend(prop={'size':8}, loc='center left', bbox_to_anchor=(1,0.5))
            if leg:
                leg.set_draggable(True)

            ax1.set_ylabel('Time')

            #Add major and Minor gridlines
            # Customize the major grid
            ax1.grid(which='major', linestyle=':', linewidth='0.5', color='grey')
            # Customize the minor grid
            ax1.grid(which='minor', linestyle=':', linewidth='0.5', color='grey')

            self.figure1.tight_layout()

            plt.isinteractive()      

            self.canvas1.draw()

        #============================================================================================================     
    def plotBar (self, df):
        ax1 = self.figure1.add_subplot(111)
        ax1.clear()

        fr = self.comboBox.currentText().strip()
        unt = self.comboBox_2.currentText().strip()

        if fr != '':
            start = datetime.strptime(str(fr), '%d %b %y').date()

        if unt != '':
            finish = datetime.strptime(str(unt), '%d %b %y').date()
            
            # print('Start = {}'.format(start))
            # print('Finish = {}'.format(finish))

            filt = []
            delta = [start + timedelta(days=x) for x in range(0, (finish-start).days+1)]
            for day in delta:
                filt.append(day.strftime('%d %b %y'))
            df_fil = df.loc[df['Date'].isin(filt)]
            
            df_bar = df_fil.groupby(['Date','Project'])['Delta'].sum().unstack()
            df_bar.plot(kind='bar',stacked=True, ax=ax1)

            for tick in ax1.get_xticklabels():
                tick.set_rotation(90)
            
            leg= ax1.legend(prop={'size':8}, loc='center left', bbox_to_anchor=(1,0.5))
            if leg:
                leg.set_draggable(True)

            ax1.set_ylabel('Time')

            #Add major and Minor gridlines
            # Customize the major grid
            ax1.grid(which='major', linestyle=':', linewidth='0.5', color='grey')
            # Customize the minor grid
            ax1.grid(which='minor', linestyle=':', linewidth='0.5', color='grey')

            self.figure1.tight_layout()

            plt.isinteractive()      

            self.canvas1.draw()
        else:
            # pass

            # df = df.loc['Date'].isin([filt])
            # df[df['Date'].isin(filt)]
            # print(df)
    
            df_bar = df.groupby(['Date','Project'])['Delta'].sum().unstack()
            df_bar.plot(kind='bar',stacked=True, ax=ax1)
    
            for tick in ax1.get_xticklabels():
                tick.set_rotation(90)
    
            leg= ax1.legend(prop={'size':8}, loc='center left', bbox_to_anchor=(1,0.5))
            if leg:
                leg.set_draggable(True)
    
            ax1.set_ylabel('Time')
    
            #Add major and Minor gridlines
            # Customize the major grid
            ax1.grid(which='major', linestyle=':', linewidth='0.5', color='grey')
    
            # Customize the minor grid
            ax1.grid(which='minor', linestyle=':', linewidth='0.5', color='grey')
    
            self.figure1.tight_layout()
    
            plt.isinteractive()      
    
            self.canvas1.draw()

        #============================================================================================================     
    def plotPie (self, df):
        ax2 = self.figure2.add_subplot(111)
        ax2.clear()

        today = datetime.now().strftime('%d %b %y')

        df_pie = df.loc[df['Date'].isin([today])]
        # print(df_pie)
        df_pie.plot(kind='pie', y = 'Delta', ax=ax2, autopct='%1.1f%%',
                    startangle=0, shadow=False,labels=df_pie['Project'])
        ax2.axis('equal')
        
        leg= ax2.legend(prop={'size':8}, loc='center left', bbox_to_anchor=(1,0.5))
        if leg:
            leg.set_draggable(True)
            
        self.figure2.tight_layout()
        self.canvas2.draw()

        #============================================================================================================
    def closeSession (self):
        choice = QMessageBox.question(self, 'Close',
                                            "Do you want to quit the application?",
                                            QMessageBox.Yes | QMessageBox.No)
        if choice == QMessageBox.Yes:
            msg = QMessageBox.question(self, 'Save Session',
                                       "Do you want to save the session before quiting?",
                                       QMessageBox.Yes | QMessageBox.No)
            if msg == QMessageBox.Yes:
                self.saveSession()
                print("Session closed")
                sys.exit()
            if msg == QMessageBox.No:
                print("Session closed")
                sys.exit()
        if choice == QMessageBox.No:
            pass

        #============================================================================================================        
    def Print (self):
        filename,_ = QFileDialog.getSaveFileName(self, "Save plots as PDF file", (os.path.join(os.path.join(os.path.expanduser('~')), 'Documents', 'Tracker')), "Portable Document File (*.pdf)")
        if filename == "":
            return
        self.canvas1.print_figure(filename)

        #============================================================================================================        
    def saveSession(self):
        self.fname, _ = QFileDialog.getSaveFileName(self, 'Select save session name', (os.path.join(os.path.join(os.path.expanduser('~')), 'Documents', 'Tracker')), 'INI files (*.ini)')
        self.settings = QSettings(self.fname, QSettings.IniFormat)
        
        for name, obj in inspect.getmembers(self):
            if isinstance(obj, QLineEdit):
                name = obj.objectName()
                value = obj.text()
                self.settings.setValue(name, value)
            if isinstance(obj, QCheckBox):
                name = obj.objectName()
                state = obj.isChecked()
                self.settings.setValue(name, state)
            if isinstance(obj, QGroupBox):
                name = obj.objectName()
                state = obj.isChecked()
                self.settings.setValue(name, state)
            if isinstance(obj, QTableWidget):
                self.settings.setValue("rowCount", obj.rowCount())
                self.settings.setValue("columnCount", obj.columnCount())
                items = QtCore.QByteArray()
                stream = QtCore.QDataStream(items, QtCore.QIODevice.WriteOnly)
                for i in range(obj.rowCount()):
                    for j in range(obj.columnCount()):
                        it = obj.item(i,j)
                        if it is not None:
                            stream.writeInt(i)
                            stream.writeInt(j)
                            stream << it
                self.settings.setValue("items", items)
                selecteditems = QtCore.QByteArray()
                stream = QtCore.QDataStream(selecteditems, QtCore.QIODevice.WriteOnly)
                for it in obj.selectedItems():
                    stream.writeInt(it.row())
                    stream.writeInt(it.column())
                self.settings.setValue("selecteditems", selecteditems)
#        print(self.fname)

    #read in previous session links and settings
    def restoreSession(self):
        fname, _ = QFileDialog.getOpenFileName(self, 'Select session to restore', (os.path.join(os.path.join(os.path.expanduser('~')), 'Documents', 'Tracker')), 'INI files (*.ini)')
        if not os.path.exists(fname):
            return
        self.settings = QSettings(fname, QSettings.IniFormat)
        for name, obj in inspect.getmembers(self):
            if isinstance(obj, QLineEdit):
                name = obj.objectName()
                value = (self.settings.value(name))
                obj.setText(value)  # restore
            # if isinstance(obj, QCheckBox):
            #     name = obj.objectName()
            #     value = self.settings.value(name).lower()
            #     if value != None:
            #         obj.setChecked(strtobool(value))
            #     else:
            #         continue
            # if isinstance(obj, QGroupBox):
            #     name = obj.objectName()
            #     value = self.settings.value(name).lower()
            #     if value != None:
            #         obj.setChecked(strtobool(value))
            #     else:
            #         continue
            if isinstance(obj, QTableWidget):
                rowCount = self.settings.value("rowCount", type=int)
                columnCount = self.settings.value("columnCount", type=int)
                obj.setRowCount(rowCount)
                obj.setColumnCount(columnCount)
                items = self.settings.value("items")
                if items is None:
                    continue
                else:
                    stream = QtCore.QDataStream(items, QtCore.QIODevice.ReadOnly)
                    while not stream.atEnd():
                        it = QtWidgets.QTableWidgetItem()
                        i = stream.readInt()
                        j = stream.readInt()
                        stream >> it
                        obj.setItem(i, j, it)
                    selecteditems = self.settings.value("selecteditems")
                    stream = QtCore.QDataStream(selecteditems, QtCore.QIODevice.ReadOnly)
                    while not stream.atEnd():
                        i = stream.readInt()
                        j = stream.readInt()
                        it = obj.item(i, j)
                        if it is not None:
                            it.setSelected(True)
        self.combo()


        #============================================================================================================        
    def newSession (self):
        #clear all fields to allow new data to be loaded - issue with clearing plot still though
        self.lineEditProject.clear()

        self.tableWidget.clear()
        self.tableWidget.setRowCount(0)
        self.tableWidget.setColumnCount(5)
        self.tableWidget.setHorizontalHeaderLabels(('Date', 'Project', 'Start', 'Finish', 'Delta')) 

        # self.tableWidget.setItem(0, 0, QTableWidgetItem('Project A'))
        self.combo()

        #============================================================================================================            
    def Remove (self):
        #Remove selected row from table
        self.tableWidget.removeRow(self.tableWidget.currentRow())

        #============================================================================================================
    def saveCSV (self):
        num_of_rows = self.tableWidget.rowCount()
        num_of_col = self.tableWidget.columnCount()
        headers = [str(self.tableWidget.horizontalHeaderItem(i).text())for i in range(num_of_col)]

        df_list = []
        for row in range(num_of_rows):
            df_list2= []
            for col in range(num_of_col):
                table_item = self.tableWidget.item(row, col)
                df_list2.append('' if table_item is None else str(table_item.text()))
            df_list.append(df_list2)

        df = pd.DataFrame(df_list, columns=headers)
        today = datetime.now().strftime('%d %b %y')

        df[['Start', 'Finish']]=df[['Start', 'Finish']].apply(pd.to_timedelta)
        df['Date']=df['Date'].astype('Datetime64').dt.strftime("%d %b %y")
        df['Delta']=df.apply(lambda row: round((row.Finish - row.Start).seconds/3600,2), axis=1)

        #Save csv file to allow reload/reuse
        filename,_ = QFileDialog.getSaveFileName(self, "Save dataframe", (os.path.join(os.path.join(os.path.expanduser('~')), 'Documents', 'Tracker')), "CSV files (*.csv)")

        df.to_csv(filename, index=False)

        df2 = df.loc[df['Date'].isin([today])]

        fr = self.comboBox.currentText().strip()
        unt = self.comboBox_2.currentText().strip()

        if fr != '':
            start = datetime.strptime(str(fr), '%d %b %y').date()
        if unt != '':
            finish = datetime.strptime(str(unt), '%d %b %y').date()

            filt = []
            delta = [start + timedelta(days=x) for x in range(0, (finish-start).days+1)]
            for day in delta:
                filt.append(day.strftime('%d %b %y'))

            df3 = df.loc[df['Date'].isin(filt)]        

        # Create a Pandas Excel writer using XlsxWriter as the engine.
        writer = pd.ExcelWriter((os.path.splitext(filename)[0]+'.xlsx'), engine='xlsxwriter')

        # Write each dataframe to a different worksheet.
        df.to_excel(writer, sheet_name='Original DF')
        df2.to_excel(writer, sheet_name='Today DF')
        df3.to_excel(writer, sheet_name='Filtered DF')

        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        #============================================================================================================
if __name__=="__main__":
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = Timerwindow()
    MainWindow.show ()
    sys.exit(app.exec_())