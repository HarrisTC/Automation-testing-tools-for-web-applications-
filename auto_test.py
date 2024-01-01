import sys, time, threading, pyautogui, openpyxl, os, win32com.client 
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QLineEdit, QFileDialog, QTextBrowser
from PyQt5 import uic
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime

class UI(QMainWindow):
    # ------------- Main Window -----------------
    def __init__(self):
        super().__init__()
        # Load the ui file 
        uic.loadUi("Automation.ui", self)

        # Define Widgets 
        self.openButton  = self.findChild(QPushButton , "Open")
        self.startButton = self.findChild(QPushButton , "start_Button")
        self.stopButton  = self.findChild(QPushButton , "stop_Button")
        self.filePath    = self.findChild(QLineEdit   , "lineEdit")
        self.output      = self.findChild(QTextBrowser, "output_text")
        self.clearbutton = self.findChild(QPushButton , "clear_Button")

        # Click the import file
        self.openButton.released.connect(self.clicker)
        self.startButton.released.connect(self.start)
        self.clearbutton.released.connect(self.clearOutput)

        # Show the App
        self.show()

    ##### Click Path Browser
    def clicker(self):
        # Open File dialog
        fname = QFileDialog.getOpenFileName(self, "Open File", "", "All Files (*);;Excel Files (*.xlsx)")
        
        # Output file name
        if fname:
            self.filePath.setText(str(fname[0]))

    ##### Click Start Button
    def start(self):
        dic_Config = {}
        self.output.append("------------Starting--------------")
        self.close_all_excel_files()
        try:
            self.output.append(str(datetime.now())+ " : Read file " + self.filePath.text() )
        except:
            self.output.append(str(datetime.now())+ " : Failed to read file " + self.filePath.text() +"\n")
            return
        # Read Config sheet
        try:
            Config_table = pd.read_excel(self.filePath.text(),'Config')
            for _,row in Config_table.iterrows():
                dic_Config[row['Name']] = row['Type']
            self.output.append(str(datetime.now())+ " : Read Config Sheet Success. " )
        except:
            self.output.append(str(datetime.now())+ " : " +"Can't Read Config Sheet."+"\n")
            return

        # Read Summary sheet
        try:    
            Summary_table = pd.read_excel(self.filePath.text(),'Summary')
            self.output.append(str(datetime.now())+ " : Read Summary Sheet Success. " )
        except:
            self.output.append(str(datetime.now())+ " : " +"Can't Read Summary Sheet."+"\n")
            return

        driver = webdriver.Edge()
        # Read Test Suites
        for index, row in Summary_table.iterrows():
            if str(row['Run']) == 'Run':
                # Read Test Case Sheet Names
                try:
                    testSuite_table = pd.read_excel(self.filePath.text(),str(row['Test Suite']) )
                    self.output.append(str(datetime.now())+ " : " + "Start - Read Test Suite " + str(row['Test Suite']) + " . " )
                except:
                    self.screenshot()
                    self.Excel_Write_Cell(self.filePath.text(),'Summary','F'+str(index+2),'NG')
                    self.Excel_Write_Cell(self.filePath.text(),'Summary','G'+str(index+2),"Read Test Suite " + str(row['Test Suite']) + " Failed. ")
                    self.output.append(str(datetime.now())+ " : " + "Read Test Suite " + str(row['Test Suite']) + " Failed. "+"\n")
                    return
                for _, row_testSuite in testSuite_table.iterrows():
                    # Read Activities in Test Case Sheet
                    try:
                        testCase_table = pd.read_excel(self.filePath.text(),str(row_testSuite['Test Case']) )
                        self.output.append(str(datetime.now())+ " : " + "Start - Read Test Case " + str(row_testSuite['Test Case']) + ". " )
                    except:
                        self.screenshot()
                        self.Excel_Write_Cell(self.filePath.text(),'Summary','F'+str(index+2),'NG')
                        self.Excel_Write_Cell(self.filePath.text(),'Summary','G'+str(index+2),"Read Test Case " + str(row_testSuite['Test Case']) + " Failed. ")
                        self.output.append(str(datetime.now())+ " : " + "Read Test Case " + str(row_testSuite['Test Case']) + " Failed. "+"\n" )
                        return
                    for _, row_testCase in testCase_table.iterrows():
                        try:
                            match dic_Config[row_testCase['Action']]: 
                                case 'Xpath':
                                    self.xpath_Activity(driver,row_testCase['Action'],row_testCase['Xpath'])
                                case 'Input and Xpath':
                                    res = self.xpath_input_Activity(driver,row_testCase['Action'],row_testCase['Xpath'],row_testCase['Input'])
                                    if res == 1:
                                        self.screenshot()
                                        self.Excel_Write_Cell(self.filePath.text(),'Summary','F'+str(index+2),'NG')
                                        self.Excel_Write_Cell(self.filePath.text(),'Summary','G'+str(index+2),str(row['Test Suite'])+" - "+str(row_testSuite['Test Case']) + " - No "+str(row_testCase['No']) + " - Failed")
                                        return self.output.append(str(datetime.now())+ " : " +str(row['Test Suite'])+" - "+row_testSuite['Test Case'] + " - No "+row_testCase['No'] + " Failed."+"\n")
                                case 'Input':
                                    self.input_Activity(driver,row_testCase['Action'],row_testCase['Input'])
                        except:
                            self.screenshot()
                            self.Excel_Write_Cell(self.filePath.text(),'Summary','F'+str(index+2),'NG')
                            self.Excel_Write_Cell(self.filePath.text(),'Summary','G'+str(index+2),str(row['Test Suite'])+" - "+str(row_testSuite['Test Case']) + " - No "+str(row_testCase['No']) + " - Failed")
                            self.output.append(str(datetime.now())+ " : " +str(row['Test Suite'])+" - "+str(row_testSuite['Test Case']) + " - No "+str(row_testCase['No']) + " - Failed."+"\n")
                            return

    def clearOutput(self):
        self.output.clear()

    ##### Xpath Activity
    def xpath_Activity(self, driver, action, xpath):
        match action:
            case 'Click':
                driver.find_element(By.XPATH, xpath).click()

    ##### input Activity
    def input_Activity(self, driver, action, input):
        match action:
            case 'OpenBrowser':
                driver.get(input)
            case 'Delays':
                threading.Thread(target=self.delays(input)).start()
    
    ##### Xpath and input Activity
    def xpath_input_Activity(self, driver, action, xpath, input):
        match action:
            case 'TypeInto':
                driver.find_element(By.XPATH,xpath).send_keys(input)
            case 'VerifyText':
                if driver.find_element(By.XPATH,xpath).text != input:
                    return 1
    
    def delays(self,input):
        time.sleep(input)
        
    def screenshot(self):
        path = 'C:/Users/HOANGTC7/Desktop/error.png'
        image = pyautogui.screenshot() 
        image.save(os.path.abspath(path))

    def Excel_Write_Cell(self, excel_path, sheet_name, cell, value):
        wb = openpyxl.load_workbook(filename=excel_path)
        for ws in wb.worksheets:
            if sheet_name in str(ws):
                ws[cell] = value
                wb.save(excel_path)
                break

    def close_all_excel_files(self):
        self.output.append(str(datetime.now())+ " : " +"Closing all Excel Applications")
        excel = win32com.client.Dispatch("Excel.Application")
        # Close each workbook
        for workbook in excel.Workbooks:
            workbook.Close(SaveChanges=False)
        # Quit Excel application
        excel.Quit()
        

app = QApplication(sys.argv)
UIWindow = UI()
app.exec_()