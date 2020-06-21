from openpyxl import Workbook
from datetime import datetime
from os import system
from sys import version_info
import subprocess
from random import uniform # ! returns floating-point value, use randint for integer value

class Randomize_App:
    def __init__(self, start_int = 0, end_int = 0, result = 0, attempts = 0):
        if version_info < (3, 5, 0):
            raise NotImplementedError("Cannot run this module from Python Version that is below than 3.5.0. Please Update your Python Version!")
        else:
            self.start_int = start_int
            self.end_int = end_int
            self.result = result
            self.attempts = attempts
            self.XLSX_Worker = Workbook()
            self.Active_XLSX = self.XLSX_Worker.active
            self.Datetime_Create = datetime.now()

    def Welcome_Intro(self):
        subprocess.call("CLS", shell=True)
        print("Hello and Welcome To Randomize Cutout Generator")
        print("Exclusively Created for EDA which uses bond paper to do data.")
        print("Data will be automatically added to Excel... Choice the data at your own sake.\n")
        self.Active_XLSX.title = input('Please enter your Filename -> ')
        return

    def Intro_DataGet(self):
        print("Data Needed is 30 Attempts of Bond Paper Cutting")
        print("Short Bond Paper Dimension is 8.5 x 11 which this concludes the possible range to cut from 2.54cm to 27.94cm")
        print("Middle Cut is 13.97cm which should give a hint on where to start...")
        print("Please specify range cm to generate...")
        try:
            self.start_int = float(input("Input Starting Range in CM -> "))
            self.end_int = float(input("Input Starting Range in CM -> "))
            self.attempts = int(input("Input Retries, Will Represent as the Nth Column - > "))
            print("Generating Data...\n")
            return
        except ValueError:
            self.start_int = 0
            self.end_int = 0
            raise ValueError("User Inserted a Non Integer Value. Please restart the script to try again.")

    def GenerateData(self):
        ColumnCount = 0
        for Column_nth in range(1, self.attempts + 1):
            ColumnCount += 1
            for Row_nth in range(1, 31):
                self.result = round(uniform(self.start_int, self.end_int), 1)
                self.Active_XLSX.cell(column = ColumnCount, row = Row_nth, value = (str(self.result) + ' cm'))
                print('Column #',Column_nth,'[ Column Position -> ',ColumnCount,'], Returned Output {}'.format(self.result), 'cm @ ', Row_nth)
        self.XLSX_Worker.save(self.Active_XLSX.title + '.xlsx')
        print("Data Creation is Finished... Please Check", self.Active_XLSX.title + '.xlsx on the location of this script... Thank you!')

if __name__ == "__main__":
    try:
        Instance_1 = Randomize_App()
        Instance_1.Welcome_Intro()
        Instance_1.Intro_DataGet()
        Instance_1.GenerateData()

    except BaseException as ScriptErr:
        print('Error Running Script: %s' % (ScriptErr,))
