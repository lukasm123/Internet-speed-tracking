import pyspeedtest
import datetime
import time
from xlwt import Workbook

class internet_test:

    #See pyspeedtest documentation for more information
    st = pyspeedtest.SpeedTest()
    #Needed for loop
    row = 1

    def testing(self, minutes):

        #Workbook creation
        wb = Workbook()
        sheet1 = wb.add_sheet('Sheet 1')

        #Naming colums
        sheet1.write(0, 0, 'ID')
        sheet1.write(0, 1, 'Speed in Bytes')
        sheet1.write(0, 2, 'Speed in Mbit')
        sheet1.write(0, 3, 'time')

        #Testing speed and writing into file
        number_of_testings = int(minutes/2)
        for i in range(number_of_testings):

            #Writing ID in file
            sheet1.write(self.row, 0, i+1)

            #Testing speed and writing in file
            result = self.st.download()
            sheet1.write(self.row, 1, result)

            #Calculating Mbit and writing in file
            mbit = round(result/1000000, 2)
            sheet1.write(self.row, 2, mbit)

            #Writing daytime
            sheet1.write(self.row, 3, datetime.datetime.now())

            #Log
            print("ID: " + str(i+1) + " " + str(result) + " Bytes / " + str(mbit) + " Mbit @ " + str(datetime.datetime.now()) + ". Added to sheet.")

            #Take next row
            self.row += 1

            #Sleeping
            time.sleep(30.0)

        #Saving data
        wb.save('speedtest_data.xls')

        print("finished")


while __name__ == '__main__':
    internet_test().testing(60)
    exit()