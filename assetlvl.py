import math
import xlrd
import xlwt

WellsGoodFile = False
while WellsGoodFile == False:
  fileToRead = input("Please enter the filename of the Wells portfolio (xlsx only)> ")
  if fileToRead == "":
    loc = ("WellsPortfolio.xlsx")
    WellsGoodFile = True
    wb = xlrd.open_workbook(loc)
  elif fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    loc = fileToRead + ".xlsx"
    try:
      WellsGoodFile = True
      wb = xlrd.open_workbook(loc)
    except:
      print("I can't find that file, try again...")

#open wells workbook
sheet = wb.sheet_by_index(0)

DLLgoodFile = False
while DLLgoodFile == False:
  fileToRead = input("Please enter the filename of the DLL portfolio (xlsx only)> ")
  if fileToRead == "":
    loc2 = ("DLLPortfolio.xlsx")
    DLLgoodFile = True
    wb2 = xlrd.open_workbook(loc2)
  elif fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    loc2 = fileToRead + ".xlsx"
    try:
      wb2 = xlrd.open_workbook(loc2)
      DLLgoodFile = True
    except:
      print("I can't find that file, try again...")

#open DLL workbook
sheet2 = wb2.sheet_by_index(0)

SRgoodFile = False
while SRgoodFile == False:
  fileToRead = input("Please enter the filename of the Sherpa report (make sure 'raw data' is in position 0)> ")
  if fileToRead == "":
    loc3 = ("SherpaReport.xlsm")
    SRgoodFile = True
    wb3 = xlrd.open_workbook(loc3)
  elif fileToRead == "exit" or fileToRead == "quit":
    print("ok, bye!")
    exit()
  else:
    loc3 = fileToRead + ".xlsm"
    try:
      wb3 = xlrd.open_workbook(loc3)
      SRgoodFile = True
    except:
      print("I can't find that file, try again...")

#Open Sherpa report
sheet3 = wb3.sheet_by_index(0)

#write to workbook
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('AssetLevelAdded')
NewWorkbookName = "myNewWb.xls"

# end of excel sheet (report for perry)
endOfSherpaReport = 3569
endOfWells = 2222
endOfDLL = 1489
DLLAssetPriceColumn = 21
DLLSerialColumn = 22
WellsAssetPriceColumn = 4
WellsSerialColumn = 2

for x in range(1,endOfSherpaReport):
  # get serial to test
  try:
    testSerial = sheet3.cell_value(x,10)
    testSerial = int(testSerial)
  except:
    testSerial = str(testSerial)

  # look in the wells portfolio for the serial
  for y in range(1, endOfWells):
    try:
      wellsAssetPrice = sheet.cell_value(y, WellsAssetPriceColumn)
      wellsSerial = sheet.cell_value(y, WellsSerialColumn)
    except:
      continue

    if testSerial == "":
      break
    if testSerial == wellsSerial:
      try:
        worksheet.write(x,0, testSerial)
        worksheet.write(x,1, wellsAssetPrice)
      except:
        continue
      break

  # look in the DLL portfolio for the serial
  for y in range(1, endOfDLL):
    try:
      DLLAssetPrice = sheet2.cell_value(y,DLLAssetPriceColumn)
      DLLserial = sheet2.cell_value(y,DLLSerialColumn)
    except:
      continue

    if testSerial == "":
      break
    if testSerial == DLLserial:
      try:
        worksheet.write(x,0, testSerial)
        worksheet.write(x,1, DLLAssetPrice)
      except:
        continue
      break


workbook.save(NewWorkbookName)
print("saved: " + str(NewWorkbookName))