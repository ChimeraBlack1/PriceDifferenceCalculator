import math
import xlrd
import xlwt

#open LSAP Rewnewal workbook
loc = ("OldUpdate.xls")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

newUpdate = ("NewUpdate.xls")
wb2 = xlrd.open_workbook(newUpdate)
newSheet = wb2.sheet_by_index(0)

#globals
newPrice = 0
xlListEnd = 10

# Check every item number's price against its old price
# if the price is different, write THE ENTIRE CONFIG to a new wb

for x in range(0, xlListEnd):

  # serial numbers to test
  serialNumberToTest = sheet.cell_value(x, 11)
  serialFromNewWorkbook = sheet.cell_value(x,11)

  if serialNumberToTest != serialFromNewWorkbook:
    continue

  #old costs
  cost = sheet.cell_value(x,14)
  cost2 = sheet.cell_value(x,15)
  cost3 = sheet.cell_value(x,16)
  cost4 = sheet.cell_value(x,17)
  cost5 = sheet.cell_value(x,18)
  outCost = sheet.cell_value(x,19)
  MSRP = sheet.cell_value(x,20)
  special1 = sheet.cell_value(x,31)
  special2 = sheet.cell_value(x,32)

  #new costs
  newCost = newSheet.cell_value(x,14)
  newCost2 = newSheet.cell_value(x,15)
  newCost3 = newSheet.cell_value(x,16)
  newCost4 = newSheet.cell_value(x,17)
  newCost5 = newSheet.cell_value(x,18)
  newOutCost = newSheet.cell_value(x,19)
  newMSRP = newSheet.cell_value(x,20)
  newSpecial1 = newSheet.cell_value(x,31)
  newSpecial2 = newSheet.cell_value(x,32)

  if serialNumberToTest == serialFromNewWorkbook:

    if cost != newCost:
      print("old cost: " + str(cost) + " new cost:" + str(newCost))
      newPrice = newPrice + 1 
    if cost2 != newCost2:
      print("old cost2: " + str(cost2) + " new newCost2:" + str(newCost2))
      newPrice = newPrice + 1 
    if cost3 != newCost3:
      print("old cost3: " + str(cost3) + " new newCost3:" + str(newCost3))
      newPrice = newPrice + 1 
    if cost4 != newCost4:
      print("old cost4: " + str(cost4) + " new newCost4:" + str(newCost4))
      newPrice = newPrice + 1 
    if cost5 != newCost5:
      print("old cost5: " + str(cost5) + " new newCost5:" + str(newCost5))
      newPrice = newPrice + 1 
    if outCost != newOutCost:
      print("old outCost: " + str(outCost) + " new newOutCost:" + str(newOutCost))
      newPrice = newPrice + 1 
    if MSRP != newMSRP:
      print("old MSRP: " + str(MSRP) + " new newMSRP:" + str(newMSRP))
      newPrice = newPrice + 1 
    if special1 != newSpecial1:
      print("old special1: " + str(special1) + " new newSpecial1:" + str(newSpecial1))  
      newPrice = newPrice + 1 
    if special2 != newSpecial2:
      print("old special2: " + str(special2) + " new newSpecial2:" + str(newSpecial2))
      newPrice = newPrice + 1

  if newPrice > 0:
    print("price has changed!")
    newPrice = 0
  else:
    newPrice = 0
