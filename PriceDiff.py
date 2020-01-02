import math
import xlrd
import xlwt

loc = ("OctUpdate(Uniques).xls")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

newUpdate = ("NovUpdate(Uniques).xls")
wb2 = xlrd.open_workbook(newUpdate)
newSheet = wb2.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wst = wbt.add_sheet('Prices that changed')

#globals
newPrice = False
xlListEnd = 735
updateXlListEnd = 601
numMatches = 0
modelStart = 0

#TODO Which Mapp?
thisMapp = "Ricoh HW MAPP"
category = "Hardware"

# Check every item number's price against its old price
# if the price is different, write THE ENTIRE CONFIG to a new wb
for x in range(0, xlListEnd):
  # serial numbers to test
  serialNumberToTest = sheet.cell_value(x, 11)

  for y in range(0, updateXlListEnd):
    serialFromNewWorkbook = newSheet.cell_value(y,11)
    if serialNumberToTest != serialFromNewWorkbook:
      continue
    else:
      numMatches = numMatches + 1
    
      #new
      newCost5 = newSheet.cell_value(y,18) # S
      newOutCost = newSheet.cell_value(y,19) # T
      newMSRP = newSheet.cell_value(y,20) # U
      newSpecial1 = newSheet.cell_value(y,31) #AF
      newSpecial2 = newSheet.cell_value(y,32) #AG

      #old
      itemType = sheet.cell_value(x,0) #A
      partNumber = sheet.cell_value(x,4) #E
      model = sheet.cell_value(x,15) #P
      cost5 = sheet.cell_value(x,18) #S
      outCost = sheet.cell_value(x,19) #T
      MSRP = sheet.cell_value(x,20) #U
      special1 = sheet.cell_value(x,31) #AF
      special2 = sheet.cell_value(x,32) #AG
      desc = sheet.cell_value(x,29) #AD
      
      if cost5 != newCost5:
        print("old cost5: " + str(cost5) + " new newCost5:" + str(newCost5))
        newPrice = True
      if outCost != newOutCost:
        print("old outCost: " + str(outCost) + " new newOutCost:" + str(newOutCost))
        newPrice = True
      if MSRP != newMSRP:
        print("old MSRP: " + str(MSRP) + " new newMSRP:" + str(newMSRP))
        newPrice = True
      if special1 != newSpecial1:
        print("old special1: " + str(special1) + " new newSpecial1:" + str(newSpecial1))
        newPrice = True
      if special2 != newSpecial2:
        print("old special2: " + str(special2) + " new newSpecial2:" + str(newSpecial2))
        newPrice = True
      
      if newPrice == True:
        print("the price for " + str(serialNumberToTest) + " has changed")
        print("..............................................")
        wst.write(modelStart, 0, itemType) #A
        wst.write(modelStart, 3, thisMapp) #D
        wst.write(modelStart, 11, serialNumberToTest) # L
        if itemType == "Model":
          wst.write(modelStart, 4, partNumber) #E
          wst.write(modelStart, 5, "Equipment") #F
          wst.write(modelStart, 6, category) #G
          wst.write(modelStart, 28, "Document Direction Ltd") #AC
        if itemType == "Access":
          wst.write(modelStart, 8, "ACCESSORY") # "I"
          wst.write(modelStart, 9, "N") #"J"
        wst.write(modelStart, 13, model) #"N"
        wst.write(modelStart, 14, 0) #O
        wst.write(modelStart, 15, 0) #P
        wst.write(modelStart, 16, 0) #Q
        wst.write(modelStart, 17, 0) #R
        wst.write(modelStart, 18, newCost5) #"S"
        wst.write(modelStart, 19, outCost) #"T"
        wst.write(modelStart, 20, MSRP) # "U"
        wst.write(modelStart, 25, 0) #Z
        wst.write(modelStart, 29, desc) #AD
        wst.write(modelStart, 31, special1) #AE
        wst.write(modelStart, 32, special2) #AG
        for l in range(0,38):
            wst.write(modelStart,33+l, 0) 

        modelStart = modelStart + 1
        newPrice = False

wbt.save("PriceChanges.xls")
print("matches: " + str(numMatches))