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
      #print(str(serialNumberToTest) + " does not equal " + str(serialFromNewWorkbook))
      # if y == updateXlListEnd - 1:
      #   print(str(serialNumberToTest) + ' Was not found ' + str(y))
      continue
    else:
      # print(str(serialNumberToTest) + " matches " + str(serialFromNewWorkbook))
      numMatches = numMatches + 1
    
      #new
      newCost5 = newSheet.cell_value(y,18) # S
      newOutCost = newSheet.cell_value(y,19) # T
      newMSRP = newSheet.cell_value(y,20) # U
      newSpecial1 = newSheet.cell_value(y,31)
      newSpecial2 = newSheet.cell_value(y,32)

      #old
      itemType = sheet.cell_value(x,0)
      partNumber = sheet.cell_value(x,4)
      model = sheet.cell_value(x,15)
      cost5 = sheet.cell_value(x,18)
      outCost = sheet.cell_value(x,19)
      MSRP = sheet.cell_value(x,20)
      special1 = sheet.cell_value(x,31)
      special2 = sheet.cell_value(x,32)
      desc = sheet.cell_value(x,30)
      
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
        wst.write(modelStart, 0, itemType)
        wst.write(modelStart, 3, thisMapp)
        wst.write(modelStart, 11, serialNumberToTest)
        if itemType == "Model":
          wst.write(modelStart, 4, partNumber)
          wst.write(modelStart, 5, "Equipment")
          wst.write(modelStart, 6, category)
          wst.write(modelStart, 28, "Document Direction Ltd")
        if itemType == "Access":
          wst.write(modelStart, 8, "ACCESSORY") # "I"
          wst.write(modelStart, 9, "N") #"J"
        wst.write(modelStart, 13, model) #"N"
        wst.write(modelStart, 14, 0)
        wst.write(modelStart, 15, 0)
        wst.write(modelStart, 16, 0)
        wst.write(modelStart, 17, 0)
        wst.write(modelStart, 19, newCost5) #"S"
        wst.write(modelStart, 20, outCost) #"T"
        wst.write(modelStart, 21, MSRP) # "U"
        wst.write(modelStart, 25, 0)
        wst.write(modelStart, 29, desc)
        wst.write(modelStart, 30, special1)
        wst.write(modelStart, 32, special2)
        for l in range(0,38):
            wst.write(modelStart,33+l, 0)

        modelStart = modelStart + 1
        newPrice = False

wbt.save("PriceChanges.xls")
print("matches: " + str(numMatches))
    # wst.write(modelStart, 3, thisMapp)
    # wst.write(modelStart, 4, modelList[i]["name"])
    # wst.write(modelStart, 5, "Equipment")
    # wst.write(modelStart, 6,  category)
    # wst.write(modelStart, 11, modelList[i]["productNumber"])
    # wst.write(modelStart, 13, modelList[i]["name"])
    # wst.write(modelStart, 14, 0)
    # wst.write(modelStart, 15, 0)
    # wst.write(modelStart, 16, 0)
    # wst.write(modelStart, 17, 0)
    # wst.write(modelStart, 18, modelList[i]["mapp"])
    # wst.write(modelStart, 19, modelList[i]["mapp"])
    # wst.write(modelStart, 20, modelList[i]["msrp"])
    # wst.write(modelStart, 28, "Document Direction Ltd")
    # wst.write(modelStart, 29, modelList[i]["desc"])
    # wst.write(modelStart, 31, modelList[i]["rmapp"])
    # wst.write(modelStart, 32, modelList[i]["rmapp2"])
    # # write a bunch of zeros in the special pricing fields
    # for j in range(0,38):
    #   wst.write(modelStart,33+j, 0)
    #     # write in accessories
    #     for k in range(0, len(globalList)):
    #       wst.write(globalStart, 0, "Access")
    #       wst.write(globalStart, 3, thisMapp)
    #       wst.write(globalStart, 8, "ACCESSORY")
    #       wst.write(globalStart, 9, "N")
    #       wst.write(globalStart, 11, globalList[k]["productNumber"])
    #       wst.write(globalStart, 13, globalList[k]["name"])
    #       wst.write(globalStart, 14, 0)
    #       wst.write(globalStart, 15, 0)
    #       wst.write(globalStart, 16, 0)
    #       wst.write(globalStart, 17, 0)
    #       wst.write(globalStart, 18, globalList[k]["mapp"])
    #       wst.write(globalStart, 19, globalList[k]["mapp"])
    #       wst.write(globalStart, 20, globalList[k]["msrp"])
    #       wst.write(globalStart, 25, 0)
    #       wst.write(globalStart, 29, globalList[k]["desc"])
    #       wst.write(globalStart, 31, globalList[k]["rmapp"])
    #       wst.write(globalStart, 32, globalList[k]["rmapp2"])
    #       # write a bunch of zeros in the special pricing fields
    #       for l in range(0,38):
    #         wst.write(globalStart,33+l, 0)

    #       globalStart = globalStart + 1

    #     accStart = globalStart
    #     # write in accessories
    #     for j in range(0, len(accList)):
    #       wst.write(accStart, 0, "Access")
    #       wst.write(accStart, 3, thisMapp)
    #       wst.write(accStart, 8, "ACCESSORY")
    #       wst.write(accStart, 9, "N")
    #       wst.write(accStart, 11, accList[j]["productNumber"])
    #       wst.write(accStart, 13, accList[j]["name"])
    #       wst.write(accStart, 14, 0)
    #       wst.write(accStart, 15, 0)
    #       wst.write(accStart, 16, 0)
    #       wst.write(accStart, 17, 0)
    #       wst.write(accStart, 18, accList[j]["mapp"])
    #       wst.write(accStart, 19, accList[j]["mapp"])
    #       wst.write(accStart, 20, accList[j]["msrp"])
    #       wst.write(accStart, 25, 0)
    #       wst.write(accStart, 29, accList[j]["desc"])
    #       wst.write(accStart, 31, accList[j]["rmapp"])
    #       wst.write(accStart, 32, accList[j]["rmapp2"])
    #       # write a bunch of zeros in the special pricing fields
    #       for m in range(0,38):
    #         wst.write(accStart,33+m, 0)

    #       accStart = accStart + 1
        
    #     globalStart = accStart
    #     modelStart = accStart




