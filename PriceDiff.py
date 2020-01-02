import math
import xlrd
import xlwt


##################################
#          EXCEL SHEETS          #
##################################
loc = ("OctUpdate(Uniques).xls")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

newUpdate = ("NovUpdate(Uniques).xls")
wb2 = xlrd.open_workbook(newUpdate)
newSheet = wb2.sheet_by_index(0)

#write to workbook
wbt = xlwt.Workbook()
wst = wbt.add_sheet('Prices that changed')

# Check every item number's price against its old price
# if the price is different, write THE ENTIRE CONFIG to a new wb


##################################
#             METHODS            #
##################################
#Get Prices
def GetPrices():
  newPrices = {
    "newCost5": newSheet.cell_value(y,18), # S
    "newOutCost": newSheet.cell_value(y,19), # T
    "newMSRP": newSheet.cell_value(y,20), # U
    "newSpecial1": newSheet.cell_value(y,31), #AF
    "newSpecial2": newSheet.cell_value(y,32), #AG
    "newDesc": newSheet.cell_value(x,29),
  }

  oldPrices = {
    "itemType": sheet.cell_value(x,0), #A
    "partNumber": sheet.cell_value(x,4), #E
    "model": sheet.cell_value(x,15), #P
    "cost5": sheet.cell_value(x,18), #S
    "outCost": sheet.cell_value(x,19), #T
    "MSRP": sheet.cell_value(x,20), #U
    "special1": sheet.cell_value(x,31), #AF
    "special2": sheet.cell_value(x,32), #AG
    "desc": sheet.cell_value(x,29), #AD
  }

  return newPrices, oldPrices


def ComparePrices(prices):
  newPrice = False
  if prices[1]["cost5"] != prices[0]["newCost5"]:
    print("old cost5: " + str(prices[1]["cost5"]) + " new newCost5:" + str(prices[0]["newCost5"]))
    newPrice = True
  if prices[1]["outCost"] != prices[0]["newOutCost"]:
    print("old outCost: " + str(prices[1]["outCost"]) + " new newOutCost:" + str(prices[0]["newOutCost"]))
    newPrice = True
  if prices[1]["MSRP"] != prices[0]["newMSRP"]:
    print("old MSRP: " + str(prices[1]["MSRP"]) + " new newMSRP:" + str(prices[0]["newMSRP"]))
    newPrice = True
  if prices[1]["special1"] != prices[0]["newSpecial1"]:
    print("old special1: " + str(prices[1]["special1"]) + " new newSpecial1:" + str(prices[0]["newSpecial1"]))
    newPrice = True
  if prices[1]["special2"] != prices[0]["newSpecial2"]:
    print("old special2: " + str(prices[1]["special2"]) + " new newSpecial2:" + str(prices[0]["newSpecial2"]))
    newPrice = True
  return newPrice


def WriteNewPrices(modelStart, newPrice, serialNumberToTest, prices):
  #TODO Which Mapp?
  thisMapp = "Ricoh HW MAPP"
  category = "Hardware"

  if newPrice == True:
    print(str(prices))
    print("the price for " + str(serialNumberToTest) + " has changed")
    print("..............................................")
    wst.write(modelStart, 0, prices[1]["itemType"]) #A
    wst.write(modelStart, 3, thisMapp) #D
    wst.write(modelStart, 11, serialNumberToTest) # L
    if prices[1]["itemType"] == "Model":
      wst.write(modelStart, 4,  prices[1]["partNumber"]) #E
      wst.write(modelStart, 5, "Equipment") #F
      wst.write(modelStart, 6, category) #G
      wst.write(modelStart, 28, "Document Direction Ltd") #AC
    if prices[1]["itemType"] == "Access":
      wst.write(modelStart, 8, "ACCESSORY") # "I"
      wst.write(modelStart, 9, "N") #"J"
    wst.write(modelStart, 13, prices[1]["model"]) #"N"
    wst.write(modelStart, 14, 0) #O
    wst.write(modelStart, 15, 0) #P
    wst.write(modelStart, 16, 0) #Q
    wst.write(modelStart, 17, 0) #R
    wst.write(modelStart, 18, prices[0]["newCost5"]) #"S"
    wst.write(modelStart, 19, prices[0]["newOutCost"]) #"T"
    wst.write(modelStart, 20, prices[0]["newMSRP"]) # "U"
    wst.write(modelStart, 25, 0) #Z
    wst.write(modelStart, 29, prices[0]["newDesc"]) #AD
    wst.write(modelStart, 31, prices[0]["newSpecial1"]) #AE
    wst.write(modelStart, 32, prices[0]["newSpecial2"]) #AG
    for l in range(0,38):
        wst.write(modelStart,33+l, 0) 
    return modelStart

def SaveWork():
  wbt.save("PriceChanges.xls")
  print("matches: " + str(numMatches))


##################################
#             GLOBALS            #
##################################
xlListEnd = 735
updateXlListEnd = 601
numMatches = 0
modelStart = 0


##################################
#             LOGIC            #
##################################
for x in range(0, xlListEnd):
  # serial numbers to test
  serialNumberToTest = sheet.cell_value(x, 11)

  for y in range(0, updateXlListEnd):
    serialFromNewWorkbook = newSheet.cell_value(y,11)

    if serialNumberToTest != serialFromNewWorkbook:
      continue
    else:
      numMatches = numMatches + 1

      prices = GetPrices()
      newPrice = ComparePrices(prices)
      WriteNewPrices(modelStart, newPrice, serialNumberToTest, prices)
      modelStart = modelStart + 1

SaveWork()