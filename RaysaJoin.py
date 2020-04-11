import xlrd
import datetime

from Openorder import openorder

from dialog import getfiledir
from easygui import *


def openorderingest(filename):
    # open file
    xl_workbook = xlrd.open_workbook(filename)

    # Grab the sheet names
    sheet_names = xl_workbook.sheet_names()
    # print('Sheet Names', sheet_names)

    # Get the cost sheet in view
    try:
        sheet = xl_workbook.sheet_by_name("Open Order Detail")
    except:
        print("sheet Open order detail not found on: ", filename)

    # Print all values, iterating through rows and columns
    #
    rows = []
    line = ""
    num_cols = sheet.ncols  # Number of columns
    for row_idx in range(2, sheet.nrows):  # Iterate through rows
        # print('-' * 40)
        if (line == ""):
            line = ""
        else:
            rows.append(line)
            line = ""
            # print(line)
            line = ""
        # print('Row: %s' % row_idx)  # Print row number
        for col_idx in range(0, num_cols):  # Iterate through columns
            # entercase to screw with the data and convert it from excel stupidity to data types that are worth something
            if ((col_idx == 6) or (col_idx == 7)):
                if (sheet.cell(row_idx, col_idx).value != ''):
                    # print(sheet.cell_value(row_idx,col_idx))
                    cell_obj = xlrd.xldate.xldate_as_datetime(sheet.cell_value(row_idx, col_idx), 0)
                    # print(str(cell_obj))
            else:
                cell_obj = sheet.cell_value(row_idx, col_idx)  # Get cell object by row, col
            line = (line + "," + str(cell_obj))
            # print('Column: [%s] cell_obj: [%s]' % (col_idx, cell_obj))

    # for i in range(len(rows)):
    # print(rows[i])
    print("returning Rows")
    return rows
def logrow(foo):
    currentDT = datetime.datetime.now()
    temp =   ("C:/Users/cburt/PycharmProjects/CostSheetConsilidation/Logs/" + (currentDT.strftime("%Y-%m-%d"))+".csv")
    outFile = open(temp, 'a')
    outFile.write(foo+",")
    outFile.write("\n")
    outFile.close()

    print("finished logging")
    return 0

# abstract the data into objects that can be easily worked with
def optomizeaccess(rows):
    orderList = []
    seenordersitemcombo = set()



    for i in range(len(rows)):
        #print(rows[i])
        # print(i)
        try:
            trash, cust, cs, so, sol, po, cpo, ship, cancel, style, cstyle, qty, price, retail, postat, extprice= rows[
                i].split(",")
        except:
            print("Too many rows to unpack, row = ", rows[i])
            logrow(rows[i])
            continue
        try:
            cs = float(cs.strip())
            item = str(style.strip())+str(cs)
        except:
            cs = None
            #if blank costsheet skip that line
        if cs == None:
            continue
            #if seen, (put into seen order set so that i can have constant access)
        elif item in seenordersitemcombo:
            for j in range(len(orderList)):
                if orderList[j].costSheet == cs and orderList[j].styleNum == str(style.strip):
                    orderList[j].shipDate = ship
            #else, add a new object to the list
        else:
            x = openorder()
            x.shipDate = ship
            x.custCode = str(cust)
            x.costSheet = cs
            seenordersitemcombo.add(item)
            x.salesorder = so
            x.poNum = po
            x.styleNum = str(style.strip())
            x.qty = qty
            x.price = float(price)
            x.retail = float(retail)
            x.lineNum = float(sol)
            orderList.append(x)


    print("finishing orderlist")
    return orderList


# remove duplicates
# search through the data and update to include newest data

def write(list):
    temp = filesavebox() + '.csv'
    outFile = open(temp, 'w')
    outFile.write(
        "Customer Code, SKU, Costsheet, PO, Shipdate, Price, Qty, Stocktype, SalesOrder, SOL")
    outFile.write("\n")
    for i in range(len(list)):
        line = (str(list[i].custCode)+ "," +str(list[i].styleNum)+ "," +str(list[i].costSheet)+ "," +str(list[i].poNum)+ "," +str(list[i].shipDate)+ "," +str(list[i].price)+ "," +str(list[i].qty)+ "," + str(list[i].stock) + "," +str(list[i].salesorder)+ "," +str(list[i].lineNum))
        outFile.write(line)
        outFile.write("\n")
    outFile.close()

    return 0


def testprinter(list):
    print(len(list))
    for i in range(len(list)):
        print("Customer Code", list[i].custCode)
        print("Item", list[i].style)
        print("Costsheet", list[i].costSheet)


def costsheetingest(filename):
    # open file
    try:
        xl_workbook = xlrd.open_workbook(filename)
        # Grab the sheet names
        sheet_names = xl_workbook.sheet_names()
    except:
        print("Cost sheet file unable to open: " + filename)
        logrow("Cost sheet file unable to open: " + filename)
        return None
    # print('Sheet Names', sheet_names)

    # Get the cost sheet in view
    try:
        sheet = xl_workbook.sheet_by_name("Cost Sheet")
    except:
        print("error unable to find cost sheet tab")
        logrow(filename)
        return None

    # first get the cost sheet number
    costsheetnumber = float(sheet.cell_value(2, 0))
    orderList = []
    # Print all values, iterating through rows and columns
    try:
        for row_idx in range(12, 120):  # Iterate through rows
            x = openorder()
            for col_idx in [0, 1, 2, 5, 6, 26, 30]:  # Iterate through relevant columns
                if col_idx == 0:
                    x.custCode = str(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 1:
                    x.stock = str(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 2:
                    x.styleNum = str(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 5:
                    x.description1 = str(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 6:
                    x.description2 = str(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 26:
                    print(sheet.cell_value(row_idx, col_idx))
                    x.price = float(sheet.cell_value(row_idx, col_idx))
                elif col_idx == 30:
                    x.qty = int(sheet.cell_value(row_idx, col_idx))
            if x.custCode is not '':
                x.costSheet = float(costsheetnumber)
                orderList.append(x)
            else:
                break
    except:
        print("error")
        logrow(str(filename))

    return orderList


def matchobjects(csList, ooList):
    #for each CS record, if CS # in set, find where and extract shipdate, canceldate, PO, Line

    #create index for ooList
    indexofoolist = set()
    for z in range(len(ooList)):
        indexofoolist.add(ooList[z].costSheet)


    for i in range(len(csList)):
        #improvement by checking IF the thing exists before we search through the whole thing. This way in best case, it becomes 1, and worst case goes from n to n-1
        if csList[i].costSheet in indexofoolist:
            for j in range(len(ooList)):
                if csList[i].costSheet == ooList[j].costSheet:
                        if csList[i].custCode == ooList[j].custCode:
                            if csList[i].styleNum == ooList[j].styleNum:
                                csList[i].shipDate = ooList[j].shipDate
                                csList[i].salesorder = ooList[j].salesorder
                                csList[i].lineNum = ooList[j].lineNum
                                csList[i].poNum = ooList[j].poNum
                            else:
                                continue
                        else:
                            continue
                else:
                    continue

        print("working: ",i)



    return csList

def getfiles(filename):
    try:
        # Try to open the file using the name given
        olympicsFile = open(filename, 'r')
        # If the name is valid, set Boolean to true to exit loop
        goodFile = True
    except:
        # If the name is not valid - IOError exception is raised
        print("Invalid filename, please try again ... ")

    Lines = olympicsFile.readlines()

    count = 0
    finallines = []
    # Strips the newline character
    for line in Lines:
        finallines.append(line.strip())
        #print(line.strip())
        #print("Line{}: {}".format(count, line.strip()))
    return finallines



def main():

    #temp = fileopenbox("Choose the OpenOrder Report list.txt")
    listoforders=getfiles("C:/Users/cburt/PycharmProjects/CostSheetConsilidation/Openorderreport.txt")
    # ingest the data from openorder
    openorderrows = []
    k = 0
    for i in range(len(listoforders)):
        preopenorderrows = openorderingest(listoforders[i])
        for j in range(len(preopenorderrows)):
            print(listoforders[i])
            openorderrows.append(preopenorderrows[j])
            k+=1

    print("Entering Optomize Access")
    # put it into objects
    orderList = optomizeaccess(openorderrows)

    # ingestdata from costsheets
    print("Waiting for costsheets")
    #tem2p = fileopenbox("Choose the Cost sheet Report list.txt")
    listofcostsheets = getfiles("C:/Users/cburt/PycharmProjects/CostSheetConsilidation/Listofcostsheets.txt")
    costsheetrows = []
    k = 0
    for i in range(len(listofcostsheets)):
        preocostsheet = costsheetingest(listofcostsheets[i])
        for j in range(len(preocostsheet)):
            costsheetrows.append(preocostsheet[j])
            k += 1

    print ("Matching starting")
    finalList = matchobjects(costsheetrows, orderList)

    print("Waiting for location of writeout file")
    write(finalList)


main()
