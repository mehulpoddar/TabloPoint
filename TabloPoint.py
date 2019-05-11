import tkinter as tk
import xlrd
import xlwt
from xlutils.copy import copy
import re

path = './../files/'
salesFile = ''
returnFile = ''
data = {}
comp = {}
msg = None
allCat = []
quit = 0

# Creating Root Window
root = tk.Tk()
root.geometry("400x320")
root.resizable(0,0)
root.title("TabloPoint")

def getNum():
    global salesFile
    
    month = salesFile[15:]

    if month == 'april': return 0
    if month == 'may': return 1
    if month == 'june': return 2
    if month == 'july': return 3
    if month == 'august': return 4
    if month == 'september': return 5
    if month == 'october': return 6
    if month == 'november': return 7
    if month == 'december': return 8
    if month == 'january': return 9
    if month == 'february': return 10
    if month == 'march': return 11
    

def loadData():
    global data, allCat, path
    
    file = (path + "Database.xlsx")
  
    wb = xlrd.open_workbook(file)
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0) 

    for i in range(sheet.nrows):
        if sheet.cell_value(i, 0) == '':
            newCategory = True
        elif newCategory:
            newCategory = False
            category = sheet.cell_value(i, 0).strip()
            allCat.append(category)
            allCat = list(set(allCat))
            data[category] = []
        else:
            data[category].append(' ' + sheet.cell_value(i, 0).strip().upper() + ' ')

def getCategory(row,string):
    global data, msg

    for category, sheets in data.items():
        for sheet in sheets:
            if sheet in (' ' + string.upper() + ' '):
                return category
    text = 'row ' + str(row) + ': ' + string + '\n'
    msg.insert('insert',text)

def loadFiles(f):
    # f = file --> 's' = sales, 'r' = return
    global salesFile, returnFile, comp, path

    if f == 's':
    	wb = xlrd.open_workbook(path + salesFile+'.xls')
    else:
        wb = xlrd.open_workbook(path + returnFile+'.xls')
    sheet = wb.sheet_by_index(0)
    sheet.cell_value(0, 0) 

    newCompany = 'ignore'
    for i in range(sheet.nrows):
        if sheet.cell_value(i, 1) == 'Grand Total':
            break

        if sheet.cell_value(i, 2) == 'GST SALES':
            continue
            
        if f == 's':
            if sheet.cell_value(i, 2) == 'SALE - GST':
                newCompany = 'true'
        else:
            if sheet.cell_value(i, 2) == 'Credit Note':
                newCompany = 'true'
            
        if newCompany == 'true':
            newCompany = 'false'
            company = sheet.cell_value(i, 1).strip()
            if company not in comp.keys():
                comp[company] = {}
                
        elif newCompany == 'false':
            category = getCategory(i, sheet.cell_value(i, 1).strip())
            
            ########################################### Wrong Categorization #########################################
            #if company == 'Shri Sumukha Agencies' and category == '1 MM TEXTURE':
                #print(i, '--->', sheet.cell_value(i, 1).strip())
            
            if f == 's':
                comp[company][category] = comp[company].setdefault(category,0.0) + sheet.cell_value(i, 3)
            else:
                comp[company][category] = comp[company].setdefault(category,0.0) - sheet.cell_value(i, 3)

def autoAdjust(wb, readerSheet, writerSheet, extraCushion):
    global path
    for row in range(readerSheet.nrows):
            for column in range(readerSheet.ncols):
                    thisCell = readerSheet.cell(row, column)
                    neededWidth = int((1 + len(str(thisCell.value))) * 256) 
                    if writerSheet.col(column).width < neededWidth:
                            writerSheet.col(column).width = neededWidth + extraCushion
    wb.save(path + 'TabloPoint Output.xls')


def saveData():
   global data, comp, allCat, path

   compNames = list(comp.keys())
   compNames.sort()
   
   rb = xlrd.open_workbook(path + 'TabloPoint Output.xls')
   read = rb.sheet_by_index(getNum())
   
   wb = copy(rb)
   sheet = wb.get_sheet(getNum())

   zeroTot = 0
   rowCount = 0
   totRowP = []
   totalTot = 0
   removing = True
   for i,name in enumerate(compNames):
       sheet.write(i+2,2,name)

       totalP = 0
       rowCount = i
       for j in range(3,10):
           category = read.cell_value(1,j)
           if removing:
               allCat.remove(category)
               totRowP.append(0)
           val = comp[name].setdefault(category,0)
           totRowP[j-3] += val
           sheet.write(i+2,j,val,xlwt.easyxf('align: horiz center'))
           totalP += read.cell_value(0,j) * val
       totalTot += totalP

       removing = False
       for cat in allCat:
           zeroTot += comp[name].setdefault(cat,0)
       sheet.write(i+2,10,totalP,xlwt.easyxf('align: horiz center'))

   style = 'font: bold on; align: horiz center'
   for j in range(3,10):
       sheet.write(rowCount+3,j,totRowP[j-3],xlwt.easyxf(style))
       
   sheet.write(rowCount+3,10,totalTot,xlwt.easyxf(style))
   sheet.write(rowCount+5,2,'Zero Point Items',xlwt.easyxf(style))
   sheet.write(rowCount+5,3,zeroTot,xlwt.easyxf(style))

   autoAdjust(wb,read,sheet,10)

def summary():
    
    rb = xlrd.open_workbook(path + 'TabloPoint Output.xls')
    read = rb.sheet_by_index(12)
   
    wb = copy(rb)
    sheet = wb.get_sheet(12)

    companies = {}

    for i in range(12):
        readi = rb.sheet_by_index(i)

        for j in range(readi.nrows-2):
            company = readi.cell_value(j+2,2)
            if company == '':
                company = 'zzzzzzz'
            
            points = readi.cell_value(j+2,10)
            
            if type(points) != type(0.0):
                break

            if company not in companies.keys():
                companies[company] = [0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0,0.0]
                
            companies[company][i] += points


    compNames = list(companies.keys())
    compNames.sort()
    i = 0
    for compName in compNames:
        points = companies[compName]
        i += 1

        if compName == 'zzzzzzz':
            compName = 'Total'
        sheet.write(i,0,compName)
        total = 0
        for j in range(12):
            total += points[j]
            sheet.write(i,j+1,points[j])
        sheet.write(i,13,total)

    autoAdjust(wb,read,sheet,10)

def main(fileName):
    global data, comp, salesFile, returnFile, msg, quit


    if quit == 1:
        root.destroy()
    else:
        quit += 1

        salesFile = 'salesregister' + fileName
        returnFile = 'goodsreturnof' + fileName
        loadData()
        loadFiles('s')
        loadFiles('r')
        saveData()
        summary()
        msg.insert('insert','Run Complete - Quit to Run again!')

def main_box(parent):
    global msg
    
    # Create Frames
    frame = tk.Frame(parent)

    runSection = tk.Frame(frame, height = 120, width = 400, highlightthickness = 3, highlightbackground = "#FFD300", bg = "#87CEFA")
    msgSection = tk.Frame(frame, height = 200, width = 400, highlightthickness = 3, highlightbackground = "#FFD300", bg = "#00BFFF")
    
    runSection.pack(side = "top")
    msgSection.pack(side = "bottom")

    runSection.pack_propagate(0)
    msgSection.pack_propagate(0)
    
    
    fileNameLabel = tk.Label(runSection, text = "Sales File Name:", bg = "#87CEFA", pady = 5)
    fileNameLabel.config(font=("Segoe Print", 9))
    fileNameLabel.pack();

    fileName = tk.Entry(runSection, bg = "#00BFFF", width = 35)
    fileName.config(font=("Segoe Print", 9))
    fileName.pack()
    
    dummy = tk.Label(runSection, text = "", bg = "#87CEFA", pady = 0)
    dummy.config(font=("Segoe Print", 2))
    dummy.pack();

    button = tk.Button(runSection, text="Tabulate Points",bg = "#00BFFF", pady = 0, padx = 50,
                         command=lambda: main(fileName.get()))
    button.config(font=("Segoe Print", 9))
    button.pack()
    
    
    heading = tk.Label(msgSection, text = "Uncategorized", bg = "#00BFFF", pady = 5)
    heading.config(font=("Eras Bold ITC", 10))
    heading.pack();

    msg = tk.Text(msgSection, bg = "#00BFFF", pady = 5)
    msg.config(font=("Segoe Print", 8))
    msg.pack()

    return frame


def buildFrame(frame_name):
    frame_container = tk.Frame(root)
    frame_container.pack(side="top", fill="both", expand=True)

    frame_container.grid_rowconfigure(0, weight=1)
    frame_container.grid_columnconfigure(0, weight=1)

    frame = frame_name(frame_container)
    frame.grid(row=0, column=0, sticky="nsew")
    frame.tkraise()

buildFrame(main_box)
root.mainloop()
