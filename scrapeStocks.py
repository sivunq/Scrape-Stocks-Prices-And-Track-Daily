import urllib,requests,openpyxl
from openpyxl.styles import colors,Color, PatternFill, Font, Border,Side,Alignment
from openpyxl.cell import Cell
from openpyxl.formatting.rule import ColorScaleRule
from datetime import datetime
from bs4 import BeautifulSoup

wb=openpyxl.load_workbook("FinTrack.xlsx")
sheet=wb["CurrentStocks"]

stocksData={ 
    "stockNames":[],
    "webLink":[],
    "buyPrice":[],
    "shareCount":[]
}

colIndex=2
#read all constant values
while(sheet.cell(row=1,column=colIndex).value != "Total"):
    stocksData["stockNames"].append(sheet.cell(row=1,column=colIndex).value)
    stocksData["webLink"].append(sheet.cell(row=2,column=colIndex).value)
    stocksData["shareCount"].append(sheet.cell(row=3,column=colIndex).value)
    stocksData["buyPrice"].append(sheet.cell(row=4,column=colIndex).value)
    colIndex=colIndex+1

todayRow=12 #Update
todayDate=str(datetime.today().strftime('%d-%m-%Y'))

#logic to find empty cell or update today's data
while(True):
    if(str(sheet.cell(row=todayRow,column=1).value) == todayDate):
        break
    if(str(sheet.cell(row=todayRow,column=1).value) == 'None'):
        sheet.cell(row=todayRow,column=1).value = todayDate
        break
    todayRow=todayRow+1

headers = {"user-agent" : "Mozilla/5.0 (Linux; Android 7.0; SM-G930V Build/NRD90M) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.125 Mobile Safari/537.36"}
total=0
for index in range(1,len(stocksData["stockNames"])+4):
    cellFill=sheet.cell(row=todayRow,column=index)
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'),top=Side(style='thin'),bottom=Side(style='thin'))
    cellFill.border = thin_border
    cellFill.alignment = Alignment(horizontal='center')
    
    if(index>1 and index <= len(stocksData["stockNames"])+1):
        resp = requests.get(stocksData["webLink"][index-2], headers=headers)
        soup = BeautifulSoup(resp.content, "html.parser")
        mydivs = soup.findAll("div", {"class": "nsecp"})
        price=mydivs[0]["rel"]
        price= round(float(price),2)

        #find current price
        sheet.cell(row=6,column=index).value=price #currentPrice
        sheet.cell(row=5,column=index).value=price-stocksData["buyPrice"][index-2] #difference
        profit=(float(price)-stocksData["buyPrice"][index-2])*stocksData["shareCount"][index-2]
        profit= round(profit,2)
        total=total+profit

        #fill profit
        sheet.cell(row=todayRow,column=index).value=profit
        print(index-1,".",stocksData["stockNames"][index-2],":",profit)

        #fill heat map
        colIndex=openpyxl.utils.cell.get_column_letter(index)
        sheet.conditional_formatting.add(colIndex+str(8)+":"+colIndex+str(todayRow) ,ColorScaleRule(start_type='min', start_value=0, start_color='F5602E',
                                            mid_type='percentile', mid_value=50, mid_color='F8F80E',
                                            end_type='max', end_value=100, end_color='51C806'))

    darkGreen = openpyxl.styles.colors.Color(rgb='70AD47')	
    myFill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=darkGreen)
    cellFill.fill = myFill
            
    index=index+1

#update total and day change
total=round(total,2)
sheet.cell(row=todayRow,column=index-2).value=total

prevTotal=sheet.cell(row=todayRow-1,column=index-2).value
sheet.cell(row=todayRow,column=index-1).value=total-prevTotal

wb.save("FinTrack.xlsx")
wb.close()
print("Done")
input("Press Enter To Exit")
