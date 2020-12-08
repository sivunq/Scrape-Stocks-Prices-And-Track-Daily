import urllib,requests,openpyxl
from openpyxl.styles import colors,Color, PatternFill, Font, Border
from openpyxl.cell import Cell
from datetime import datetime
from bs4 import BeautifulSoup

wb=openpyxl.load_workbook("FinTrack.xlsx")
sheet=wb["CurrentStocks"]
stocks=[]
way=[]
link=[]
initialPrice=[]
sharesCount=[]
colIndex=2

#read all constant values
while(sheet.cell(row=1,column=colIndex).value != "Total"):
    stocks.append(sheet.cell(row=1,column=colIndex).value)
    way.append(sheet.cell(row=2,column=colIndex).value)
    link.append(sheet.cell(row=3,column=colIndex).value)
    initialPrice.append(sheet.cell(row=4,column=colIndex).value)
    sharesCount.append(sheet.cell(row=5,column=colIndex).value)
    colIndex=colIndex+1

todayRow=95 #Update
todayDate=str(datetime.today().strftime('%d-%m-%Y'))

#logic to find empty cell or update today's data
while(True):
    if(str(sheet.cell(row=todayRow,column=1).value) == todayDate):
        break
    if(str(sheet.cell(row=todayRow,column=1).value) == 'None'):
        sheet.cell(row=todayRow,column=1).value = todayDate
        break
    todayRow=todayRow+1

index=0
total=0
headers = {"user-agent" : "Mozilla/5.0 (Linux; Android 7.0; SM-G930V Build/NRD90M) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.125 Mobile Safari/537.36"}
for stock in stocks:
    #money control has two ways of showing stock prices, both handled seperately
    if(way[index]=="new"):
        resp = requests.get(link[index], headers=headers)
        soup = BeautifulSoup(resp.content, "html.parser")
        mydivs = soup.findAll("div", {"class": "nsecp"})
        price=mydivs[0]["rel"]
    elif(way[index]=="old"):
        resp = requests.get(link[index], headers=headers)
        soup = BeautifulSoup(resp.content, "html.parser")
        mydivs = soup.findAll("span", {"class": "span_price_wrap"})
        price=mydivs[1].text
    price=(float(price)-initialPrice[index])*sharesCount[index]
    price= round(price,2)
    total=total+price
    sheet.cell(row=todayRow,column=index+2).value=price
    print(index+1,".",stock,":",price)
    cellFill=sheet.cell(row=todayRow,column=index+2)

    #make cell pink color if today's price is lesser than yesterday
    if(price < (sheet.cell(row=todayRow-1,column=index+2).value)):
        myColor = openpyxl.styles.colors.Color(rgb='FCE4D6')
        myFill = openpyxl.styles.fills.PatternFill(patternType='solid', fgColor=myColor)
        cellFill.fill = myFill
            
    index=index+1
#update total and day change
total=round(total,2)
sheet.cell(row=todayRow,column=index+2).value=total
prevTotal=sheet.cell(row=todayRow-1,column=index+2).value
sheet.cell(row=todayRow,column=index+3).value=total-prevTotal
wb.save("FinTrack.xlsx")
wb.close()
print("Done")
input("Press Enter To Exit")
