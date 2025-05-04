from bs4 import BeautifulSoup
import requests
import xlsxwriter


workbook=xlsxwriter.Workbook("spreadsheet.xlsx")
worksheet=workbook.add_worksheet('linked2')


def rma(dl):
    for i in dl:
        if i=="":
            dl.remove(i)
    return dl

k=0
l=["0"]
row=1
re=(requests.get("https://biod.co.uk/stockists/stores").text)
req=BeautifulSoup(re).find_all("div",{"class":"map-marker d-none"})

for r in req:
    col=0
    l.append(r.text)
    n=r.find("div",{"class":"card-header font-weight-bold lh-1 px-2"}).text
    langlot=r["data-latlong"]
    details=r.find("div",{"class":"card-body rounded-0 m-0 lh-1 p-2"}).text
    link=r.find("a").text
    details1=details.split("\n")
    details1=rma(details1)
    worksheet.write(row,col,n)
    
    if len(details1)<5 and len(details1)>=3:
        worksheet.write(row,col+1,details1[0])
        worksheet.write(row,col+2,details1[1])
        worksheet.write(row,col+3,details1[2])
        try:
            worksheet.write(row,col+5,details1[3])
        except:
            worksheet.write(row,col+5,"")
        try:
            worksheet.write(row,col+6,details1[4])
        except:
            pass
    try:
        if details[3].startswith("Tel:") or details[3].startswith("Telephone"):
            worksheet.write(row,col+1,details1[0])
            worksheet.write(row,col+2,details1[1])
            worksheet.write(row,col+3,details1[2])
            try:
                worksheet.write(row,col+5,details1[3])
            except:
                worksheet.write(row,col+5,"")
            try:
                worksheet.write(row,col+6,details1[4])
            except:
                worksheet.write(row,col+6,"")
        elif details1[2].startswith("Tel:") or details1[2].startswith("Telephone"):
            worksheet.write(row,col+1,details1[0])
            worksheet.write(row,col+2,details1[1])
            try:
                worksheet.write(row,col+5,details1[2])
                worksheet.write(row,col+6,details1[3])
            except:
                worksheet.write(row,col+6,"")
            try:
                worksheet.write(row,col+7,details1[4])
            except:
                worksheet.write(row,col+7,"")
        if (len(details1)>6 or len(details1)>3) and not(details1[3].startswith("Tel:") or details1[3].startswith("Telephone")):
            details1.pop(2)
            worksheet.write(row,col+1,details1[0])
            worksheet.write(row,col+2,details1[1])
            worksheet.write(row,col+3,details1[2])
            try:
                worksheet.write(row,col+5,details1[3])
            except:
                worksheet.write(row,col+5,"")
            try:
                worksheet.write(row,col+6,details1[4])
            except:
                worksheet.write(row,col+6,"")
            try:
                worksheet.write(row,col+7,details1[5])
            except:
                worksheet.write(row,col+7,"")
            try:
                worksheet.write(row,col+8,details1[6])
            except:
                worksheet.write(row,col+8,"")
 
    except:
        pass
    else:
        try:
            worksheet.write(row,col+1,details1[0])
        except:
            worksheet.write(row,col+1,".................")
        try:
            
            worksheet.write(row,col+2,details1[1])
        except:
            worksheet.write(row,col+2,".................")
        try:
            
            worksheet.write(row,col+3,details1[2])
        except:
            worksheet.write(row,col+3,".................")
        try:
            
            worksheet.write(row,col+4,details1[3])
        except:
            worksheet.write(row,col+4,".................")
        try:
            
            worksheet.write(row,col+5,details1[4])
        except:
            worksheet.write(row,col+5,".................")
        try:
            worksheet.write(row,col+6,details1[5])
        except:
            worksheet.write(row,col+6,".................")
    
    worksheet.write(row,col+7,langlot)
    worksheet.write(row,col+8,link)
    row+=1
workbook.close()
print(len(req),k)


