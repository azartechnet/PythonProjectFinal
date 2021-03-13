from pyexcel_xls import read_data
from bs4 import BeautifulSoup
import urllib.request as url
import xlsxwriter
import sqlite3
import re


# TO CREATE IGNORE TEXT FILE
file=open("E:\\file1.txt","r")
ignoreword=file.read().split()
ignoreset=set(ignoreword)
file.close()
print("\nSET TYPE FUNCTION:\n",ignoreset)# to avoid repeated words

heading=['TOP LIST','FREQUENCY','DENSITY']

    #TO CREATE DATA BASE
db=sqlite3.connect('data.db')
print("Data Base Created Successfully")
c=db.cursor()
#c.execute('''CREATE TABLE List(KEYWORD TEXT NOT NULL,FREQUENCY INT NOT NULL)''')
#c.execute('''CREATE TABLE Content(D1 INT NOT NULL,D2 INT NOT NULL,D3 INT NOT NULL,D4 INT NOT NULL,D5 INT NOT NULL)''')
print("Table Created Successfully") 


#COLLECT THE CONTENT FROM WEB PAGE & CALCULATE TOP 5 WORDS AND FREQUENCY  
req=url.Request("https://www.javatpoint.com/java-tutorial",data=None,headers={'user_Agent':'Mozilla 5.0(Macintosh;Indel MAc Os X 10-9-3)applewebkit/537.36(KHTML,like Gecko) chrome/35.0.1916.47 safari/537.36'})

f=url.urlopen(req)
#print(f.read())#.decode('utf-8'))

soup=BeautifulSoup(f,'html.parser')
head=[soup.title.string]

for script in soup(['script','style','[document]','head','title']):
    script.extract()
    text=soup.get_text().lower()
    #print("***",text)
    fill=filter(None,re.split('\W|\d',text))
    #print("@@@",fill)
    d={}
    #print("dict",d)
    #print("Text",text)
    wordcount=len(text)
    #print("WordCount",wordcount)
for word in fill:
     word=word.lower()
     if word not in ignoreset:
         if word not in d:d[word]=1 
         else:d[word]+=1 

a=sorted(d.items(),key=lambda x:x[1],reverse=True,)[:5]
print("%%%%",a)
density=[]
for ke,va in a:
    key=len(ke)
    print("Key::",key)
    print("Value::",va)
    dens=((key/wordcount*100))
    density.append(dens)

print("Here Calculations")

    #TO INSERT THE DATABASE
steps=[(k,v)for k,v in a]
#c.executemany("INSERT INTO List(KEYWORD,FREQUENCY) VALUES(?,?)",(steps))
#c.execute("INSERT INTO Content(D1,D2,D3,D4,D5) VALUES(?,?,?,?,?)",(density))
db.commit()
c.execute("SELECT * FROM List")

word_collection=[]
for mat in c.fetchall():
    word_collection.append(mat)
    
top_word=[]
top_frequency=[]
for k1,v1 in word_collection:
    top_word.append(k1)
    top_frequency.append(v1)
    
c.execute("SELECT * FROM Content")
value=[]
for res in c.fetchall():
    for val in res:
        value.append(val)

    
#TO APPLY THE RESULT TO XLSX WRITER
workbook=xlsxwriter.Workbook("E:\\report.xlsx")
worksheet=workbook.add_worksheet()
bold = workbook.add_format({'bold': True})

data=[top_word,top_frequency,value]

worksheet.set_column('A:A',20)
worksheet.set_column('B:B',15)
worksheet.set_column('C:C',20)

worksheet.write_row('A1',heading,bold)
worksheet.write_column('A2',data[0])
worksheet.write_column('B2',data[1])
worksheet.write_column('C2',data[2])
chart=workbook.add_chart({'type':'pie'})

chart.add_series({
        'name':       '=Sheet1!$A$2:$A$6',
        'categories': '=Sheet1!$B$2:$B$6',
        'values':     '=Sheet1!$C$2:$C$6',
      })

worksheet.write('A8',"END",bold)

    #DRAW A CHART
chart.set_title ({'name': 'MY ANALYSIS'})
worksheet.insert_chart('E2', chart)
workbook.close()

print("Find The Result in Excel")
























    


















    
