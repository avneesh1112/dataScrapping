from bs4 import BeautifulSoup
import requests, openpyxl
excel= openpyxl.Workbook()
sheet=excel.active
sheet.title='2015_boxoffice_collection'
sheet.append(['Movie_name','Release_Date','Opening_Day(in₹ crores)','Opening_Weekend(in₹ crores)','End_of_week_1(in₹ crores)','lifetime(in₹ crores)'])

url = "https://www.bollywoodhungama.com/box-office-collections/filterbycountry/IND/2015/"
page = requests.get(url)
soup = BeautifulSoup(page.content,'html.parser')
lists = soup.find_all('tbody')
newlist=lists[0].find_all('tr')

for list in newlist:
     data = list.find_all('td')
     Movie_name=data[0].text
     Release_date = data[1].text.split(' ')[2]
     o = data[2].text.split(' ')
     if len(o)==3:
          Opening_Day=data[2].text.split(' ')[2]
     elif len(o)==2:
              Opening_Day = data[2].text.split(' ')[1]
     else:
          Opening_Day = data[2].text.split(' ')[0]

     op = data[3].text.split(' ')
     if len(op)==3:
          Opening_Weekend=data[3].text.split(' ')[2]
     elif len(op)==2:
              Opening_Weekend = data[3].text.split(' ')[1]
     else:
          Opening_Weekend = data[2].text.split(' ')[0]


     w = data[4].text.split(' ')
     if len(w)==3:
          End_of_week_1=data[4].text.split(' ')[2]
     elif len(w)==2:
              End_of_week_1 = data[4].text.split(' ')[1]
     else:
          End_of_week_1 = data[2].text.split(' ')[0]

     lifetime = data[5].text
     print(Movie_name,Release_date,Opening_Day,Opening_Weekend,End_of_week_1,lifetime)

     sheet.append([Movie_name,Release_date,Opening_Day,Opening_Weekend,End_of_week_1,lifetime])
     excel.save('2015_boxoffice_collection.xlsx')


