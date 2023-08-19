from bs4 import BeautifulSoup
import requests

import openpyxl

excel = openpyxl.Workbook()

Sheet = excel.active
Sheet.title="Top Rated Movies"
print(excel.sheetnames)

Sheet.append(['Movie Rank','Movie Name','Year of Release','Rating'])

try:
  source = requests.get('https://www.imdb.com/search/title/?groups=top_250&sort=user_rating')
  source.raise_for_status()

  soup = BeautifulSoup(source.text,'html.parser')
  #print(soup)

  movies = soup.find('div',class_="lister-list").find_all('div',class_="lister-item mode-advanced")
  #print(len(movies))
  for movie in movies:
    name = movie.find('h3',class_="lister-item-header").a.text
    rank = movie.find('h3',class_="lister-item-header").get_text(strip=True).split('.')[0]
    year = movie.find('h3',class_="lister-item-header").get_text(strip=False).split('\n')[3].strip('()')
    rating = movie.find('div',class_="ratings-bar").strong.text
    Sheet.append([rank,name,year,rating])

except Exception as e:
  print(e)

excel.save('Top Movies.xlsx')

