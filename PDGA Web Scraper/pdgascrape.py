from bs4 import BeautifulSoup
import requests
import time

SLEEP_CONST = 0.5

link = 'https://www.pdga.com'

playerAddon = '/player/215058'

html_data = requests.get(link + playerAddon)
time.sleep(SLEEP_CONST)

soup = BeautifulSoup(html_data.text, 'html.parser')

for i in soup.find_all('td',{"class": "tournament"}):
    for j in i.find_all('a',href = True):
        eventAddon = j.get('href')
        print(link + eventAddon)
        tourny_data = requests.get(link + eventAddon)
        time.sleep(SLEEP_CONST)

        soup = BeautifulSoup(tourny_data.text, 'html.parser')

        for k in soup.find_all('tr',{"class": ["even","odd"]}):
            num = 0
            # Change find_all here to just find 
            for l in k.find_all('td',{"class": "pdga-number"}):
                if(len(l.contents) > 0):
                    num = int(l.contents[0])
                    print(num)
            if num == 215058:
                for m in k.find_all('td',{"class": "round"}):
                    print(m)



        break
    break


