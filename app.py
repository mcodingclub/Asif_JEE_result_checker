import requests
from bs4 import BeautifulSoup

url = 'https://ntaresults.nic.in/resultservices/JEEMAINauth23s2p1'
r = requests.get(url)

htmlcontent = r.content
soup = BeautifulSoup(htmlcontent,'html.parser')
print(soup)
