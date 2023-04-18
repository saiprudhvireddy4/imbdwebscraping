import openpyxl
import requests
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = "Top Rated Movies"
print(excel.sheetnames)
sheet.append(["RANK", "NAME", "YEAR", "RATING"])

try:
    req = requests.get("https://www.imdb.com/chart/top/")
    req.raise_for_status()
    soup = BeautifulSoup(req.content, "html.parser")

    movies = soup.find("tbody", class_="lister-list").find_all("tr")

    for movie in movies:
        name = movie.find("td", class_="titleColumn").a.text
        rank = movie.find("td", class_="titleColumn").get_text(strip=True).split(".")[0]
        year = movie.find("td", class_="titleColumn").span.text.strip("()")
        rating = movie.find("td", class_="ratingColumn imdbRating").strong.text
        print(rank, name, year, rating)
        sheet.append([rank, name, year, rating])


except Exception as e:
    print(e)

excel.save("IMBD MOVIES")
