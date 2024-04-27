from selenium import webdriver
from bs4 import BeautifulSoup
from textblob import TextBlob
import time
import re
import csv
import os
import requests
import openpyxl
from PIL import Image
from io import BytesIO
from vaderSentiment.vaderSentiment import SentimentIntensityAnalyzer

search_term = "Congress elections"
driver = webdriver.Chrome()
n = 1 # pages to see :)
count = 0 # just to keep count of how many headlines are fetched :)

images = {}
# // this dictionary will store the list of images

sentiments = []

for i in range(n):
    start = i * 10
    driver.get(f"https://news.google.com/search?q={search_term}&hl=en-IN&gl=IN&ceid=IN%3Aen")
    time.sleep(5)
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')
    article_tags = soup.find_all('article', {"class": "IFHyqb"})
    for article in article_tags:
        headline_tag = article.find('a', {"class": "JtKRv"})
        headline = str(headline_tag.get_text()).lower()
        headline = headline.replace("lok sabha elections 2024", "")\
            .replace("lok sabha election 2024", "")\
            .replace("lok sabha elections", "")\
            .replace("'", "")\
            .replace(":", "")\
            .replace("mint", "")\
            .replace("...", "").replace("|", ",")


        image_tag = article.find('img', {"class": ["msvBD", "zC7z7b"]})
        if image_tag:
            image_url = image_tag['src']
        blob = TextBlob(headline)
        sentiment = blob.sentiment
        #MIGRATE::::
        analyzer = SentimentIntensityAnalyzer()
        sentiment_scores = analyzer.polarity_scores(headline)
        sentiment = sentiment_scores['compound']
        sentiments.append(
            {
                'who': image_url, # who posted the article?
                'headline': headline,
                'sentiment': sentiment
            }
        )
        count += 1

driver.quit()
print("total number of articles examined: ", count)

wb = openpyxl.Workbook()
sheet = wb.active
sheet.title = "Data"
headers = ["who", "headline", "sentiment"]
sheet.append(headers)

for content in sentiments:
    image_url = content["who"]
    response = requests.get(image_url)
    if response.status_code==200:
        image = Image.open(BytesIO(response.content))
        png_image = BytesIO()
        image.save(png_image, format="PNG")
        png_image.seek(0)
        image = Image.open(png_image)
        img_cell = openpyxl.drawing.image.Image(image)
        sheet.add_image(img_cell, f"A{sheet.max_row+1}")
    else:
        print("Image download failed :(")
    sheet.append([
        content['who'],
        content['headline'],
        str(content['sentiment'])
    ])

wb.save("data_with_images.xlsx")