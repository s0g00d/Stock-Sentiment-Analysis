import gspread
from oauth2client.service_account import ServiceAccountCredentials
from os import path
from bs4 import BeautifulSoup as bs
import pandas as pd
from urllib.request import Request, urlopen
import nltk
nltk.download('vader_lexicon')
import PySimpleGUI as sg



DATA_DIR = 'C:/Users/xbsqu/Desktop/Python Learning/Projects/Premarket Stock Price'

#Connecting to G Sheet...

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(path.join(DATA_DIR, 'client_secret.json'), scope)
client = gspread.authorize(creds)

sheet = client.open('Stock Watcher')
worksheet = sheet.get_worksheet(0)
row_limit = 1+len(worksheet.col_values(1))
#Now connected to the G Sheet.

 


for i in range(2, row_limit):
    sg.one_line_progress_meter('Stock Headline Sentiment Analyzer', i+1, row_limit, 'key', 'Analyzing the sentiment of recent news for your tracked companies...')
    
    if worksheet.acell(f'A{i}') == "":

        pass

    else:
        stock_symbol = worksheet.acell(f'A{i}').value

#Here is where we use a web scraper to grab the premarket price information

        url_slug = 'https://finviz.com/quote.ashx?t='
        stock_url = url_slug + stock_symbol

    
#We need to send a header so as to not get 403 errors
    try:        
        req = Request(stock_url, headers={'User-Agent': 'Mozilla/5.0'})
        webpage = urlopen(req).read()

#Make the Soup call & scrape the table
        page_soup = bs(webpage,'html.parser')
        headline_table = page_soup.find('table',{'class': 'fullview-news-outer'})

#Remove the publishing blog name
        for span_tag in headline_table.findAll('span'):
            span_tag.decompose()


#Turn the Soup data into a df and add headers
        headline_table = pd.read_html(str(headline_table))[0]
        headline_table.columns = ['Date', 'Title']


#Now we need to fix the date column
#First let's just use the first 9 characters since that's the length of the date portion
        headline_table['Date'] = headline_table['Date'].str.slice(stop=9)
#Now use Foward Fill to overwrite where needed
        headline_table['Date'] = headline_table.Date.where(headline_table.Date.str.contains('-')).ffill()

#Running the sentiment analyzer with VADER

        from nltk.sentiment.vader import SentimentIntensityAnalyzer as SIA

        results = []

        for headline in headline_table['Title']:
            pol_score = SIA().polarity_scores(headline) # run analysis
            pol_score['headline'] = headline # add headlines for viewing
            results.append(pol_score)

        results

#Creating the column
        headline_table['Sentiment'] = pd.DataFrame(results)['compound']

#Removing zero values to make aggregating more impactful
        headline_table = headline_table[headline_table['Sentiment'] != 0].reset_index(drop=True)
        daily_mean = headline_table.groupby('Date')['Sentiment'].mean()      
        
        median_sent_score = daily_mean.median()
        worksheet.update_cell(i, 25, median_sent_score)
    
    except:      
        worksheet.update_cell(i, 25, "")
        
        

