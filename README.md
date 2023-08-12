# Automated-Stock-Selector

This stock selection and ranking software is a python based project that leverages uses webscraping, mathematical analysis and financial information to eliminate and rank stocks within each sector. For more information about this project, please feel free to visit my website:

https://lucasluwa.com/Projects/Individual/Stonks/Stocks

Project Overview:
This project is split into three parts. All parts rely on master spreadsheets and include automatic file naming systems when generating files/folders
- Part 1: Webscraper using BeautifulSoup to pull information on 6000+ stocks on the NASDAQ and NYSE. All the raw data is split apart and added to an excel spreadsheet. A recovery system has also been added in the event that a problem occurs during the scraping process. This allows users to resume processing at a certain spot.
- Part 2: This part of the project takes the raw data from part 1 and splits the data into meaningful segments of data. 
- Part 3: This stage performs statistical analysis on the data and eliminates stocks based on a predefined criteria. Stocks that make it through this stage are added to the final excel spreadsheet. Partially implemented reference code is provided for those that are interested in developing their own strategy/algorithms.







