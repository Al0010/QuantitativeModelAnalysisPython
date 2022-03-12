# Quantitative Model Analysis With Python
Quantitative analysis model with Python. Extrapolate variance, standard deviation, correlation and beta of a list of securities in a few clicks.

This model can help quantitative analysts in optimising their market research. It simplifies complex calculations such as variance, standard deviation, correlation and beta. The model is capable of calculating data on a plurality of financial assets. 

It is used to quickly derive data for analysis. The analyst can verify the data and choose the best alternatives in a short time. 

The script produces two outputs; the first is an Excel database with the required calculation data, the second is a heatmap showing the correlation of the basket of securities. 

Through this script:

1) Download data from Yahoo Finance 
2) Create a dataframe using Pandas 
3) Extracts specific data from the dataframe
4) Calculate variance
5) Calculate standard deviation 
6) Calculate correlation coefficient
7) Calculate beta 
8) Create a database of financial data in Excel 
9) Create a coefficient correlation heatmap

In this example, it is used to search for the most volatile asset that is least correlated to the reference index. 
The basket of stocks concerns assets with a core business in the commodities sector. An oil future is used as an index.

![HeatmapTest](https://user-images.githubusercontent.com/100917872/158033758-7f8bcf6f-7352-42f8-8ede-7ff8f2b4d281.png)

Version: Python V 3.10.2

# How to run this script 
1) Import libraries 
2) Set the script 
3) Run 

You can copy and paste the code into your Virtual Studio Code to test the script with the default settings in the example.

# Install Libraries
- pip3 install pandas 
- pip3 install datareader 
- pip3 install yahoo-finance
- pip3 install math 
- pip3 install matplotlib.pyplot
- pip3 install mplfinance
- pip3 install seaborn
- pip3 install openpyxl

# Libraries to import 
- import pandas as pd
- import pandas_datareader.data as web
- import datetime as dt
- import math
- import matplotlib.pyplot as plt
- import mplfinance as mpf
- import seaborn as sns
- from openpyxl import Workbook, load_workbook
- from openpyxl.utils import get_column_letter
- from openpyxl.styles import Font
