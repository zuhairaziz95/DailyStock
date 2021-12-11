# DailyStock
DailyStockInformation

The worksheet consist of multiple stock information with their name and corrseponds shortform in the Yahoo Finance.
The price of the stock can be obtain directly from Yahoo Finance from historical data and at closing price column.

![image](https://user-images.githubusercontent.com/65912073/145676224-64cecf67-f108-45d9-9a48-c9f0b98ebf02.png)


![image](https://user-images.githubusercontent.com/65912073/145676150-344e6300-38e9-4e04-8808-df6e26a58b82.png)

The automation was made using Excel VBA Macro which query data from Yahoo Finance and update it automatically
with the stock name. The stock ticker column is case sensitive , the stock name must be the same from the Yahoo Finance
website. Otherwise, it will prompt wiith no table query.

For example Vanguard Berkshire Hathaway Inc. the shortform is (BRK-B) there should be dash in between. The code is flexible and it
will look at the stock ticker column and read the data and update it accordingly.More stock name can be added underneath the table. 

There is button added inside the worksheet so user can
run it daily. 
