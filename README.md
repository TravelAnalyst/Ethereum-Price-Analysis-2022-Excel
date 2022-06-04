# Ethereum-Price-Analysis-2022-Excel

Using Visual Basic Editor a few lines of code are implemented on the webpage CoingGecko to extract historical price data for months of April and March of 2020. 
This quick Ethereum price prediction for the month of June is performed using analytics tools such as Regression Analysis, Forcast Price Prediction as well as Descriptive Analysis. 
Detailed analysis sheet includes formated data, price returns, moving average and several more insights as well as data vizualization. 

VBA CODE

Public Sub UseQueryTable()


Dim url As String
url = "https://www.coingecko.com/en/coins/ethereum/historical_data?end_date=2022-06-01&start_date=2021-06-01#panel"
Dim table As QueryTable
Set table = Sheet1.QueryTables.Add("URL;" & url, Sheet1.Range("A1"))
With table
.WebSelectionType = xlAllTables
.WebTables = "1"
.Refresh
End With


End Sub
