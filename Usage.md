Typical usage of the package

# Introduction #

there are two basic ways how to use the package to download historical data form finance.yahoo.com. Either to use menu or spreadsheet function ydh()


# Using Menu #

Click on YDownload main menu entry. There are two possible choices either:
"Get History - Single Ticket", which is used for single ticket downloads or "Get History - Bulk" which is used for downloads for list of tickets at the same time.

In the dialog window fill in Yahoo ticker (or column range with tickets for bulk downloads), start and end dates and data frequency. The output range specifies the target place for the data. You could choose what data are you interested in (Date, Open, High, Low, Close, Volume or Adjusted Close). Click "Download!" button and the data will be placed in the "Output Range".

# Using ydh() Function #

Select region with 7 columns and enough rows to place all the historical data. Type
=ydh(ticker, startDate, endDate, frequency)
and press CTRL-SHIFT-Enter (Array formula). The historical data will be downloaded to the selected range.