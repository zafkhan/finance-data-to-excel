Attribute VB_Name = "YahooDownloadHistory"
Option Explicit

Public Sub ShowYDHForm()
    ydhForm.Show
End Sub

Public Sub ShowYDHBulkForm()
    ydhBulkForm.Show
End Sub

'
' Download history for several securites (bulk download)
' The list of tickers is in range rTics
'
Public Sub loadBulkDataToRange(rTics As Range, startDate As Date, endDate As Date, freq As String, target As Range, _
        Optional showDate As Boolean, Optional showOpen As Boolean, _
        Optional showHigh As Boolean, Optional showLow As Boolean, _
        Optional showClose As Boolean, Optional showVolume As Boolean, _
        Optional showAdjClose As Boolean)
    Dim rTarget As Range
    Dim nOfCols As Integer, nOfTics As Integer, i As Integer
    Dim ticker As String
    nOfCols = calculateNOfColumns(showDate, showOpen, showHigh, showLow, showClose, showVolume, showAdjClose)
    nOfTics = rTics.Rows.Count
    For i = 0 To nOfTics - 1
        ticker = rTics.Cells(1, 1).Offset(i, 0)
        Set rTarget = target.Cells(1, 1).Offset(0, i * (nOfCols + 2))
        loadDataToRange ticker, startDate, endDate, freq, rTarget, showDate, showOpen, showHigh, showLow, showClose, showVolume, showAdjClose
    Next i
End Sub
'
' Download single security history from finance.yahoo.com and place them in target range
'
Public Sub loadDataToRange(ticker As String, startDate As Date, endDate As Date, freq As String, target As Range, _
        Optional showDate As Boolean, Optional showOpen As Boolean, _
        Optional showHigh As Boolean, Optional showLow As Boolean, _
        Optional showClose As Boolean, Optional showVolume As Boolean, _
        Optional showAdjClose As Boolean)
    ' load from Yahoo
    Dim result() As Variant
    result = ydh(ticker, startDate, endDate, freq, _
                showDate, showOpen, showHigh, showLow, showClose, showVolume, showAdjClose)
    
    ' Place in target range
    Dim nOfRows As Integer, nOfCols As Integer, rowIx As Integer, colIx As Integer
    nOfRows = UBound(result, 1)
    nOfCols = UBound(result, 2)
    If nOfCols < 1 Or nOfCols > 6 Then
        MsgBox "Error in load. nOfCols=" & nOfCols & ", expected 6"
        Exit Sub
    End If
    With Range(target.Cells(1, 1), target.Cells(1, 1).Offset(nOfRows, nOfCols))
        .Value = result
        .Cells(1, 1) = ticker
        .Font.Bold = False
        .Rows(1).Font.Bold = True
        .NumberFormat = "#,##0.00"
        If showDate = True Then .Columns(1).NumberFormat = "YYYY-mm-dd"
    End With ' region
End Sub

Public Function ydh(ticker As String, startDate As Date, endDate As Date, freq As String, _
        Optional showDate As Boolean = True, Optional showOpen As Boolean = False, _
        Optional showHigh As Boolean = False, Optional showLow As Boolean = False, _
        Optional showClose As Boolean = False, Optional showVolume As Boolean = False, _
        Optional showAdjClose As Boolean = True)
Attribute ydh.VB_Description = "Array formula downloading historical quotes from Yahoo Finance. "
Attribute ydh.VB_ProcData.VB_Invoke_Func = " \n14"
    ' load from Yahoo
    Dim data() As Variant
    data = ydhLoadData(ticker, startDate, endDate, freq)
    ' Check sanity
    If UBound(data, 1) < 6 Then
        MsgBox "Error in load. nOfCols=" & UBound(data, 1) & ", expected 6"
        Exit Function
    End If
    
    ' filter only requested columns
    Dim nOfRows As Integer, nOfCols As Integer, rowIx As Integer, colIx As Integer
    nOfRows = UBound(data, 1)
    nOfCols = calculateNOfColumns(showDate, showOpen, showHigh, showLow, showClose, showVolume, showAdjClose)
    
    If nOfCols < 1 Or nOfCols > 6 Then
        ydh = "Wrong input. Select column"
        Exit Function
    End If
    ReDim result(nOfRows, nOfCols)
    For rowIx = 0 To nOfRows
        colIx = 0
        If showDate = True Then
            result(rowIx, colIx) = data(rowIx, 0)
            colIx = colIx + 1
        End If
        If showOpen = True Then
            result(rowIx, colIx) = data(rowIx, 1)
            colIx = colIx + 1
        End If
        If showHigh = True Then
            result(rowIx, colIx) = data(rowIx, 2)
            colIx = colIx + 1
        End If
        If showLow = True Then
            result(rowIx, colIx) = data(rowIx, 3)
            colIx = colIx + 1
        End If
        If showClose = True Then
            result(rowIx, colIx) = data(rowIx, 4)
            colIx = colIx + 1
        End If
        If showVolume = True Then
            result(rowIx, colIx) = data(rowIx, 5)
            colIx = colIx + 1
        End If
        If showAdjClose = True Then
            result(rowIx, colIx) = data(rowIx, 6)
            colIx = colIx + 1
        End If
    Next rowIx
    ydh = result
End Function

'
' Download raw historical data from finance.yahoo.com
'
Private Function ydhLoadData(ticker As String, startDate As Date, endDate As Date, freq As String)
Attribute ydhLoadData.VB_Description = "Yahoo Download History array function. Loads historical data for given ticker from Yahoo server. Enter security ticker, start of the period, end of the period and data frequency. Use CTRL-SHIFT-Enter!"
Attribute ydhLoadData.VB_ProcData.VB_Invoke_Func = " \n14"
    ' create URL
    Dim occUrl As String
    occUrl = "http://ichart.finance.yahoo.com/table.csv?s=" & ticker & _
        "&a=" & (Month(startDate) - 1) & "&b=" & Day(startDate) & "&c=" & Year(startDate) & _
        "&d=" & (Month(endDate) - 1) & "&e=" & Day(endDate) & "&f=" & Year(endDate) & _
        "&g=" & freq & "&ignore=.csv"
    Debug.Print occUrl
    ' download from Yahoo
    Dim tableText As String
    Dim xml As Object
    
    Set xml = CreateObject("Microsoft.XMLHTTP")
    xml.Open "GET", occUrl, False
    xml.send
    tableText = xml.ResponseText
    Set xml = Nothing
    
    ' parse the result
    Dim lines() As String, field() As String
    Dim nOfCols As Integer, nOfRows As Integer, i As Integer, j As Integer, rowIx As Integer
    lines = Split(tableText, vbLf)  ' Asc(vbLf)=10; tried vbCrLf and vbNewLine
    nOfCols = UBound(Split(lines(0), ","))
    nOfRows = UBound(lines) - 1 ' The last element in "result" is empty string and must be ignored
    ReDim result(nOfRows, nOfCols) As Variant
    
    For i = 0 To nOfRows
        field = Split(lines(i), ",")
        For j = 0 To nOfCols
            If i = 0 Then
                rowIx = 0
            Else
                rowIx = nOfRows - i + 1 ' revert the date order to ascending
            End If
            result(rowIx, j) = field(j)
        Next j
    Next i
    ydhLoadData = result
End Function

Private Function calculateNOfColumns(Optional showDate As Boolean, Optional showOpen As Boolean, _
                                    Optional showHigh As Boolean, Optional showLow As Boolean, _
                                    Optional showClose As Boolean, Optional showVolume As Boolean, _
                                    Optional showAdjClose As Boolean)
    calculateNOfColumns = Abs(CInt(showDate) + CInt(showOpen) + CInt(showHigh) + CInt(showLow) + CInt(showClose) + CInt(showVolume) + CInt(showAdjClose)) - 1
End Function


