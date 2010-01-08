Attribute VB_Name = "YahooDownloadHistory"
Option Explicit

Public Sub ShowYDHForm()
    ydhForm.show
End Sub

'
' Download price history from finance.yahoo.com and place them in target range
'
Public Sub loadDataToRange(ticker As String, startDate As Date, endDate As Date, freq As String, target As Range, _
        Optional showDate As Boolean = True, Optional showOpen As Boolean = False, _
        Optional showHigh As Boolean = False, Optional showLow As Boolean = False, _
        Optional showClose As Boolean = False, Optional showVolume As Boolean = False, _
        Optional showAdjClose As Boolean = True)
    ' load from Yahoo
    Dim result() As Variant
    result = ydh(ticker, startDate, endDate, freq)
    
    ' Place in target range
    Dim nOfRows As Integer, nOfCols As Integer, rowIx As Integer, colIx As Integer
    nOfRows = UBound(result, 1)
    nOfCols = UBound(result, 2)
    If nOfCols < 6 Then
        MsgBox "Error in load. nOfCols=" & nOfCols & ", expected 6"
        Exit Sub
    End If
    With target.Cells(1, 1)
        For rowIx = 0 To nOfRows
            colIx = 0
            If showDate = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 0)
                    .NumberFormat = "yyyy-mm-dd"
                End With
                colIx = colIx + 1
            End If
            If showOpen = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 1)
                    .NumberFormat = "0.00"
                End With
                colIx = colIx + 1
            End If
            If showHigh = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 2)
                    .NumberFormat = "0.00"
                End With
                colIx = colIx + 1
            End If
            If showLow = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 3)
                    .NumberFormat = "0.00"
                End With
                colIx = colIx + 1
            End If
            If showClose = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 4)
                    .NumberFormat = "0.00"
                End With
                colIx = colIx + 1
            End If
            If showVolume = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 5)
                    .NumberFormat = "#,##0"
                End With
                colIx = colIx + 1
            End If
            If showAdjClose = True Then
                With .Offset(rowIx, colIx)
                    .Value = result(rowIx, 6)
                    .NumberFormat = "0.00"
                End With
                colIx = colIx + 1
            End If
        Next rowIx
        .Offset(0, 0).Value = ticker  ' write security ticker in the upper left corner
    End With ' target
End Sub

'
' Download historical prices from finance.yahoo.com
'
Public Function ydh(ticker As String, startDate As Date, endDate As Date, freq As String)
Attribute ydh.VB_Description = "Yahoo Download History array function. Loads historical data for given ticker from Yahoo server. Enter security ticker, start of the period, end of the period and data frequency. Use CTRL-SHIFT-Enter!"
Attribute ydh.VB_ProcData.VB_Invoke_Func = " \n14"
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
    ydh = result
End Function

