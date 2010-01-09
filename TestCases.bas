Attribute VB_Name = "TestCases"
Option Explicit

Private testerM As cTester
Private Const testSheetNameM As String = "TEST"

Public Sub runAllTests()
    Set testerM = New cTester
    testYDHFunction
    testLoadDataToRangeProc
    testLoadBulkDataToRange
End Sub

Private Sub testLoadBulkDataToRange()
    Const startDate As Date = #6/1/2008#
    Const endDate As Date = #6/15/2008#
    Const freq As String = "d"
    Dim target As Range, rTics As Range
    
    With Worksheets(testSheetNameM)
        .Cells.ClearContents
        Set target = .Range("B2")
        Set rTics = .Range("A1:A2")
        rTics.Cells(1, 1).Value = "^FTSE"
        rTics.Cells(2, 1).Value = "SEB-A.ST"
    End With ' sheet(TEST)
    
    YahooDownloadHistory.loadBulkDataToRange rTics, startDate, endDate, freq, target, _
                showDate:=True, showOpen:=False, showHigh:=True, showLow:=True, _
                showClose:=False, showVolume:=False, showAdjClose:=True
    With target
        Call testerM.assertS("loadBulkDataToRange headers 1", .Offset(0, 0), "^FTSE")
        Call testerM.assertS("loadBulkDataToRange headers 2", .Offset(0, 5), "SEB-A.ST")
        Call testerM.assertS("loadBulkDataToRange headers 3", .Offset(0, 6), "High")
        Call testerM.assertS("loadBulkDataToRange headers 4", .Offset(0, 8), "Adj Close")
        Call testerM.assertD("loadBulkDataToRange data 1", .Offset(1, 0), #6/2/2008#) ' Date
        Call testerM.assertD("loadBulkDataToRange data 2", .Offset(2, 1), 6059)  ' High
        Call testerM.assertD("loadBulkDataToRange data 3", .Offset(3, 2), 5933.3)   ' Low
        Call testerM.assertD("loadBulkDataToRange data 4", .Offset(4, 3), 5995.3)   ' Close
        Call testerM.assertD("loadBulkDataToRange data 2.1", .Offset(5, 5), #6/10/2008#) ' Date
        Call testerM.assertD("loadBulkDataToRange data 2.2", .Offset(6, 6), 130)  ' High
        Call testerM.assertD("loadBulkDataToRange data 2.3", .Offset(7, 7), 123.25)   ' Low
        Call testerM.assertD("loadBulkDataToRange data 2.4", .Offset(8, 8), 128.75)   ' Close
        Call testerM.assertS("loadBulkDataToRange data 2.5", .Offset(9, 7), "")   ' No Data
    End With ' target
End Sub

Private Sub testLoadDataToRangeProc()
    Const ticker As String = "ERIC-B.ST"
    Const startDate As Date = #6/1/2008#
    Const endDate As Date = #7/15/2008#
    Const freq As String = "d"
    Dim target As Range
    
    With Worksheets(testSheetNameM)
        .Cells.ClearContents
        Set target = .Range("A1")
    End With ' sheet(TEST)
    
    YahooDownloadHistory.loadDataToRange ticker, startDate, endDate, freq, target, _
                showDate:=True, showOpen:=True, showHigh:=True, showLow:=True, _
                showClose:=True, showVolume:=True, showAdjClose:=True
    With target
        Call testerM.assertS("loadDataToRange headers 1", .Offset(0, 0), ticker)
        Call testerM.assertS("loadDataToRange headers 2", .Offset(0, 1), "Open")
        Call testerM.assertS("loadDataToRange headers 3", .Offset(0, 2), "High")
        Call testerM.assertS("loadDataToRange headers 4", .Offset(0, 3), "Low")
        Call testerM.assertS("loadDataToRange headers 5", .Offset(0, 4), "Close")
        Call testerM.assertS("loadDataToRange headers 6", .Offset(0, 5), "Volume")
        Call testerM.assertS("loadDataToRange headers 7", .Offset(0, 6), "Adj Close")
        Call testerM.assertD("loadDataToRange data 1", .Offset(1, 0), #6/2/2008#) ' Date
        Call testerM.assertD("loadDataToRange data 2", .Offset(2, 1), 80.7)  ' Open
        Call testerM.assertD("loadDataToRange data 3", .Offset(3, 2), 82.4)   ' High
        Call testerM.assertD("loadDataToRange data 4", .Offset(4, 3), 77.2)   ' Low
        Call testerM.assertD("loadDataToRange data 5", .Offset(5, 4), 73#)     ' Close
        Call testerM.assertD("loadDataToRange data 6", .Offset(6, 5), 18380800)   ' Volume
        Call testerM.assertD("loadDataToRange data 7", .Offset(7, 6), 70.29)   ' Adj Close
    End With ' target
    YahooDownloadHistory.loadDataToRange ticker, startDate, endDate, freq, target, _
                showDate:=True, showOpen:=False, showHigh:=False, showLow:=False, _
                showClose:=False, showVolume:=True, showAdjClose:=True
    With target
        Call testerM.assertS("loadDataToRange headers 2.1", .Offset(0, 0), ticker)
        Call testerM.assertS("loadDataToRange headers 2.2", .Offset(0, 1), "Volume")
        Call testerM.assertS("loadDataToRange headers 2.3", .Offset(0, 2), "Adj Close")
        Call testerM.assertD("loadDataToRange data 2.1", .Offset(4, 2), 75.85)   ' Adj Close
    End With ' target
End Sub

Private Sub testYDHFunction()
    Const ticker As String = "^OMX"
    Const startDate As Date = #1/1/2009#
    Const endDate As Date = #6/1/2009#
    Const freq As String = "m"
    Dim result() As Variant
    result = YahooDownloadHistory.ydh(ticker, startDate, endDate, freq, showDate:=True, showOpen:=True, showHigh:=True, showLow:=True, _
                showClose:=True, showVolume:=True, showAdjClose:=True)
    Call testerM.assertS("ydh function headers 1", CStr(result(0, 0)), "Date")
    Call testerM.assertS("ydh function headers 2", CStr(result(0, 5)), "Volume")
    Call testerM.assertS("ydh function headers 3", CStr(result(0, 6)), "Adj Close")
    Call testerM.assertD("ydh function data 1", CDate(result(1, 0)), #1/2/2009#) 'Date
    Call testerM.assertD("ydh function data 2", CDbl(result(2, 1)), 617.38) ' Open
    Call testerM.assertD("ydh function data 3", CDbl(result(3, 2)), 689.96) ' High
    Call testerM.assertD("ydh function data 4", CDbl(result(4, 3)), 641.02) ' Low
    Call testerM.assertD("ydh function data 4", CDbl(result(5, 4)), 776.5)  ' Close
    Call testerM.assertD("ydh function data 4", CDbl(result(5, 5)), 0)     ' Volume
    Call testerM.assertD("ydh function data 4", CDbl(result(6, 6)), 801.57)  ' Adj Close
End Sub
