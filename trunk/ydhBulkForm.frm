VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ydhBulkForm 
   Caption         =   "Yahoo Download History - Bulk"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   OleObjectBlob   =   "ydhBulkForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ydhBulkForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttonDownload_Click()
    Dim rTics As Range
    Dim startDate As Date
    Dim endDate As Date
    Dim freq As String
    Dim target As Range
    
    labelErr.Caption = ""
    On Error GoTo ErrHandler

    Set rTics = Range(refTickers.Value)
    startDate = CDate(boxStartDate.Value)
    endDate = CDate(boxEndDate.Value)
    freq = cboFrequency.Value
    Set target = Range(refOutputRange.Value)
    
    YahooDownloadHistory.loadBulkDataToRange rTics, startDate, endDate, freq, target, _
            showDate:=cbDate, showOpen:=cbOpen, showHigh:=cbHigh, showLow:=cbLow, _
            showClose:=cbClose, showVolume:=cbVolume, showAdjClose:=cbAdjClose

    Unload Me ' close the form after loading data
ErrHandler:
    labelErr.Caption = "Error!!!"
End Sub

Private Sub UserForm_Initialize()
    refTickers.Value = ""
    boxStartDate.Value = "2009-01-01"
    boxEndDate.Value = "2010-01-01"
    With cboFrequency
        .AddItem "d"
        .AddItem "m"
        .AddItem "y"
        .Value = "m"
    End With ' cboFrequency
    refOutputRange.Value = "A1"
    cbDate = True
    cbOpen = False
    cbHigh = False
    cbLow = False
    cbClose = False
    cbVolume = False
    cbAdjClose = True
    labelErr.Caption = ""
End Sub
