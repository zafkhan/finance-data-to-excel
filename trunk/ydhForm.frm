VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ydhForm 
   Caption         =   "Yahoo Download History"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4620
   OleObjectBlob   =   "ydhForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ydhForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttonDownload_Click()
    Dim ticker As String
    Dim startDate As Date
    Dim endDate As Date
    Dim freq As String
    Dim target As Range
    
    labelErr.Caption = ""
    On Error GoTo ErrHandler

    ticker = Split(cboTicker.Value, " ")(0)
    startDate = CDate(boxStartDate.Value)
    endDate = CDate(boxEndDate.Value)
    freq = cboFrequency.Value
    Set target = Range(refOutputRange.Value)
    
    YahooDownloadHistory.loadDataToRange ticker, startDate, endDate, freq, target, _
            showDate:=cbDate, showOpen:=cbOpen, showHigh:=cbHigh, showLow:=cbLow, _
            showClose:=cbClose, showVolume:=cbVolume, showAdjClose:=cbAdjClose

    Unload Me ' close the form after loading data
ErrHandler:
    labelErr.Caption = "Error!!!"
End Sub


Private Sub UserForm_Initialize()
    Const tickers As String = "^DJI   Dow Jones Industrial Average|^FTSE    FTSE 100|^HSI   HANG SENG INDEX|^IXIC   NASDAQ Composite|^N225  NIKKEI 225|^OMX  OMXS30|OMXS60.ST   OMXS60|^STOXX50E    DJ EURO STOXX 50|" & _
        "ABB.ST    ABB|ALFA.ST  ALFA LAVAL|ASSA-B.ST    ASSA ABLOY-B-|ATCO-A.ST     ATLAS COPCO -A-|AZN.ST  ASTRAZENECA|" & _
        "BOL.ST  BOLIDEN|ELUX-B.ST   ELECTROLUX -B-|ERIC-B.ST   ERICSSON -B-|GETI-B.ST   GETINGE -B-|HM-B.ST HENNES & MAURITZ-B-|INVE-B.ST   INVESTOR -B-|LUPE.ST LUNDIN PETROL|MTG-B.ST    MODERN TIMES GR -B-|NDA-SEK.ST  NORDEA BANK|" & _
        "NOKI-SEK.ST    Nokia|SAND.ST   SANDVIK|SCA-B.ST    SVENSKA CELLULO -B-|SCV-B.ST    SCANIA -B-|SEB-A.ST    S-E-BANKEN -A-|SECU-B.ST   SECURITAS -B-|SHB-A.ST    SV HANDBK -A-|SKA-B.ST    SKANSKA -B-|SKF-B.ST    SKF -B-|SSAB-A.ST   SSAB -A-|SWED-A.ST   SWEDBANK -A-|SWMA.ST SWEDISH MATCH|" & _
        "TEL2-B.ST   TELE2 -B-|TLSN.ST TELIASONERA|VOLV-B.ST   VOLVO -B-"
    Dim tics() As String, i As Integer
    tics = Split(tickers, "|")
    With cboTicker
        For i = 0 To UBound(tics)
            .AddItem tics(i)
        Next i
        .Value = "^OMX"
    End With ' cboTicker
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
