VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' --------- Constants and Private Variables ---------
Private sheetM As Worksheet
Private rowNoM As Integer ' the last used row
Private Const sheetNameM As String = "UnitTests"

' ------------ Constructor -----------------
Private Sub Class_Initialize()
    runSetUp
End Sub

' ---------------- Public Methods ----------------
Public Sub assert(name As String, result As Boolean, Optional calculated As Variant, Optional expected As Variant)
    rowNoM = rowNoM + 1
    With sheetM
        .Range("total").Value = .Range("total").Value + 1
        .Cells(rowNoM, 1).Value = name
        With .Cells(rowNoM, 3)
            .Value = Now()
            .NumberFormat = "dd-mmm h:mm:ss"
        End With
        With .Cells(rowNoM, 2)
            If result Then
                .Value = "PASS"
                .Font.Color = RGB(0, 255, 0)
                sheetM.Range("passed").Value = sheetM.Range("passed").Value + 1
            Else
                .Value = "FAILED"
                .Font.Color = RGB(255, 0, 0)
                sheetM.Range("failed").Value = sheetM.Range("failed").Value + 1
                .Offset(0, 1).Value = "calc=" & calculated & ", exp=" & expected
            End If
        End With ' .cells(rowNoM,2)
    End With 'sheetM
End Sub

' Double assert
Public Sub assertD(name As String, calculated As Double, expected As Double)
    assert name, Abs(calculated - expected) < 0.00000001, calculated, expected
End Sub

' String assert
Public Sub assertS(name As String, calculated As String, expected As String)
    assert name, calculated = expected, calculated, expected
End Sub

' --------------- Private Methods ----------------

' Clear the test page and reset all test statistics
Private Sub runSetUp()
    ' set sheet to sheetNameM. Create if it does not exist
    On Error Resume Next
    Set sheetM = Worksheets(sheetNameM)
    On Error GoTo 0
    If sheetM Is Nothing Then
        Set sheetM = Sheets.Add
        sheetM.name = sheetNameM
    End If
    
    ' Clean up
    With sheetM
        .Tab.colorIndex = 7
        .Range("1:1000").ClearContents
        With .Rows(1)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
        End With
        .Range("A1").Value = "Test Name"
        .Range("B1").Value = "Result"
        .Range("C1").Value = "Time"
        
        With .Range("E1:F1")
            .Merge
            .Value = "Statistics"
            .HorizontalAlignment = xlCenter
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
        End With
        .Range("E2").Value = "Total"
        .Range("E3").Value = "Passed"
        .Range("E4").Value = "Failed"
        .Names.Add name:="total", RefersTo:=.Range("F2")
        .Names.Add name:="passed", RefersTo:=.Range("F3")
        .Names.Add name:="failed", RefersTo:=.Range("F4")
        .Range("total").Value = 0
        .Range("passed").Value = 0
        .Range("failed").Value = 0
    End With 'sheetM
    
    rowNoM = 1
End Sub

' -------------- Example Usage ------------
'Private testerM As cTester
'
'Public Sub runAllTests()
'    Set testerM = New cTester
'    Call test1
'    Call test2
'End Sub
'
'Private Sub test1()
'    Call testerM.assertD("Test1", 1, 2)
'End Sub
'
'Private Sub test2()
'    Call testerM.assertD("Test2", 1, 1)
'End Sub
' -------------------------------------------

