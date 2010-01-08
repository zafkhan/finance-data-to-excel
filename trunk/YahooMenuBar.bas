Attribute VB_Name = "YahooMenuBar"
Option Explicit

Private Const menuNameC As String = "&YDownload"

' Tools-References: Microsoft Office Object 12.0 Library
Sub AddYahooMenu()
    Dim cMenu1 As CommandBarControl
    Dim cbMainMenuBar As CommandBar
    Dim iHelpMenu As Integer
    Dim cbcCutomMenu As CommandBarControl


    '(1)Set a CommandBar variable to Worksheet menu bar
    Set cbMainMenuBar = Application.CommandBars("Worksheet Menu Bar")
    '(2)Delete any existing one. We must use On Error Resume next in case it does not exist.
    On Error Resume Next
    cbMainMenuBar.Controls(menuNameC).Delete
    On Error GoTo 0
         
    '(3)Add a Control to the "Worksheet Menu Bar" before Help.
    iHelpMenu = cbMainMenuBar.Controls("Help").Index
    Set cbcCutomMenu = cbMainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=iHelpMenu)
                      
    '(5)Give the control a caption
    cbcCutomMenu.Caption = menuNameC
         
    '(6) Working with our new Control, add a sub control and give it a Caption and tell it which macro to run (OnAction).
     With cbcCutomMenu.Controls.Add(Type:=msoControlButton)
        .Caption = "Yahoo Download History"
        .OnAction = "YahooDownloadHistory.ShowYDHForm"
        .FaceId = 29
     End With
End Sub

Sub DeleteYahooMenu()
    On Error Resume Next
    Application.CommandBars("Worksheet Menu Bar").Controls(menuNameC).Delete
    On Error GoTo 0
End Sub


