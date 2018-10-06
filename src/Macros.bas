Attribute VB_Name = "Macros"
Attribute VB_Description = "Application entry points."
'@Folder("Battleship")
'@Description("Application entry points.")
Option Explicit
'@Ignore MoveFieldCloserToUsage
Private controller As GameController

Public Sub PlayWorksheetInterface()
    Dim view As WorksheetView
    Set view = New WorksheetView
    
    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer
    
    Set controller = New GameController
    controller.NewGame GridViewAdapter.Create(view), randomizer
End Sub

'@Ignore StopKeyword
Public Sub PlayOtherInterface()

    Const message As String = _
        "No, really - this UI isn't implemented." & vbNewLine & _
        "Will you implement it?"
        
    If MsgBox(message, vbInformation + vbYesNo, "Battleship") = vbYes Then
        Stop
        ' ~> Didn't mean to stop here?
        ' ~> Press F5 and close this window.
        ' ~> Nobody will know ;)
        End
    End If
    
End Sub



