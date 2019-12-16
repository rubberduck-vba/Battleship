Attribute VB_Name = "Macros"
Attribute VB_Description = "Application entry points."
'@Folder("Battleship")
'@ModuleDescription("Application entry points.")
Option Explicit
'@Ignore MoveFieldCloserToUsage
Private controller As IGameController

Public Sub PlayWorksheetInterface()
    
    Dim adapter As GridViewAdapter
    Set adapter = GridViewAdapter.Create(New WorksheetView)
    
    Dim randomizer As IRandomizer
    Set randomizer = New GameRandomizer
    
    Dim players As IPlayerFactory
    Set players = PlayerFactory.Create(randomizer)
    
    Set controller = StandardGameController.Create(adapter, randomizer, players)
    controller.NewGame
    
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



