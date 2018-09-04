VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameRandomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship")
Option Explicit
Implements IRandomizer

Private Sub Class_Initialize()
    Randomize
End Sub

Public Function Between(ByVal lower As Long, ByVal upper As Long) As Long
    Between = CLng(Int((upper - lower + 1) * VBA.Math.Rnd + lower))
End Function

Public Function NextSingle() As Single
    NextSingle = VBA.Math.Rnd
End Function

Private Function IRandomizer_Between(ByVal lower As Long, ByVal upper As Long) As Long
    IRandomizer_Between = Between(lower, upper)
End Function

Private Function IRandomizer_NextSingle() As Single
    IRandomizer_NextSingle = NextSingle
End Function


