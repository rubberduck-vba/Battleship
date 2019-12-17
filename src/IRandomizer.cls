VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IRandomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Battleship.Model.Abstract")
'@Interface
Option Explicit

Public Enum Probability
    Never        '   0.0%
    AlmostNever  '   7.5%
    Rarely       '  20.0%
    Sometimes    '  50.0%
    Often        '  75.0%
    AlmostAlways '  92.5%
    Always       ' 100.0%
End Enum

Public Function Maybe(ByVal likelihood As Probability) As Boolean
End Function

Public Function NextSingle() As Single
End Function

Public Function Between(ByVal lower As Long, ByVal upper As Long) As Long
End Function
