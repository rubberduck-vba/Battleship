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

'Private Const ThresholdNever As Single = 0
Private Const ThresholdAlmostNever As Single = 0.1
Private Const ThresholdRarely As Single = 0.25
Private Const ThresholdSometimes As Single = 0.5 ' median
Private Const ThresholdOften As Single = 1 - ThresholdRarely
Private Const ThresholdAlmostAlways As Single = 1 - ThresholdAlmostNever
'Private Const ThresholdAlways As Single = 1

Option Explicit
Implements IRandomizer

Private Sub Class_Initialize()
    Randomize
End Sub

Public Property Get Probability(ByVal value As Probability) As Single
    Select Case value
        Case AlmostNever
            Probability = ThresholdAlmostNever
        Case Rarely
            Probability = ThresholdRarely
        Case Sometimes
            Probability = ThresholdSometimes
        Case Often
            Probability = ThresholdOften
        Case AlmostAlways
            Probability = ThresholdAlmostAlways
        Case Else
            'tried invoking with Probability.Always or Probability.Never? why use randomizer then?
            Debug.Assert False
            Probability = 1
    End Select
End Property

Public Function Between(ByVal lower As Long, ByVal upper As Long) As Long
    Between = CLng(Int((upper - lower + 1) * VBA.Math.Rnd + lower))
End Function

Public Function NextSingle() As Single
    NextSingle = VBA.Math.Rnd
End Function

Private Function IRandomizer_Between(ByVal lower As Long, ByVal upper As Long) As Long
    IRandomizer_Between = Between(lower, upper)
End Function

Private Function IRandomizer_Maybe(ByVal likelihood As Probability) As Boolean
    IRandomizer_Maybe = Me.NextSingle < GameRandomizer.Probability(likelihood)
End Function

Private Function IRandomizer_NextSingle() As Single
    IRandomizer_NextSingle = NextSingle
End Function

