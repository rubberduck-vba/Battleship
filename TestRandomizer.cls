VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TestRandomizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A configurable, deterministic, mock implementation of a randomizer."
'@Folder("Tests.Stubs")
'@Description("A configurable, deterministic, mock implementation of a randomizer.")
Option Explicit
Implements IRandomizer

Private Type TState
    currentIndex As Long
    sequence As Variant
End Type

Private this As TState

Public Sub Setup(ParamArray sequence() As Variant)
    ReDim this.sequence(LBound(sequence) To UBound(sequence))
    Dim i As Long
    For i = LBound(sequence) To UBound(sequence)
        this.sequence = sequence(i)
    Next
    this.currentIndex = 0
End Sub

Public Property Get NextValue() As Single
    If this.currentIndex > UBound(this.sequence) Then Reset
    
    NextValue = this.sequence(this.currentIndex)
    this.currentIndex = this.currentIndex + 1
End Property

Private Sub Reset()
    this.currentIndex = 0
End Sub

Private Function IRandomizer_Between(ByVal lower As Long, ByVal upper As Long) As Long
    IRandomizer_Between = NextValue
End Function

Private Function IRandomizer_NextSingle() As Single
    IRandomizer_NextSingle = NextValue
End Function
