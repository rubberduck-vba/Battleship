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
'@PredeclaredId
'@Folder("Tests.Stubs")
'@ModuleDescription("A configurable, deterministic, mock implementation of a randomizer.")
Option Explicit
Implements IRandomizer

Private Type TState
    currentIndex As Long
    sequence As VBA.Collection
    MaybeInvokes As VBA.Collection
End Type

Private this As TState

Public Sub Setup(ParamArray sequence() As Variant)
    
    Dim i As Long
    Dim gridCoords As Long
    For i = LBound(sequence) To UBound(sequence)
        Dim value As Variant
        value = sequence(i)
        
        If TypeOf value Is IGridCoord Then
            
            Dim position As IGridCoord
            Set position = value
            
            this.sequence.Add position.X
            this.sequence.Add position.Y
        
        Else
            this.sequence.Add CSng(sequence(i))
        End If
        
    Next
    this.currentIndex = 0
End Sub

Public Property Get NextValue() As Single
    If this.currentIndex = this.sequence.Count Then Reset
    
    this.currentIndex = this.currentIndex + 1
    NextValue = this.sequence(this.currentIndex)
End Property

Private Sub Reset()
    this.currentIndex = 0
End Sub

Private Sub Class_Initialize()
    Set this.sequence = New VBA.Collection
    Set this.MaybeInvokes = New VBA.Collection
End Sub

Private Function IRandomizer_Between(ByVal lower As Long, ByVal upper As Long) As Long
    IRandomizer_Between = NextValue
End Function

Private Function IRandomizer_Maybe(ByVal likelihood As Probability) As Boolean
    MaybeInvokes.Add likelihood
    IRandomizer_Maybe = Me.NextValue < GameRandomizer.Probability(likelihood)
End Function

Public Property Get MaybeInvokes() As VBA.Collection
    Dim result As VBA.Collection
    Set result = New VBA.Collection
    Dim i As Variant
    For Each i In this.MaybeInvokes
        result.Add i
    Next
    Set MaybeInvokes = result
End Property

Private Function IRandomizer_NextSingle() As Single
    IRandomizer_NextSingle = NextValue
End Function
