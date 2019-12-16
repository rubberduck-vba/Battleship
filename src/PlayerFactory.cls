VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "A concrete implementation for the IPlayer abstract factory."
'@Folder("Battleship.Model")
'@ModuleDescription("A concrete implementation for the IPlayer abstract factory.")
'@PredeclaredId
Option Explicit

Private Type TFactory
    Random As IRandomizer
End Type

Private this As TFactory
Implements IPlayerFactory

Public Function Create(ByVal randomizer As IRandomizer) As IPlayerFactory
    With New PlayerFactory
        Set .Random = randomizer
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IPlayerFactory
    Set Self = Me
End Property

Public Property Get Random() As IRandomizer
    Set Random = this.Random
End Property

Public Property Set Random(ByVal value As IRandomizer)
    Set this.Random = value
End Property

Private Function IPlayerFactory_Create(ByVal gridId As PlayGridId, Optional ByVal difficulty As AIDifficulty = AIDifficulty.Unspecified) As IPlayer
    
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(gridId)
    
    If difficulty = Unspecified Then
        Set IPlayerFactory_Create = CreateHumanPlayer(grid)
    Else
        Set IPlayerFactory_Create = CreateAIPlayer(grid, difficulty)
    End If

End Function

Private Function CreateHumanPlayer(ByVal grid As PlayerGrid) As IPlayer
    Set CreateHumanPlayer = HumanPlayer.Create(grid)
End Function

Private Function CreateAIPlayer(ByVal grid As PlayerGrid, ByVal difficulty As AIDifficulty) As IPlayer
    Select Case difficulty
        Case AIDifficulty.RandomAI
            Set CreateAIPlayer = AIPlayer.Create(grid, RandomShotStrategy.Create(this.Random))
            
        Case AIDifficulty.FairplayAI
            Set CreateAIPlayer = AIPlayer.Create(grid, FairPlayStrategy.Create(this.Random))
            
        Case AIDifficulty.MercilessAI
            Set CreateAIPlayer = AIPlayer.Create(grid, MercilessStrategy.Create(this.Random))
            
        Case Else
            Debug.Assert False
    End Select
End Function
