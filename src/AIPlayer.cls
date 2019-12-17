VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AIPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IPlayer that is AI-controlled."
'@Folder("Battleship.Model")
'@ModuleDescription("An implementation of IPlayer that is AI-controlled.")
'@PredeclaredId
Option Explicit
Implements IPlayer

Private Const Delay As Long = 800

Private Type TPlayer
    PlayerType As PlayerType
    PlayGrid As PlayerGrid
    Strategy As IGameStrategy
End Type

Private this As TPlayer

Public Function Create(ByVal grid As PlayerGrid, ByVal GameStrategy As IGameStrategy) As IPlayer
    With New AIPlayer
        .PlayerType = ComputerControlled
        Set .Strategy = GameStrategy
        Set .PlayGrid = grid
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IPlayer
    Set Self = Me
End Property

Public Property Get Strategy() As IGameStrategy
    Set Strategy = this.Strategy
End Property

Public Property Set Strategy(ByVal value As IGameStrategy)
    Set this.Strategy = value
End Property

Public Property Get PlayGrid() As PlayerGrid
    Set PlayGrid = this.PlayGrid
End Property

Public Property Set PlayGrid(ByVal value As PlayerGrid)
    Set this.PlayGrid = value
End Property

Public Property Get PlayerType() As PlayerType
    PlayerType = this.PlayerType
End Property

Public Property Let PlayerType(ByVal value As PlayerType)
    this.PlayerType = value
End Property

Private Property Get IPlayer_PlayGrid() As PlayerGrid
    Set IPlayer_PlayGrid = this.PlayGrid
End Property

Private Sub IPlayer_PlaceShip(ByVal CurrentShip As IShip)
    this.Strategy.PlaceShip this.PlayGrid, CurrentShip
End Sub

Private Function IPlayer_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    
    Win32API.Sleep Delay
    Set IPlayer_Play = this.Strategy.Play(enemyGrid)
    
    Static shots As Long
    shots = shots + 1
    Debug.Print "AI Player " & this.PlayGrid.gridId & "(" & TypeName(this.Strategy) & ") has played " & shots & " turns"
    
End Function

Private Property Get IPlayer_PlayerType() As PlayerType
    IPlayer_PlayerType = this.PlayerType
End Property
