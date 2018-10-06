VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship.Model.Player")
Option Explicit
Implements IPlayer

Private Type TPlayer
    GridIndex As Byte
    PlayerType As PlayerType
    PlayGrid As PlayerGrid
End Type

Private this As TPlayer

Public Function Create(ByVal gridId As Byte) As IPlayer
    With New HumanPlayer
        .GridIndex = gridId
        .PlayerType = HumanControlled
        Set .PlayGrid = PlayerGrid.Create(gridId)
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IPlayer
    Set Self = Me
End Property

Public Property Get PlayGrid() As PlayerGrid
    Set PlayGrid = this.PlayGrid
End Property

Public Property Set PlayGrid(ByVal value As PlayerGrid)
    Set this.PlayGrid = value
End Property

Public Property Get GridIndex() As Byte
    GridIndex = this.GridIndex
End Property

Public Property Let GridIndex(ByVal value As Byte)
    this.GridIndex = value
End Property

Public Property Get PlayerType() As PlayerType
    PlayerType = this.PlayerType
End Property

Public Property Let PlayerType(ByVal value As PlayerType)
    this.PlayerType = value
End Property

Private Sub IPlayer_PlaceShip(ByVal currentShip As IShip)
    'human-controlled / no-op by design
End Sub

'@Ignore NonReturningFunction
Private Function IPlayer_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    'human-controlled / no-op by design
End Function

Private Property Get IPlayer_PlayGrid() As PlayerGrid
    Set IPlayer_PlayGrid = this.PlayGrid
End Property

Private Property Get IPlayer_PlayerType() As PlayerType
    IPlayer_PlayerType = this.PlayerType
End Property
