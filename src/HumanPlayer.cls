VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "HumanPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "An implementation of IPlayer that is human-controlled."
'@Folder("Battleship.Model")
'@ModuleDescription("An implementation of IPlayer that is human-controlled.")
'@PredeclaredId
Option Explicit
Implements IPlayer

Private Type TPlayer
    PlayerType As PlayerType
    PlayGrid As PlayerGrid
End Type

Private this As TPlayer

'@Description("Creates a new HumanPlayer with the specified grid.")
Public Function Create(ByVal grid As PlayerGrid) As IPlayer
Attribute Create.VB_Description = "Creates a new HumanPlayer with the specified grid."
    With New HumanPlayer
        .PlayerType = HumanControlled
        Set .PlayGrid = grid
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

Public Property Get PlayerType() As PlayerType
    PlayerType = this.PlayerType
End Property

Public Property Let PlayerType(ByVal value As PlayerType)
    this.PlayerType = value
End Property

Private Sub IPlayer_PlaceShip(ByVal CurrentShip As IShip)
    'no-op, ship placement is handled by view events.
End Sub

Private Function IPlayer_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    'return value is irrelevant, play coord is provided by view events.
    Set IPlayer_Play = GridCoord.Default
End Function

Private Property Get IPlayer_PlayGrid() As PlayerGrid
    Set IPlayer_PlayGrid = this.PlayGrid
End Property

Private Property Get IPlayer_PlayerType() As PlayerType
    IPlayer_PlayerType = this.PlayerType
End Property
