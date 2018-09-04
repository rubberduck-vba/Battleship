VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Battleship.Model.Player")
'@Interface
Option Explicit

Public Enum PlayerType
    HumanControlled
    ComputerControlled
End Enum

'@Description("Identifies whether the player is human or computer-controlled.")
Public Property Get PlayerType() As PlayerType
End Property

'@Description("Gets the player's grid/state.")
Public Property Get PlayGrid() As PlayerGrid
End Property

'@Description("Places specified ship on game grid.")
Public Sub PlaceShip(ByVal currentShip As IShip)
End Sub

'@Description("Attempts to make a hit on the enemy grid.")
Public Function Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
End Function


