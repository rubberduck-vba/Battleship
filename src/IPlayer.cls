VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object representing a player."
'@Exposed
'@Folder("Battleship.Model.Abstract")
'@ModuleDescription("Describes an object representing a player.")
'@Interface

Public Enum PlayerType
    HumanControlled
    ComputerControlled
End Enum

Option Explicit

'@Description("Identifies whether the player is human or computer-controlled.")
Public Property Get PlayerType() As PlayerType
Attribute PlayerType.VB_Description = "Identifies whether the player is human or computer-controlled."
End Property

'@Description("Gets the player's grid/state.")
Public Property Get PlayGrid() As PlayerGrid
Attribute PlayGrid.VB_Description = "Gets the player's grid/state."
End Property

'@Description("Places specified ship on game grid.")
Public Sub PlaceShip(ByVal CurrentShip As IShip)
Attribute PlaceShip.VB_Description = "Places specified ship on game grid."
End Sub

'@Description("Attempts to make a hit on the enemy grid.")
Public Function Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
Attribute Play.VB_Description = "Attempts to make a hit on the enemy grid."
End Function


