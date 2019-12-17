VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlayerGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object representing a player's game grid."
'@Exposed
'@Folder("Battleship.Model.Abstract")
'@ModuleDescription("Describes an object representing a player's game grid.")
'@Interface
Option Explicit

Public Enum PlayGridId
    '@Ignore UseMeaningfulName
    PlayerGrid1 = 1
    '@Ignore UseMeaningfulName
    PlayerGrid2 = 2
End Enum

Public Enum PlayerGridErrors
    KnownGridStateError = vbObjectError Or 127
    CannotAddShipAtPosition
    CannotAddMoreShips
End Enum

Public Enum AttackResult
    Marked
    Miss
    Hit
    Sunk
End Enum

Public Enum GridState
    Unknown = -1
    PreviewShipPosition = 0
    ShipPosition = 1
    InvalidPosition = 2
    PreviousMiss = 3
    PreviousHit = 4
End Enum

'@Description("Gets the ID of this grid (Player1/Player2).")
'@Ignore UseMeaningfulName
Public Property Get Id() As PlayGridId
Attribute Id.VB_Description = "Gets the ID of this grid (Player1/Player2)."
End Property

'@Description("Gets the number of ships placed on the grid.")
Public Property Get ShipCount() As Long
Attribute ShipCount.VB_Description = "Gets the number of ships placed on the grid."
End Property

'@Description("Gets a collection containing all ships on this grid, sunken or afloat.")
Public Property Get Fleet() As VBA.Collection
Attribute Fleet.VB_Description = "Gets a collection containing all ships on this grid, sunken or afloat."
End Property

'@Description("Gets the size of the smallest ship still afloat on this grid.")
Public Property Get SmallestShipSize() As Byte
Attribute SmallestShipSize.VB_Description = "Gets the size of the smallest ship still afloat on this grid."
End Property

'@Description("Adds the specified ship to the grid. Throws if position is illegal.")
Public Sub AddShip(ByVal Item As IShip)
Attribute AddShip.VB_Description = "Adds the specified ship to the grid. Throws if position is illegal."
End Sub

'@Description("Gets a value indicating whether a ship can be added at the specified position/direction/size.")
Public Function CanAddShip(ByVal position As IGridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As Boolean
Attribute CanAddShip.VB_Description = "Gets a value indicating whether a ship can be added at the specified position/direction/size."
End Function

'@Description("Gets the IGridCoord of the intersecting coordinate where the specified position/direction/size is intersecting with another ship. Returns 'Nothing' if no other ships intersect.")
Public Function IntersectsAny(ByVal position As IGridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As IGridCoord
Attribute IntersectsAny.VB_Description = "Gets the IGridCoord of the intersecting coordinate where the specified position/direction/size is intersecting with another ship. Returns 'Nothing' if no other ships intersect."
End Function

'@Description("Gets a value indicating whether the specified position/direction/size has any adjacent existing ship.")
Public Function HasAdjacentShip(ByVal position As IGridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As Boolean
Attribute HasAdjacentShip.VB_Description = "Gets a value indicating whether the specified position/direction/size has any adjacent existing ship."
End Function

'@Description("Attempts a hit at the specified position; returns the result of the attack, and a reference to the hit ship if successful. Function is side-effecting: alters ship's hit state.")
Public Function TryHit(ByVal position As IGridCoord, Optional ByRef outShip As IShip) As AttackResult
Attribute TryHit.VB_Description = "Attempts a hit at the specified position; returns the result of the attack, and a reference to the hit ship if successful. Function is side-effecting: alters ship's hit state."
End Function

'@Description("True if specified position contains a ship that was previously hit, but not sunken.")
Public Property Get HasDamagedShip(ByVal position As IGridCoord) As Boolean
Attribute HasDamagedShip.VB_Description = "True if specified position contains a ship that was previously hit, but not sunken."
End Property

'@Description("Gets the GridState value at the specified position.")
Public Property Get State(ByVal position As IGridCoord) As GridState
Attribute State.VB_Description = "Gets the GridState value at the specified position."
End Property

'@Description("Gets a 2D array containing the GridState of each coordinate in the grid.")
Public Property Get StateArray() As Variant
Attribute StateArray.VB_Description = "Gets a 2D array containing the GridState of each coordinate in the grid."
End Property

'@Description("Gets a value indicating whether the ship at the specified position is sunken.")
Public Property Get IsSunken(ByVal position As IGridCoord) As Boolean
Attribute IsSunken.VB_Description = "Gets a value indicating whether the ship at the specified position is sunken."
End Property

'@Description("Gets a value indicating whether all ships have been sunken.")
Public Property Get IsAllSunken() As Boolean
Attribute IsAllSunken.VB_Description = "Gets a value indicating whether all ships have been sunken."
End Property

'@Description("Finds area around a damaged ship, if one exists.")
Public Function FindHitArea() As VBA.Collection
Attribute FindHitArea.VB_Description = "Finds area around a damaged ship, if one exists."
End Function

'@Description("Removes confirmed ship positions from grid state.")
Public Sub Scramble()
Attribute Scramble.VB_Description = "Removes confirmed ship positions from grid state."
End Sub

