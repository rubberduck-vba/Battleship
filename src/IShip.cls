VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IShip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object representing a ship."
'@Exposed
'@Folder("Battleship.Model.Abstract")
'@ModuleDescription("Describes an object representing a ship.")
'@Interface
Option Explicit

Public Enum ShipType
    Carrier
    Battleship
    Submarine
    Cruiser
    Destroyer
End Enum

Public Enum ShipOrientation
    Horizontal
    Vertical
End Enum

'@Description("Gets the type of the ship.")
Public Property Get ShipKind() As ShipType
Attribute ShipKind.VB_Description = "Gets the type of the ship."
End Property

'@Description("The name/description of the ship. Must be unique in a grid.")
Public Property Get Name() As String
Attribute Name.VB_Description = "The name/description of the ship. Must be unique in a grid."
End Property

'@Description("Use in 'With' blocks to get a reference to the scope variable.")
Public Property Get GridPosition() As IGridCoord
Attribute GridPosition.VB_Description = "Use in 'With' blocks to get a reference to the scope variable."
End Property

'@Description("The number of grid squares (1-5) occupied by this ship.")
Public Property Get Size() As Byte
Attribute Size.VB_Description = "The number of grid squares (1-5) occupied by this ship."
End Property

'@Description("The orientation of the ship.")
Public Property Get Orientation() As ShipOrientation
Attribute Orientation.VB_Description = "The orientation of the ship."
End Property

'@Description("True if this ship is sunken.")
Public Property Get IsSunken() As Boolean
Attribute IsSunken.VB_Description = "True if this ship is sunken."
End Property

'@Description("Gets a collection of collections containing contiguous grid coordinates this ship was hit at.")
Public Property Get HitAreas() As Collection
Attribute HitAreas.VB_Description = "Gets a collection of collections containing contiguous grid coordinates this ship was hit at."
End Property

'@Description("Gets an array containing the state of each grid coordinate of the ship.")
Public Property Get StateArray() As Variant
Attribute StateArray.VB_Description = "Gets an array containing the state of each grid coordinate of the ship."
End Property

'@Description("If the specified coordinate hits this ship, marks coordinate as a hit and returns True.")
Public Function Hit(ByVal coord As IGridCoord) As Boolean
Attribute Hit.VB_Description = "If the specified coordinate hits this ship, marks coordinate as a hit and returns True."
End Function

'@Description("Returns intersection coordinate if specified ship intersects with this instance.")
Public Function Intersects(ByVal shipSize As Byte, ByVal direction As ShipOrientation, ByVal position As IGridCoord) As IGridCoord
Attribute Intersects.VB_Description = "Returns intersection coordinate if specified ship intersects with this instance."
End Function

