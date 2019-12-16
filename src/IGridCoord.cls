VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGridCoord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes a coordinate in a 2D grid."
'@Folder("Battleship.Model.Abstract")
'@ModuleDescription "Describes a coordinate in a 2D grid."
'@Exposed
'@Interface
Option Explicit

'@Description("Gets the horizontal position.")
Public Property Get X() As Long
Attribute X.VB_Description = "Gets the horizontal position."
End Property

'@Description("Gets the vertical position.")
Public Property Get Y() As Long
Attribute Y.VB_Description = "Gets the vertical position."
End Property

'@Description("Creates and returns a new coordinate by offsetting this instance.")
Public Function Offset(Optional ByVal xOffset As Long, Optional ByVal yOffset As Long) As IGridCoord
Attribute Offset.VB_Description = "Creates and returns a new coordinate by offsetting this instance."
End Function

'Description("Returns True if the specified coordinate is adjacent to this instance.")
Public Function IsAdjacent(ByVal other As IGridCoord) As Boolean
End Function

'@Description("Returns True if the specified coordinate describes the same location as this instance.")
Public Function Equals(ByVal other As IGridCoord) As Boolean
Attribute Equals.VB_Description = "Returns True if the specified coordinate describes the same location as this instance."
End Function

'@Description("Returns a (x,y) string representation of this instance.")
Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a (x,y) string representation of this instance."
End Function

'@Description("Returns a A1 string representation of this instance.
Public Function ToA1String() As String
End Function

