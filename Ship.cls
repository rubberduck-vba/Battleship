VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Default instance factory; avoid altering default instance state."
'@Folder("Battleship.Model.Ship")
'@Description("Default instance factory; avoid altering default instance state.")
Option Explicit

Private ShipSizes As Dictionary
Private shipNames As Dictionary

Private Const ShipNameCarrier As String = "Aircraft Carrier"
Private Const ShipNameBattleship As String = "Battleship"
Private Const ShipNameSubmarine As String = "Submarine"
Private Const ShipNameCruiser As String = "Cruiser"
Private Const ShipNameDestroyer As String = "Destroyer"

Private Type TShip
    ShipKind As ShipType
    Name As String
    GridPosition As IGridCoord
    Orientation As ShipOrientation
    State As Dictionary
    IsHit As Boolean
End Type

Private this As TShip
Implements IShip

'@Description("Populates and returns a dictionary associating all ship names with their respective size.")
Public Function Fleet() As Dictionary
    Dim Names As Variant
    Names = shipNames.Items
    
    Dim sizes As Variant
    sizes = ShipSizes.Items
    
    Dim result As Dictionary
    Set result = New Dictionary
    Dim i As Long
    For i = LBound(Names) To UBound(Names)
        result.Add Names(i), sizes(i)
    Next
    Set Fleet = result
End Function

'@Description("Gets an array of all valid ShipKind enum values.")
Public Function ShipKinds() As Variant
    ShipKinds = shipNames.Keys
End Function

'@Description("Gets an array of all ship names.")
Public Function Names() As Variant
    Names = shipNames.Items
End Function

'@Description("Use from the class' default instance to create a new ship instance using parameters.")
Public Function Create(ByVal kind As ShipType, ByVal direction As ShipOrientation, ByVal position As IGridCoord) As Ship
    ValidateInputs kind, direction, position
    With New Ship
        .ShipKind = kind
        .Name = shipNames(kind)
        .Orientation = direction
        Set .GridPosition = position
        Dim Offset As Byte
        For Offset = 0 To ShipSizes(kind) - 1
            
            Dim currentPoint As GridCoord
            Set currentPoint = GridCoord.Default
            
            currentPoint.X = position.X + IIf(direction = Horizontal, Offset, 0)
            currentPoint.Y = position.Y + IIf(direction = Vertical, Offset, 0)
            
            ' each element is a Boolean, keyed with a grid coordinate:
            .State.Add Item:=False, Key:=currentPoint.ToString
        Next
        Set Create = .Self
    End With
End Function

Private Sub ValidateInputs(ByVal kind As ShipType, ByVal Orientation As ShipOrientation, ByVal position As GridCoord)
    Dim shipSize As Byte
    shipSize = ShipSizes(kind)
    
    Select Case True
        
        Case Orientation <> Horizontal And Orientation <> Vertical
            OnInvalidArgument "orientation", "Invalid orientation."
        
        Case Orientation = Horizontal And position.X + shipSize - 1 > PlayerGrid.Size
            OnInvalidArgument "position", "Invalid position; ship exceeds right edge of the grid."
        
        Case Orientation = Vertical And position.Y + shipSize - 1 > PlayerGrid.Size
            OnInvalidArgument "position", "Invalid position; ship exceeds bottom edge of the grid."
    
    End Select

End Sub

Private Sub OnInvalidArgument(ByVal argName As String, ByVal message As String)
    Err.Raise 5, TypeName(Me), message
End Sub

Public Property Get Self() As Ship
    Set Self = Me
End Property

Public Property Get ShipKind() As ShipType
    ShipKind = this.ShipKind
End Property

Public Property Let ShipKind(ByVal Value As ShipType)
    this.ShipKind = Value
End Property

Public Property Get Name() As String
    Name = this.Name
End Property

Public Property Let Name(ByVal Value As String)
    this.Name = Value
End Property

Public Property Get Orientation() As ShipOrientation
    Orientation = this.Orientation
End Property

Public Property Let Orientation(ByVal Value As ShipOrientation)
    this.Orientation = Value
End Property

Public Property Get GridPosition() As GridCoord
    Set GridPosition = this.GridPosition
End Property

Public Property Set GridPosition(ByVal Value As GridCoord)
    Set this.GridPosition = Value
End Property

Public Property Get State() As Dictionary
    Set State = this.State
End Property

Private Sub Class_Initialize()
    If Me Is Ship Then
        'default instance
        Set ShipSizes = New Dictionary
        With ShipSizes
            .Add ShipType.Carrier, 5
            .Add ShipType.Battleship, 4
            .Add ShipType.Submarine, 3
            .Add ShipType.Cruiser, 3
            .Add ShipType.Destroyer, 2
        End With
        Set shipNames = New Dictionary
        With shipNames
            .Add ShipType.Carrier, ShipNameCarrier
            .Add ShipType.Battleship, ShipNameBattleship
            .Add ShipType.Submarine, ShipNameSubmarine
            .Add ShipType.Cruiser, ShipNameCruiser
            .Add ShipType.Destroyer, ShipNameDestroyer
        End With
    Else
        Set this.State = New Dictionary
    End If
End Sub

Private Sub Class_Terminate()
    Set ShipSizes = Nothing
    Set shipNames = Nothing
    Set this.State = Nothing
End Sub

Private Property Get IShip_GridPosition() As IGridCoord
    Set IShip_GridPosition = this.GridPosition
End Property

Private Function IShip_Hit(ByVal coord As IGridCoord) As Boolean
    Dim coordString As String
    coordString = coord.ToString
    If this.State.Exists(coordString) Then
        'this.State.Remove coordString
        this.State(coordString) = True
        this.IsHit = True
        IShip_Hit = this.State(coordString)
    End If
End Function

Private Function IShip_Intersects(ByVal shipSize As Byte, ByVal direction As ShipOrientation, ByVal position As IGridCoord) As IGridCoord
    Dim gridOffset As Long
    For gridOffset = 0 To shipSize - 1
        Dim current As GridCoord
        Set current = position.Offset( _
            IIf(direction = Horizontal, gridOffset, 0), _
            IIf(direction = Vertical, gridOffset, 0))
        If this.State.Exists(current.ToString) Then
            Set IShip_Intersects = current
            Exit Function
        End If
    Next
End Function

Private Property Get IShip_HitAreas() As VBA.Collection
    Dim result As VBA.Collection
    Set result = New VBA.Collection
    Dim currentArea As VBA.Collection
    Set currentArea = New VBA.Collection
    Dim currentPoint As Variant
    For Each currentPoint In this.State.Keys
        If this.State(currentPoint) Then
            currentArea.Add GridCoord.FromString(currentPoint)
        Else
            If currentArea.Count > 0 Then
                result.Add currentArea
                Set currentArea = New VBA.Collection
            End If
        End If
    Next
    If currentArea.Count > 0 Then result.Add currentArea
    Set IShip_HitAreas = result
End Property

Private Property Get IShip_IsSunken() As Boolean
    If Not this.IsHit Then Exit Property
    Dim currentPoint As Variant
    For Each currentPoint In this.State.Items
        If Not currentPoint Then Exit Property
    Next
    IShip_IsSunken = True
End Property

Private Property Get IShip_Name() As String
    IShip_Name = this.Name
End Property

Private Property Get IShip_Orientation() As ShipOrientation
    IShip_Orientation = this.Orientation
End Property

Private Property Get IShip_ShipKind() As ShipType
    IShip_ShipKind = this.ShipKind
End Property

Private Property Get IShip_Size() As Byte
    IShip_Size = this.State.Count
End Property

Private Property Get IShip_StateArray() As Variant
    IShip_StateArray = this.State.Items
End Property

