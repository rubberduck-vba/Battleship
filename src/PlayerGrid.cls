VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Battleship.Model.Player")
Option Explicit

Private Const GridSize As Byte = 10
Private Const MaxShipsPerGrid As Byte = 5

Private Const KnownGridStateErrorMsg As String _
    = "Specified coordinate is not in an unknown state."
Private Const CannotAddShipAtPositionMsg As String _
    = "Cannot add a ship of this size at this position."
Private Const CannotAddMoreShipsMsg As String _
    = "Cannot add more ships to this grid."

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
    '@Description("Content at this coordinate is unknown.")
    Unknown = -1
    '@Description("Unconfirmed friendly ship position.")
    PreviewShipPosition = 0
    '@Description("Confirmed friendly ship position.")
    ShipPosition = 1
    '@Description("Unconfirmed invalid/overlapping ship position.")
    InvalidPosition = 2
    '@Description("No ship at this coordinate.")
    PreviousMiss = 3
    '@Description("An enemy ship occupies this coordinate.")
    PreviousHit = 4
End Enum

Private Type TPlayGrid
    Id As Byte
    ships As Collection
    State(1 To GridSize, 1 To GridSize) As GridState
End Type

Private this As TPlayGrid

Public Function Create(ByVal gridId As Byte) As PlayerGrid
    With New PlayerGrid
        .gridId = gridId
        Set Create = .Self
    End With
End Function

Public Property Get Self() As PlayerGrid
    Set Self = Me
End Property

Public Property Get Size() As Byte
    Size = GridSize
End Property

Public Property Get ShipsPerGrid() As Byte
    ShipsPerGrid = MaxShipsPerGrid
End Property

'@Description("Gets the ID of this grid. 1 for Player1, 2 for Player2.")
Public Property Get gridId() As Byte
    gridId = this.Id
End Property

Public Property Let gridId(ByVal Value As Byte)
    this.Id = Value
End Property

'@Description("Gets the number of ships placed on the grid.")
Public Property Get shipCount() As Long
    shipCount = this.ships.Count
End Property

'@Description("Gets a collection containing all ships on this grid.")
Public Property Get Fleet() As Collection
    Dim result As Collection
    Set result = New Collection
    Dim currentShip As IShip
    For Each currentShip In this.ships
        result.Add currentShip
    Next
    Set Fleet = result
End Property

Public Property Get SmallestShipSize() As Byte
    Dim result As Byte
    Dim currentShip As IShip
    Dim smallestShip As IShip
    For Each currentShip In this.ships
        If Not currentShip.IsSunken Then
            If smallestShip Is Nothing Then
                Set smallestShip = currentShip
            Else
                If currentShip.Size < smallestShip.Size Then Set smallestShip = currentShip
                If smallestShip.Size = 2 Then Exit For
            End If
        End If
    Next
    SmallestShipSize = smallestShip.Size
End Property

Private Sub Class_Initialize()
    Set this.ships = New Collection
    Dim currentX As Long
    For currentX = LBound(this.State, 1) To UBound(this.State, 1)
        Dim currentY As Long
        For currentY = LBound(this.State, 2) To UBound(this.State, 2)
            this.State(currentX, currentY) = GridState.Unknown
        Next
    Next
End Sub

'@Description("Adds the specified ship to the grid. Throws if position is illegal.")
Public Sub AddShip(ByVal Item As IShip)
    
    If Not CanAddShip(Item.GridPosition, Item.Orientation, Item.Size) Then
        Err.Raise PlayerGridErrors.CannotAddShipAtPosition, TypeName(Me), CannotAddShipAtPositionMsg
    End If
    
    If this.ships.Count >= ShipsPerGrid Then
        Err.Raise PlayerGridErrors.CannotAddMoreShips, TypeName(Me), CannotAddMoreShipsMsg
    End If
    
    ' will throw a duplicate key error if item.Name is already in collection
    this.ships.Add Item, Item.Name
    
    Dim currentX As Long
    For currentX = Item.GridPosition.X To Item.GridPosition.X + IIf(Item.Orientation = Horizontal, Item.Size - 1, 0)
        Dim currentY As Long
        For currentY = Item.GridPosition.Y To Item.GridPosition.Y + IIf(Item.Orientation = Vertical, Item.Size - 1, 0)
            this.State(currentX, currentY) = GridState.ShipPosition
        Next
    Next

End Sub

'@Description("Gets a value indicating whether a ship can be added at the specified position/direction/size.")
Public Function CanAddShip(ByVal position As GridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As Boolean
    CanAddShip = (position.X + IIf(direction = Horizontal, shipSize - 1, 0) <= UBound(this.State, 1)) _
             And (position.Y + IIf(direction = Vertical, shipSize - 1, 0) <= UBound(this.State, 2)) _
             And (position.X > 0 And position.Y > 0) _
             And IntersectsAny(position, direction, shipSize) Is Nothing
End Function

'@Description("Gets a value indicating whether the specified position/direction/size intersects with any existing ship.")
Public Function IntersectsAny(ByVal position As GridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As GridCoord
    Dim currentShip As IShip
    For Each currentShip In this.ships
        Dim intersecting As GridCoord
        Set intersecting = currentShip.Intersects(shipSize, direction, position)
        If Not intersecting Is Nothing Then
            Set IntersectsAny = intersecting
            Exit Function
        End If
    Next
End Function

'@Description("Gets a value indicating whether the specified position/direction/size has any adjacent existing ship.")
Public Function HasAdjacentShip(ByVal position As GridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As Boolean
        
    Dim positionX As Long
    Dim positionY As Long
    
    If direction = Horizontal Then
        positionY = position.Y
        For positionX = position.X To position.X + shipSize - 1
            If HasAnyAdjacentShips(GridCoord.Create(positionX, positionY)) Then
                HasAdjacentShip = True
                Exit Function
            End If
        Next
    Else
        positionX = position.X
        For positionY = position.Y To position.Y + shipSize - 1
            If HasAnyAdjacentShips(GridCoord.Create(positionX, positionY)) Then
                HasAdjacentShip = True
                Exit Function
            End If
        Next
    End If
End Function

Private Function HasAnyAdjacentShips(ByVal coord As GridCoord) As Boolean
    Dim currentX As Long
    Dim currentY As Long
    Dim currentShip As IShip
    For Each currentShip In this.ships
        If currentShip.Orientation = Horizontal Then
            currentY = currentShip.GridPosition.Y
            For currentX = currentShip.GridPosition.X To currentShip.GridPosition.X + currentShip.Size - 1
                If GridCoord.Create(currentX, currentY).IsAdjacent(coord) Then
                    HasAnyAdjacentShips = True
                    Exit Function
                End If
            Next
        Else
            currentX = currentShip.GridPosition.X
            For currentY = currentShip.GridPosition.Y To currentShip.GridPosition.Y + currentShip.Size - 1
                If GridCoord.Create(currentX, currentY).IsAdjacent(coord) Then
                    HasAnyAdjacentShips = True
                    Exit Function
                End If
            Next
        End If
    Next
End Function

'@Description("(side-effecting) Attempts a hit at the specified position; returns the result of the attack, and a reference to the hit ship if successful.")
Public Function TryHit(ByVal position As IGridCoord, Optional ByRef hitShip As IShip) As AttackResult
    
    If this.State(position.X, position.Y) = GridState.PreviousHit Or _
       this.State(position.X, position.Y) = GridState.PreviousMiss Then
        Err.Raise PlayerGridErrors.KnownGridStateError, TypeName(Me), KnownGridStateErrorMsg
    End If
    
    Dim currentShip As IShip
    For Each currentShip In this.ships
        If currentShip.Hit(position) Then
            this.State(position.X, position.Y) = GridState.PreviousHit
            If currentShip.IsSunken Then
                TryHit = Sunk
            Else
                TryHit = Hit
            End If
            Set hitShip = currentShip
            Exit Function
        End If
    Next
    
    this.State(position.X, position.Y) = GridState.PreviousMiss
    TryHit = Miss
    
End Function

'@Description("True if specified position contains a ship that was previously hit, but not sunken.")
Public Property Get HasDamagedShip(ByVal position As GridCoord) As Boolean
    
    Dim currentShip As IShip
    For Each currentShip In this.ships
        
        Dim currentX As Long
        Dim currentY As Long
        
        If currentShip.GridPosition.Y = position.Y And currentShip.Orientation = Horizontal Then
            For currentX = currentShip.GridPosition.X To currentShip.GridPosition.X + currentShip.Size - 1
                If currentX = position.X Then
                    HasDamagedShip = Not currentShip.IsSunken And currentShip.HitAreas.Count > 0
                    Exit Property
                End If
            Next
        ElseIf currentShip.GridPosition.X = position.X Then
            For currentY = currentShip.GridPosition.Y To currentShip.GridPosition.Y + currentShip.Size - 1
                If currentY = position.Y Then
                    HasDamagedShip = Not currentShip.IsSunken And currentShip.HitAreas.Count > 0
                    Exit Property
                End If
            Next
        End If
        
    Next
    
    HasDamagedShip = False
    
End Property

'@Description("Gets the GridState value at the specified position.")
Public Property Get State(ByVal position As GridCoord) As GridState
    On Error Resume Next
    State = this.State(position.X, position.Y)
    On Error GoTo 0
End Property

'@Description("Gets a 2D array containing the GridState of each coordinate in the grid.")
Public Property Get StateArray() As Variant
    Dim result(1 To GridSize, 1 To GridSize) As Variant
    Dim currentX As Long
    For currentX = 1 To GridSize
        Dim currentY As Long
        For currentY = 1 To GridSize
            Dim Value As GridState
            Value = this.State(currentX, currentY)
            result(currentX, currentY) = IIf(Value = Unknown, Empty, Value)
        Next
    Next
    StateArray = result
End Property

'@Description("Gets a value indicating whether the ship at the specified position is sunken.")
Public Property Get IsSunken(ByVal position As GridCoord) As Boolean
    Dim currentShip As IShip
    For Each currentShip In this.ships
        If currentShip.IsSunken Then
            If currentShip.Orientation = Horizontal Then
                If currentShip.GridPosition.Y = position.Y Then
                    If position.X >= currentShip.GridPosition.X And _
                       position.X <= currentShip.GridPosition.X + currentShip.Size - 1 _
                    Then
                        IsSunken = True
                        Exit Property
                    End If
                End If
            End If
        End If
    Next
End Property

'@Descrition("Gets a value indicating whether all ships have been sunken.")
Public Property Get IsAllSunken() As Boolean
    Dim currentShip As IShip
    For Each currentShip In this.ships
        If Not currentShip.IsSunken Then
            IsAllSunken = False
            Exit Property
        End If
    Next
    IsAllSunken = True
End Property

'@Description("Finds area around a damaged ship, if one exists.")
Public Function FindHitArea() As Collection
    Dim currentShip As IShip
    For Each currentShip In this.ships
        If Not currentShip.IsSunken Then
            Dim currentAreas As Collection
            Set currentAreas = currentShip.HitAreas
            If currentAreas.Count > 0 Then
                Set FindHitArea = currentAreas(1)
                Exit Function
            End If
        End If
    Next
End Function

'@Description("Removes confirmed ship positions from grid state.")
Public Sub Scramble()
    Dim currentX As Long
    For currentX = 1 To GridSize
        Dim currentY As Long
        For currentY = 1 To GridSize
            If this.State(currentX, currentY) = GridState.ShipPosition Then
                this.State(currentX, currentY) = GridState.Unknown
            End If
        Next
    Next
End Sub




