VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PlayerGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "An object representing a player's game grid."
'@PredeclaredId
'@Folder("Battleship.Model")
'@ModuleDescription("An object representing a player's game grid.")
'@Exposed
Option Explicit

Private Const GridSize As Byte = 10
Private Const MaxShipsPerGrid As Byte = 5

Private Const KnownGridStateErrorMsg As String = "Specified coordinate is not in an unknown state."
Private Const CannotAddShipAtPositionMsg As String = "Cannot add a ship of this size at this position."
Private Const CannotAddMoreShipsMsg As String = "Cannot add more ships to this grid."

Public Enum PlayGridId
    PlayerGrid1 = 1
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

Private Type TPlayGrid
    Id As PlayGridId
    ships As Collection
    State(1 To GridSize, 1 To GridSize) As GridState
End Type

Private this As TPlayGrid

Public Function Create(ByVal gridId As PlayGridId) As PlayerGrid
    With New PlayerGrid
        .gridId = gridId
        Set Create = .Self
    End With
End Function

Public Property Get Self() As PlayerGrid
    Set Self = Me
End Property

'@Description("Gets the size of a player game grid.")
Public Property Get Size() As Byte
Attribute Size.VB_Description = "Gets the size of a player game grid."
    Size = GridSize
End Property

'@Description("Gets the number of ships in a player's grid at the end of the deployment phase.")
Public Property Get ShipsPerGrid() As Byte
Attribute ShipsPerGrid.VB_Description = "Gets the number of ships in a player's grid at the end of the deployment phase."
    ShipsPerGrid = MaxShipsPerGrid
End Property

'@Description("Gets the ID of this grid. 1 for Player1, 2 for Player2.")
Public Property Get gridId() As PlayGridId
Attribute gridId.VB_Description = "Gets the ID of this grid. 1 for Player1, 2 for Player2."
    gridId = this.Id
End Property

Public Property Let gridId(ByVal value As PlayGridId)
    this.Id = value
End Property

'@Description("Gets the number of ships placed on the grid.")
Public Property Get shipCount() As Long
Attribute shipCount.VB_Description = "Gets the number of ships placed on the grid."
    shipCount = this.ships.Count
End Property

'@Description("Gets a collection containing all ships on this grid, sunken or afloat.")
Public Property Get Fleet() As Collection
Attribute Fleet.VB_Description = "Gets a collection containing all ships on this grid, sunken or afloat."
    Dim result As Collection
    Set result = New Collection
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        result.Add CurrentShip
    Next
    Set Fleet = result
End Property

'@Description("Gets the size of the smallest ship still afloat on this grid.")
Public Property Get SmallestShipSize() As Byte
Attribute SmallestShipSize.VB_Description = "Gets the size of the smallest ship still afloat on this grid."
    Dim result As Byte
    Dim CurrentShip As IShip
    Dim smallestShip As IShip
    For Each CurrentShip In this.ships
        If Not CurrentShip.IsSunken Then
            If smallestShip Is Nothing Then
                Set smallestShip = CurrentShip
            Else
                If CurrentShip.Size < smallestShip.Size Then Set smallestShip = CurrentShip
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
Attribute AddShip.VB_Description = "Adds the specified ship to the grid. Throws if position is illegal."
    
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
Attribute CanAddShip.VB_Description = "Gets a value indicating whether a ship can be added at the specified position/direction/size."
    CanAddShip = (position.X + IIf(direction = Horizontal, shipSize - 1, 0) <= UBound(this.State, 1)) _
             And (position.Y + IIf(direction = Vertical, shipSize - 1, 0) <= UBound(this.State, 2)) _
             And (position.X > 0 And position.Y > 0) _
             And IntersectsAny(position, direction, shipSize) Is Nothing
End Function

'@Description("Gets a value indicating whether the specified position/direction/size intersects with any existing ship.")
Public Function IntersectsAny(ByVal position As GridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As GridCoord
Attribute IntersectsAny.VB_Description = "Gets a value indicating whether the specified position/direction/size intersects with any existing ship."
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        Dim intersecting As GridCoord
        Set intersecting = CurrentShip.Intersects(shipSize, direction, position)
        If Not intersecting Is Nothing Then
            Set IntersectsAny = intersecting
            Exit Function
        End If
    Next
End Function

'@Description("Gets a value indicating whether the specified position/direction/size has any adjacent existing ship.")
Public Function HasAdjacentShip(ByVal position As GridCoord, ByVal direction As ShipOrientation, ByVal shipSize As Byte) As Boolean
Attribute HasAdjacentShip.VB_Description = "Gets a value indicating whether the specified position/direction/size has any adjacent existing ship."
        
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
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        If CurrentShip.Orientation = Horizontal Then
            currentY = CurrentShip.GridPosition.Y
            For currentX = CurrentShip.GridPosition.X To CurrentShip.GridPosition.X + CurrentShip.Size - 1
                If GridCoord.Create(currentX, currentY).IsAdjacent(coord) Then
                    HasAnyAdjacentShips = True
                    Exit Function
                End If
            Next
        Else
            currentX = CurrentShip.GridPosition.X
            For currentY = CurrentShip.GridPosition.Y To CurrentShip.GridPosition.Y + CurrentShip.Size - 1
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
Attribute TryHit.VB_Description = "(side-effecting) Attempts a hit at the specified position; returns the result of the attack, and a reference to the hit ship if successful."
    
    If this.State(position.X, position.Y) = GridState.PreviousHit Or _
       this.State(position.X, position.Y) = GridState.PreviousMiss Then
        Err.Raise PlayerGridErrors.KnownGridStateError, TypeName(Me), KnownGridStateErrorMsg
    End If
    
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        If CurrentShip.Hit(position) Then
            this.State(position.X, position.Y) = GridState.PreviousHit
            If CurrentShip.IsSunken Then
                TryHit = Sunk
            Else
                TryHit = Hit
            End If
            Set hitShip = CurrentShip
            Exit Function
        End If
    Next
    
    this.State(position.X, position.Y) = GridState.PreviousMiss
    TryHit = Miss
    
End Function

'@Description("True if specified position contains a ship that was previously hit, but not sunken.")
Public Property Get HasDamagedShip(ByVal position As GridCoord) As Boolean
Attribute HasDamagedShip.VB_Description = "True if specified position contains a ship that was previously hit, but not sunken."
    
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        
        Dim currentX As Long
        Dim currentY As Long
        
        If CurrentShip.GridPosition.Y = position.Y And CurrentShip.Orientation = Horizontal Then
            For currentX = CurrentShip.GridPosition.X To CurrentShip.GridPosition.X + CurrentShip.Size - 1
                If currentX = position.X Then
                    HasDamagedShip = Not CurrentShip.IsSunken And CurrentShip.HitAreas.Count > 0
                    Exit Property
                End If
            Next
        ElseIf CurrentShip.GridPosition.X = position.X Then
            For currentY = CurrentShip.GridPosition.Y To CurrentShip.GridPosition.Y + CurrentShip.Size - 1
                If currentY = position.Y Then
                    HasDamagedShip = Not CurrentShip.IsSunken And CurrentShip.HitAreas.Count > 0
                    Exit Property
                End If
            Next
        End If
        
    Next
    
    HasDamagedShip = False
    
End Property

'@Description("Gets the GridState value at the specified position.")
Public Property Get State(ByVal position As GridCoord) As GridState
Attribute State.VB_Description = "Gets the GridState value at the specified position."
    On Error Resume Next
    State = this.State(position.X, position.Y)
    On Error GoTo 0
End Property

'@Description("Gets a 2D array containing the GridState of each coordinate in the grid.")
Public Property Get StateArray() As Variant
Attribute StateArray.VB_Description = "Gets a 2D array containing the GridState of each coordinate in the grid."
    Dim result(1 To GridSize, 1 To GridSize) As Variant
    Dim currentX As Long
    For currentX = 1 To GridSize
        Dim currentY As Long
        For currentY = 1 To GridSize
            Dim value As GridState
            value = this.State(currentX, currentY)
            result(currentX, currentY) = IIf(value = Unknown, Empty, value)
        Next
    Next
    StateArray = result
End Property

'@Description("Gets a value indicating whether the ship at the specified position is sunken.")
Public Property Get IsSunken(ByVal position As GridCoord) As Boolean
Attribute IsSunken.VB_Description = "Gets a value indicating whether the ship at the specified position is sunken."
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        If CurrentShip.IsSunken Then
            If CurrentShip.Orientation = Horizontal Then
                If CurrentShip.GridPosition.Y = position.Y Then
                    If position.X >= CurrentShip.GridPosition.X And _
                       position.X <= CurrentShip.GridPosition.X + CurrentShip.Size - 1 _
                    Then
                        IsSunken = True
                        Exit Property
                    End If
                End If
            End If
        End If
    Next
End Property

'@Description("Gets a value indicating whether all ships have been sunken.")
Public Property Get IsAllSunken() As Boolean
Attribute IsAllSunken.VB_Description = "Gets a value indicating whether all ships have been sunken."
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        If Not CurrentShip.IsSunken Then
            IsAllSunken = False
            Exit Property
        End If
    Next
    IsAllSunken = True
End Property

'@Description("Finds area around a damaged ship, if one exists.")
Public Function FindHitArea() As Collection
Attribute FindHitArea.VB_Description = "Finds area around a damaged ship, if one exists."
    Dim CurrentShip As IShip
    For Each CurrentShip In this.ships
        If Not CurrentShip.IsSunken Then
            Dim currentAreas As Collection
            Set currentAreas = CurrentShip.HitAreas
            If currentAreas.Count > 0 Then
                Set FindHitArea = currentAreas(1)
                Exit Function
            End If
        End If
    Next
End Function

'@Description("Removes confirmed ship positions from grid state.")
Public Sub Scramble()
Attribute Scramble.VB_Description = "Removes confirmed ship positions from grid state."
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


