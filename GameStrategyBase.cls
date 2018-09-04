VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameStrategyBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Battleship.AI")
Option Explicit

Public Function PlaceShip(ByVal randomizer As IRandomizer, ByVal grid As PlayerGrid, ByVal currentShip As IShip, ByRef outDirection As ShipOrientation) As IGridCoord
    Do
        Dim gridX As Long
        gridX = randomizer.Between(1, PlayerGrid.Size)
        
        Dim direction As ShipOrientation
        If gridX + currentShip.Size - 1 > PlayerGrid.Size Then
            direction = Vertical
        Else
            direction = IIf(randomizer.NextSingle < 0.5, Horizontal, Vertical)
        End If
        
        Dim gridY As Long
        If direction = Horizontal Then
            gridY = randomizer.Between(1, PlayerGrid.Size)
        Else
            gridY = randomizer.Between(1, PlayerGrid.Size - currentShip.Size)
        End If
        
        Dim position As GridCoord
        Set position = GridCoord.Create(gridX, gridY)
        
    Loop Until grid.CanAddShip(position, direction, currentShip.Size)
    outDirection = direction
    Set PlaceShip = position
End Function

'@Description("Gets a random grid position with an unknown state, within the grid boundaries.")
Public Function ShootRandomPosition(ByVal randomizer As IRandomizer, ByVal grid As PlayerGrid) As IGridCoord
    Dim position As IGridCoord
    Do
        Set position = GridCoord.Create( _
            xPosition:=randomizer.Between(1, PlayerGrid.Size), _
            yPosition:=randomizer.Between(1, PlayerGrid.Size))
        
    Loop Until IsLegalPosition(grid, position)
    Set ShootRandomPosition = position
End Function

'@Description("Returns True if the specified position is a legal attack position.")
Public Function IsLegalPosition(ByVal grid As PlayerGrid, ByVal position As IGridCoord) As Boolean
    Dim isInsideGrid As Boolean
    If Not position Is Nothing Then
        isInsideGrid = position.X >= 1 And position.X <= PlayerGrid.Size And _
                       position.Y >= 1 And position.Y <= PlayerGrid.Size
        If isInsideGrid Then
            IsLegalPosition = grid.State(position) <> PreviousHit And _
                              grid.State(position) <> PreviousMiss
        End If
    End If
End Function

'@Description("Attempts to expand the provided area to destroy an identified target. May attempt to shoot outside the grid boundaries.")
Public Function DestroyTarget(ByVal randomizer As IRandomizer, ByVal grid As PlayerGrid, ByVal area As VBA.Collection) As IGridCoord

    Dim inferredDirection As ShipOrientation
    inferredDirection = TryInferDirection(randomizer, area)
    
    Dim result As IGridCoord
    
    If inferredDirection = Horizontal Then
        If randomizer.NextSingle < 0.5 Then
            Set result = FindLeftMostHit(area)
            If result.X > 1 Then Set result = result.Offset(xOffset:=-1)
            If result.X < 1 Or grid.State(result) = PreviousMiss Then
                Set result = FindRightMostHit(area).Offset(xOffset:=1)
            End If
        Else
            Set result = FindRightMostHit(area)
            If result.X < PlayerGrid.Size Then Set result = result.Offset(xOffset:=1)
            If result.X > PlayerGrid.Size Or grid.State(result) = PreviousMiss Then
                Set result = FindLeftMostHit(area).Offset(xOffset:=-1)
            End If
        End If
    Else
        If randomizer.NextSingle < 0.5 Then
            Set result = FindTopMostHit(area)
            If result.Y > 1 Then Set result = result.Offset(yOffset:=-1)
            If result.Y < 1 Or grid.State(result) = PreviousMiss Then
                Set result = FindBottomMostHit(area).Offset(yOffset:=1)
            End If
        Else
            Set result = FindBottomMostHit(area)
            If result.Y < PlayerGrid.Size Then Set result = result.Offset(yOffset:=1)
            If result.Y > PlayerGrid.Size Or grid.State(result) = PreviousMiss Then
                Set result = FindTopMostHit(area).Offset(yOffset:=-1)
            End If
        End If
    End If
                    
    Set DestroyTarget = result
    
End Function

'@Description("Attempts to infer a direction from the given area.")
Public Function TryInferDirection(ByVal randomizer As IRandomizer, ByVal area As Collection) As ShipOrientation
    
    Dim previousPosition As GridCoord
    Dim currentPosition As GridCoord
    
    For Each currentPosition In area
        If previousPosition Is Nothing Then
            Set previousPosition = currentPosition
            'could be either (ignoring remaining enemy ships)
            TryInferDirection = IIf(randomizer.NextSingle < 0.5, Horizontal, Vertical)
        Else
            If currentPosition.Y = previousPosition.Y Then
                TryInferDirection = Horizontal
                Exit Function
            Else
                TryInferDirection = Vertical
                Exit Function
            End If
        End If
    Next

End Function

'@Description("Gets the left-most hit in the specified area.")
Private Function FindLeftMostHit(ByVal area As Collection) As IGridCoord
    Dim leftMost As IGridCoord
    Set leftMost = area(1)
    
    Dim current As IGridCoord
    For Each current In area
        If current.X < leftMost.X Then Set leftMost = current
    Next
    
    Set FindLeftMostHit = leftMost
End Function

'@Description("Gets the right-most hit in the specified area.")
Private Function FindRightMostHit(ByVal area As Collection) As IGridCoord
    Dim rightMost As IGridCoord
    Set rightMost = area(1)
    
    Dim current As IGridCoord
    For Each current In area
        If current.X > rightMost.X Then Set rightMost = current
    Next
    
    Set FindRightMostHit = rightMost
End Function

'@Description("Gets the top-most hit in the specified area.")
Private Function FindTopMostHit(ByVal area As Collection) As IGridCoord
    Dim topMost As IGridCoord
    Set topMost = area(1)
    
    Dim current As IGridCoord
    For Each current In area
        If current.Y < topMost.Y Then Set topMost = current
    Next
    
    Set FindTopMostHit = topMost
End Function

'@Description("Gets the bottom-most hit in the specified area.")
Private Function FindBottomMostHit(ByVal area As Collection) As IGridCoord
    Dim bottomMost As IGridCoord
    Set bottomMost = area(1)
    
    Dim current As IGridCoord
    For Each current In area
        If current.Y > bottomMost.Y Then Set bottomMost = current
    Next
    
    Set FindBottomMostHit = bottomMost
End Function

'@Description("Returns True if the smallest remaining ship could fit at the specified position.")
Public Function VerifyShipFits(ByVal enemyGrid As PlayerGrid, ByVal position As IGridCoord, ByVal SmallestShipSize As Byte) As Boolean
    VerifyShipFits = FitsHorizontally(enemyGrid, position, SmallestShipSize) Or _
                     FitsVertically(enemyGrid, position, SmallestShipSize)
End Function

Private Function FitsHorizontally(ByVal enemyGrid As PlayerGrid, ByVal position As IGridCoord, ByVal SmallestShipSize As Byte) As Boolean
    
    Dim fits As Boolean
    Dim positionState As GridState
    
    Dim currentX As Long
    currentX = position.X
    
    If position.X + SmallestShipSize - 1 <= PlayerGrid.Size Then
        fits = True
        For currentX = position.X To position.X + SmallestShipSize - 1
            positionState = enemyGrid.State(GridCoord.Create(currentX, position.Y))
            If enemyGrid.HasDamagedShip(position) And positionState = Unknown Or positionState = ShipPosition Then
                fits = True
                Exit For
            Else
                fits = fits And positionState = Unknown
                If Not fits Then Exit For
            End If
        Next
        If fits Then
            FitsHorizontally = True
            Exit Function
        End If
    End If
    
    If position.X - SmallestShipSize + 1 < 1 Then Exit Function
    
    fits = True
    For currentX = position.X To position.X - SmallestShipSize + 1 Step -1
        positionState = enemyGrid.State(GridCoord.Create(currentX, position.Y))
        If enemyGrid.HasDamagedShip(position) And positionState = Unknown Or positionState = ShipPosition Then
            fits = True
            Exit For
        Else
            fits = fits And positionState = Unknown
        End If
    Next
    If fits Then
        FitsHorizontally = True
        Exit Function
    End If
    
End Function

Private Function FitsVertically(ByVal enemyGrid As PlayerGrid, ByVal position As IGridCoord, ByVal SmallestShipSize As Byte) As Boolean
    
    Dim fits As Boolean
    Dim positionState As GridState
    
    Dim currentY As Long
    currentY = position.Y
    
    If position.Y + SmallestShipSize - 1 <= PlayerGrid.Size Then
        fits = True
        For currentY = position.Y To position.Y + SmallestShipSize - 1
            positionState = enemyGrid.State(GridCoord.Create(position.X, currentY))
            If enemyGrid.HasDamagedShip(position) And positionState = Unknown Or positionState = ShipPosition Then
                fits = True
                Exit For
            Else
                fits = fits And positionState = Unknown
            End If
        Next
        If fits Then
            FitsVertically = True
            Exit Function
        End If
    End If
    
    If position.Y - SmallestShipSize + 1 < 1 Then Exit Function
    
    fits = True
    For currentY = position.Y To position.Y - SmallestShipSize + 1 Step -1
        positionState = enemyGrid.State(GridCoord.Create(position.X, currentY))
        If enemyGrid.HasDamagedShip(position) And positionState = Unknown Or positionState = ShipPosition Then
            fits = True
            Exit For
        Else
            fits = fits And positionState = Unknown
        End If
    Next
    If fits Then
        FitsVertically = True
        Exit Function
    End If
    
End Function



