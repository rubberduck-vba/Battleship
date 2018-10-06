VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MercilessStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship.AI")
'@Description("A strategy that deploys ships more carefully, and attacks in patterns.")
Option Explicit
Implements IGameStrategy

Private Type TStrategy
    Random As IRandomizer
End Type

Private Enum GridEdge
    LeftEdge
    TopEdge
    RightEdge
    BottomEdge
End Enum

Private base As GameStrategyBase
Private this As TStrategy

Public Function Create(ByVal randomizer As IRandomizer) As IGameStrategy
    With New MercilessStrategy
        Set .Random = randomizer
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGameStrategy
    Set Self = Me
End Property

Public Property Get Random() As IRandomizer
    Set Random = this.Random
End Property

Public Property Set Random(ByVal value As IRandomizer)
    Set this.Random = value
End Property

Private Sub Class_Initialize()
    Set base = New GameStrategyBase
End Sub

Private Sub IGameStrategy_PlaceShip(ByVal grid As PlayerGrid, ByVal currentShip As IShip)
    Do
        Dim direction As ShipOrientation
        Dim position As IGridCoord
        Set position = base.PlaceShip(Random, grid, currentShip, direction)
    Loop Until Not grid.HasAdjacentShip(position, direction, currentShip.Size) Or Random.Maybe(AlmostNever)
    
    grid.AddShip Ship.Create(currentShip.ShipKind, direction, position)
    If grid.shipCount = PlayerGrid.ShipsPerGrid Then grid.Scramble
End Sub

Private Function IGameStrategy_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    Dim position As GridCoord
    Do
        Dim area As Collection
        Set area = enemyGrid.FindHitArea
        
        If Not area Is Nothing Then
            Set position = base.DestroyTarget(Random, enemyGrid, area)
        
        Else
            If this.Random.Maybe(AlmostNever) Then
                Set position = base.ShootRandomPosition(this.Random, enemyGrid)
                
            ElseIf this.Random.Maybe(Sometimes) Then
                Set position = ScanCenter(enemyGrid)
            
            Else
                Set position = ScanEdges(enemyGrid)
            
            End If
        End If
        
    Loop Until base.IsLegalPosition(enemyGrid, position) And _
               base.VerifyShipFits(enemyGrid, position, enemyGrid.SmallestShipSize) And _
               AvoidAdjacentHitPosition(enemyGrid, position)
    Set IGameStrategy_Play = position
End Function

Private Function ScanCenter(ByVal enemyGrid As PlayerGrid) As IGridCoord
    Dim position As IGridCoord
    Dim iterations As Long
    Static lastPosition As IGridCoord
    Do
        If iterations < 10 Then
            If lastPosition Is Nothing Then
                Set position = GridCoord.Create(this.Random.Between(3, 8), this.Random.Between(3, 8))
            Else
                If enemyGrid.State(lastPosition) = Unknown Then
                    'legal, but smallest ship wouldn't fit there. start over.
                    Set lastPosition = Nothing
                Else
                    Set position = lastPosition.Offset(IIf(this.Random.Maybe(Sometimes), 1, -1), IIf(this.Random.Maybe(Sometimes), 1, -1))
                    If position.X < 3 Or position.X > 8 Or position.Y < 3 Or position.Y > 8 Then
                        Set lastPosition = GridCoord.Create(this.Random.Between(3, 8), this.Random.Between(3, 8))
                    End If
                End If
            End If
        Else
            Set position = base.ShootRandomPosition(this.Random, enemyGrid)
        End If
        iterations = iterations + 1
    Loop Until base.IsLegalPosition(enemyGrid, position)
    Set lastPosition = position
    Set ScanCenter = position
End Function

Private Function ScanEdges(ByVal enemyGrid As PlayerGrid) As IGridCoord
    Dim position As IGridCoord
    Dim iterations As Long
    
    Static lastEdge As GridEdge
    Static lastPosition As IGridCoord
    
    Do
        If iterations < 10 Then
            If lastPosition Is Nothing Then
                If this.Random.Maybe(Sometimes) Then
                    'scan vertical edge
                    If this.Random.Maybe(Sometimes) Then
                        lastEdge = LeftEdge
                        Set position = GridCoord.Create(this.Random.Between(1, 2), this.Random.Between(1, 10))
                    Else
                        lastEdge = RightEdge
                        Set position = GridCoord.Create(this.Random.Between(9, 10), this.Random.Between(1, 10))
                    End If
                Else
                    'scan horizontal edge
                    If this.Random.Maybe(Sometimes) Then
                        lastEdge = TopEdge
                        Set position = GridCoord.Create(this.Random.Between(1, 10), this.Random.Between(1, 2))
                    Else
                        lastEdge = BottomEdge
                        Set position = GridCoord.Create(this.Random.Between(1, 10), this.Random.Between(9, 10))
                    End If
                    
                End If
            Else
                If enemyGrid.State(lastPosition) = Unknown Then
                    'legal, but smallest ship wouldn't fit there. start over.
                    Set lastPosition = Nothing
                Else
                    'allow suddenly jumping to another edge
                    If this.Random.Maybe(Rarely) Then lastEdge = this.Random.Between(0, 3)
                    
                    Select Case lastEdge
                        Case LeftEdge
                            Set position = GridCoord.Create(IIf(lastPosition.X = 1, 2, 1), IIf(lastPosition.Y + 1 > 10, 1, lastPosition.Y + 1))
                        Case RightEdge
                            Set position = GridCoord.Create(IIf(lastPosition.X = 10, 9, 10), IIf(lastPosition.Y - 1 < 1, 1, lastPosition.Y - 1))
                        Case TopEdge
                            Set position = GridCoord.Create(IIf(lastPosition.X + 1 > 10, 1, lastPosition.X + 1), IIf(lastPosition.Y = 1, 2, 1))
                        Case BottomEdge
                           Set position = GridCoord.Create(IIf(lastPosition.X - 1 < 1, 10, lastPosition.X - 1), IIf(lastPosition.Y = 10, 9, 10))
                    End Select
                End If
            End If
        Else
            Set position = base.ShootRandomPosition(this.Random, enemyGrid)
        End If
        iterations = iterations + 1
    Loop Until base.IsLegalPosition(enemyGrid, position)
    Set lastPosition = position
    Set ScanEdges = position
End Function

Private Function AvoidAdjacentHitPosition(ByVal enemyGrid As PlayerGrid, ByVal position As IGridCoord) As Boolean
'prefer to avoid shooting in positions adjacent to known ship positions; return false to deny position
    Dim allowPosition As Boolean
    allowPosition = True
    If Not enemyGrid.HasDamagedShip(position) Then
    'if there's a damaged ship (hit but not sunken) at that position, let it be allowed.
        
        If this.Random.Maybe(AlmostNever) Then
            allowPosition = True
        
        Else
            allowPosition = True
            
            Dim leftSide As IGridCoord
            Set leftSide = position.Offset(-1)
            If leftSide.X >= 1 Then
                allowPosition = allowPosition And enemyGrid.State(leftSide) <> PreviousHit
            End If
            
            Dim rightSide As IGridCoord
            Set rightSide = position.Offset(1)
            If rightSide.X <= PlayerGrid.Size Then
                allowPosition = allowPosition And enemyGrid.State(rightSide) <> PreviousHit
            End If
            
            Dim topSide As IGridCoord
            Set topSide = position.Offset(0, -1)
            If topSide.Y >= 1 Then
                allowPosition = allowPosition And enemyGrid.State(topSide) <> PreviousHit
            End If
            
            Dim bottomSide As IGridCoord
            Set bottomSide = position.Offset(0, 1)
            If bottomSide.Y <= PlayerGrid.Size Then
                allowPosition = allowPosition And enemyGrid.State(bottomSide) <> PreviousHit
            End If
        End If
    End If
    AvoidAdjacentHitPosition = allowPosition
End Function

