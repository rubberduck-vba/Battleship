VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RandomShotStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship.AI")
'@Description("A strategy that deploys ships in the first legal random position found, and shoots at random until all enemy ships are found.")
Option Explicit
Implements IGameStrategy

Private Type TStrategy
    Random As IRandomizer
End Type

Private base As GameStrategyBase
Private this As TStrategy

Public Function Create(ByVal randomizer As IRandomizer) As IGameStrategy
    With New RandomShotStrategy
        Set .Random = randomizer
        Set Create = .Self
    End With
End Function

Public Property Get Self() As RandomShotStrategy
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
    
    Dim direction As ShipOrientation
    Dim position As IGridCoord
    Set position = base.PlaceShip(Random, grid, currentShip, direction)
    
    grid.AddShip Ship.Create(currentShip.ShipKind, direction, position)
    If grid.shipCount = PlayerGrid.ShipsPerGrid Then grid.Scramble
    
End Sub

Private Function IGameStrategy_Play(ByVal enemyGrid As PlayerGrid) As IGridCoord
    Dim position As IGridCoord
    Do
        If EnemyShipsToFind(enemyGrid) > 0 Then
            Set position = base.ShootRandomPosition(Random, enemyGrid)
        Else
            Set position = base.DestroyTarget(Random, enemyGrid, enemyGrid.FindHitArea)
        End If
    Loop Until base.IsLegalPosition(enemyGrid, position)
    Set IGameStrategy_Play = position
End Function

Private Function EnemyShipsToFind(ByVal enemyGrid As PlayerGrid) As Byte
    Dim enemyFleet As Collection
    Set enemyFleet = enemyGrid.Fleet
    
    Dim result As Byte
    
    Dim currentShip As IShip
    For Each currentShip In enemyFleet
        If currentShip.HitAreas.Count = 0 Then result = result + 1
    Next
    
    EnemyShipsToFind = result
End Function



