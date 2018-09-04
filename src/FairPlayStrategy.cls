VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FairPlayStrategy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("Battleship.AI")
Option Explicit
Implements IGameStrategy

Private Type TStrategy
    Random As IRandomizer
End Type

Private base As GameStrategyBase
Private this As TStrategy

Public Function Create(ByVal randomizer As IRandomizer) As IGameStrategy
    With New FairPlayStrategy
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

Public Property Set Random(ByVal Value As IRandomizer)
    Set this.Random = Value
End Property

Private Sub Class_Initialize()
    Set base = New GameStrategyBase
End Sub

Private Sub IGameStrategy_PlaceShip(ByVal grid As PlayerGrid, ByVal currentShip As IShip)
    Do
        Dim direction As ShipOrientation
        Dim position As IGridCoord
        Set position = base.PlaceShip(Random, grid, currentShip, direction)
        
    Loop Until Not grid.HasAdjacentShip(position, direction, currentShip.Size)
    
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
            Set position = base.ShootRandomPosition(Random, enemyGrid)
        End If
    Loop Until base.IsLegalPosition(enemyGrid, position)
    Set IGameStrategy_Play = position
End Function



