Attribute VB_Name = "GameStrategyBaseTests"
'@Folder("Tests")
'@Description("Tests covering the Battleship.GameStrategyBase class.")
'@TestModule
Option Explicit
Option Private Module

Private Assert As Rubberduck.AssertClass
'Private Fakes As Rubberduck.FakesProvider

'@ModuleInitialize
Public Sub ModuleInitialize()
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

'@TestMethod
Public Sub VerifyShipFits_TrueGivenEmptyGrid()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsTrue sut.VerifyShipFits(grid, GridCoord.Create(1, 1), 5)
End Sub

'@TestMethod
Public Sub VerifyShipFits_FalseGivenPreviousMissesAroundTarget()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missB1 As IGridCoord
    Set missB1 = GridCoord.Create(2, 1)
    
    Dim missA2 As IGridCoord
    Set missA2 = GridCoord.Create(1, 2)
    
    grid.TryHit missB1
    grid.TryHit missA2
    
    If grid.State(missB1) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missB1.ToA1String & " must be 'MISS'."
    If grid.State(missA2) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missA2.ToA1String & " must be 'MISS'."
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsFalse sut.VerifyShipFits(grid, position, 5)
    
End Sub

'@TestMethod
Public Sub VerifyShipFits_TrueGivenEnoughHorizontalUnknownState()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missA2 As IGridCoord
    Set missA2 = GridCoord.Create(1, 2)
    
    grid.TryHit missA2
    If grid.State(missA2) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missA2.ToA1String & " must be 'MISS'."
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsTrue sut.VerifyShipFits(grid, position, 5)
    
End Sub

'@TestMethod
Public Sub VerifyShipFits_TrueGivenEnoughVerticalUnknownState()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missB1 As IGridCoord
    Set missB1 = GridCoord.Create(2, 1)
    
    grid.TryHit missB1
    If grid.State(missB1) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missB1.ToA1String & " must be 'MISS'."
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(1, 1)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsTrue sut.VerifyShipFits(grid, position, 5)
    
End Sub

'@TestMethod
Public Sub VerifyShipFits_FalseGivenHorizontalEdgeOfGrid()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missH2 As IGridCoord
    Set missH2 = GridCoord.Create(8, 2)
    
    Dim missI3 As IGridCoord
    Set missI3 = GridCoord.Create(9, 3)
    
    grid.TryHit missH2
    If grid.State(missH2) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missH2.ToA1String & " must be 'MISS'."
    
    grid.TryHit missI3
    If grid.State(missI3) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missI3.ToA1String & " must be 'MISS'."
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(9, 2)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsFalse sut.VerifyShipFits(grid, position, 5)
    
End Sub

'@TestMethod
Public Sub VerifyShipFits_FalseGivenVerticalEdgeOfGrid()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missB8 As IGridCoord
    Set missB8 = GridCoord.Create(2, 8)
    
    Dim missC9 As IGridCoord
    Set missC9 = GridCoord.Create(3, 9)
    
    grid.TryHit missB8
    If grid.State(missB8) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missB8.ToA1String & " must be 'MISS'."
    
    grid.TryHit missC9
    If grid.State(missC9) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missC9.ToA1String & " must be 'MISS'."
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(2, 9)
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsFalse sut.VerifyShipFits(grid, position, 5)
    
End Sub

'@TestMethod
Public Sub VerifyShipFits_TrueGivenAdjacentHitArea()
    Dim grid As PlayerGrid
    Set grid = PlayerGrid.Create(1)
    
    Dim missJ3 As IGridCoord
    Set missJ3 = GridCoord.Create(10, 3)
    
    Dim missI4 As IGridCoord
    Set missI4 = GridCoord.Create(9, 4)
    
    Dim hitJ5 As IGridCoord
    Set hitJ5 = GridCoord.Create(10, 5)
    
    Dim position As IGridCoord
    Set position = GridCoord.Create(10, 4)
    
    Dim target As IShip
    Set target = Ship.Create(Carrier, Vertical, position)
    
    grid.AddShip target
    grid.Scramble ' make ShipPosition states Unknown
    
    grid.TryHit missJ3
    If grid.State(missJ3) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missJ3.ToA1String & " must be 'MISS'."
    
    grid.TryHit missI4
    If grid.State(missI4) <> PreviousMiss Then Assert.Inconclusive "Grid state at " & missI4.ToA1String & " must be 'MISS'."
    
    grid.TryHit hitJ5
    If grid.State(hitJ5) <> PreviousHit Then Assert.Inconclusive "Grid state at " & hitJ5.ToA1String & " must be 'HIT'."
    
    Dim sut As GameStrategyBase
    Set sut = New GameStrategyBase
    
    Assert.IsTrue sut.VerifyShipFits(grid, position, 5)
    
End Sub
