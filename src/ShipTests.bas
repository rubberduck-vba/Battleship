Attribute VB_Name = "ShipTests"
Attribute VB_Description = "Tests covering the Battleship.Ship class."
'@Folder("Tests")
'@Description("Tests covering the Battleship.Ship class.")
'@TestModule
Option Explicit
Option Private Module

Private Assert As Object 'Rubberduck.AssertClass
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

'@TestInitialize
Public Sub TestInitialize()
End Sub

'@TestCleanup
Public Sub TestCleanup()
End Sub

'@TestMethod
Public Sub CreatesShipOfSpecifiedType()
    Dim expected As ShipType
    expected = ShipType.Battleship
    
    Dim sut As IShip
    Set sut = Ship.Create(expected, Horizontal, GridCoord.Default)
    
    Dim actual As ShipType
    actual = sut.ShipKind
    
    Assert.AreEqual expected, actual
End Sub

'@TestMethod
Public Sub SuccessfulHit_SetsStateToTrue()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Battleship, Horizontal, position)
    
    If Not position.Equals(GridCoord.Default) Then Assert.Inconclusive
    If obj.State(position.ToString) Then Assert.Inconclusive
    
    Dim sut As IShip
    Set sut = obj
    sut.Hit position
    
    Assert.IsTrue obj.State(position.ToString)
End Sub

'@TestMethod
Public Sub SunkenShip_IsSunken()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Battleship, Horizontal, position)
    
    If Not position.Equals(GridCoord.Default) Then Assert.Inconclusive
    
    Dim sut As IShip
    Set sut = obj
    
    Dim current As Variant
    For Each current In obj.State.Keys
        sut.Hit GridCoord.FromString(current)
    Next
    
    Assert.IsTrue sut.IsSunken
End Sub

'@TestMethod
Public Sub MissedHit_ReturnsFalse()
    Dim position As IGridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Battleship, Horizontal, position)
    
    Dim sut As IShip
    Set sut = obj
    
    Dim target As GridCoord
    Set target = position.Offset( _
        xOffset:=IIf(sut.Orientation = Horizontal, 0, 1), _
        yOffset:=IIf(sut.Orientation = Vertical, 0, 1))
    
    Assert.IsFalse sut.Hit(target)
End Sub

'@TestMethod
Public Sub SuccessfulHit_ReturnsTrue()
    Dim position As GridCoord
    Set position = GridCoord.Default
    
    Dim obj As Ship
    Set obj = Ship.Create(ShipType.Battleship, Horizontal, position)
    
    Dim sut As IShip
    Set sut = obj
    
    Assert.IsTrue sut.Hit(position)
End Sub

'@TestMethod
Public Sub ShipOrientationInvalidEnumValue_Throws()
    
    Const ExpectedError As Long = 5
    Const badValue = 42
    
    On Error GoTo TestFail
    
    If badValue = ShipOrientation.Horizontal Or _
       badValue = ShipOrientation.Vertical _
    Then
        Assert.Inconclusive
    End If
    
    Dim sut As Ship
    Set sut = Ship.Create(ShipType.Battleship, badValue, GridCoord.Default)
    
Assert:
    Assert.Fail "Expected error was not raised. A Ship instance could be created with an unknown orientation."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub

'@TestMethod
Public Sub IntersectingShip_IntersectReturnsGridCoord()
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Battleship, Horizontal, GridCoord.Default)
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Battleship, Vertical, GridCoord.Default)
    
    Assert.IsNotNothing Ship1.Intersects(Ship2.Size, Ship2.Orientation, Ship2.GridPosition)
End Sub

'@TestMethod
Public Sub NonIntersectingShip_IntersectReturnsNothing()
    Dim Ship1 As IShip
    Set Ship1 = Ship.Create(ShipType.Battleship, Horizontal, GridCoord.Default)
    
    Dim Ship2 As IShip
    Set Ship2 = Ship.Create(ShipType.Battleship, Horizontal, GridCoord.Default.Offset(yOffset:=1))
    
    Assert.IsNothing Ship1.Intersects(Ship2.Size, Ship2.Orientation, Ship2.GridPosition)
End Sub

