Attribute VB_Name = "GridCoordTests"
Attribute VB_Description = "Tests covering the Battleship.GridCoord class."
'@Folder("Tests")
'@ModuleDescription("Tests covering the Battleship.GridCoord class.")
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

'@TestMethod("GridCoord")
Public Sub CreatesAtSpecifiedXCoordinate()
    Const expectedX As Long = 42
    Const expectedY As Long = 74
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(expectedX, expectedY)
    
    Assert.AreEqual expectedX, sut.X, "X coordinate mismatched."
    Assert.AreEqual expectedY, sut.Y, "Y coordinate mismatched."
End Sub

'@TestMethod("GridCoord")
Public Sub DefaultIsZeroAndZero()
    Const expectedX As Long = 0
    Const expectedY As Long = 0
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Assert.AreEqual expectedX, sut.X, "X coordinate mismatched."
    Assert.AreEqual expectedY, sut.Y, "Y coordinate mismatched."
End Sub

'@TestMethod("GridCoord")
Public Sub OffsetAddsX()
    Const xOffset As Long = 1
    Const yOffset As Long = 0

    Dim initial As IGridCoord
    Set initial = GridCoord.Default
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Dim actual As IGridCoord
    Set actual = sut.Offset(xOffset, yOffset)
    
    Assert.AreEqual initial.X + xOffset, actual.X
End Sub

'@TestMethod("GridCoord")
Public Sub OffsetAddsY()
    Const xOffset As Long = 0
    Const yOffset As Long = 1

    Dim initial As IGridCoord
    Set initial = GridCoord.Default
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Dim actual As IGridCoord
    Set actual = sut.Offset(xOffset, yOffset)
    
    Assert.AreEqual initial.Y + yOffset, actual.Y
End Sub

'@TestMethod("GridCoord")
Public Sub FromToString_RoundTrips()
    Dim initial As IGridCoord
    Set initial = GridCoord.Default
    
    Dim asString As String
    asString = initial.ToString
    
    Dim sut As IGridCoord
    Set sut = GridCoord.FromString(asString)
            
    Assert.AreEqual initial.X, sut.X, "X coordinate mismatched."
    Assert.AreEqual initial.Y, sut.Y, "Y coordinate mismatched."
End Sub

'@TestMethod("GridCoord")
Public Sub ToStringFormat_NoSpaceCommaSeparatedInParentheses()
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Dim expected As String
    expected = "(" & sut.X & "," & sut.Y & ")"
    
    Dim actual As String
    actual = sut.ToString
    
    Assert.AreEqual expected, actual
End Sub

'@TestMethod("GridCoord")
Public Sub EqualsReturnsTrueForMatchingCoords()
    Dim other As IGridCoord
    Set other = GridCoord.Default
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Assert.IsTrue sut.Equals(other)
End Sub

'@TestMethod("GridCoord")
Public Sub EqualsReturnsFalseForMismatchingCoords()
    Dim other As IGridCoord
    Set other = GridCoord.Default.Offset(1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Default
    
    Assert.IsFalse sut.Equals(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenOneLeftAndSameY_IsAdjacentReturnsTrue()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(2, 1)
    
    Assert.IsTrue sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenTwoLeftAndSameY_IsAdjacentReturnsFalse()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(3, 1)
    
    Assert.IsFalse sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenOneRightAndSameY_IsAdjacentReturnsTrue()
    Dim other As IGridCoord
    Set other = GridCoord.Create(3, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(2, 1)
    
    Assert.IsTrue sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenTwoRightAndSameY_IsAdjacentReturnsFalse()
    Dim other As IGridCoord
    Set other = GridCoord.Create(5, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(3, 1)
    
    Assert.IsFalse sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenOneDownAndSameX_IsAdjacentReturnsTrue()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 2)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(1, 1)
    
    Assert.IsTrue sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenTwoDownAndSameX_IsAdjacentReturnsFalse()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 3)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(1, 1)
    
    Assert.IsFalse sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenOneUpAndSameX_IsAdjacentReturnsTrue()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(1, 2)
    
    Assert.IsTrue sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenTwoUpAndSameX_IsAdjacentReturnsFalse()
    Dim other As IGridCoord
    Set other = GridCoord.Create(1, 1)
    
    Dim sut As IGridCoord
    Set sut = GridCoord.Create(1, 3)
    
    Assert.IsFalse sut.IsAdjacent(other)
End Sub

'@TestMethod("GridCoord")
Public Sub GivenInvalidString_FromStringThrows()
    Const ExpectedError As Long = 5
    On Error GoTo TestFail
    
    Dim sut As IGridCoord
    Set sut = GridCoord.FromString("invalid string")
    
Assert:
    Assert.Fail "Expected error was not raised."

TestExit:
    Exit Sub
TestFail:
    If Err.Number = ExpectedError Then
        Resume TestExit
    Else
        Resume Assert
    End If
End Sub


