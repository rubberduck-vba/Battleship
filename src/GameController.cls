VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Battleship")
Option Explicit

Private player1 As IPlayer
Private player2 As IPlayer

Private currentPlayer As IPlayer
Private currentTarget As IPlayer
Private currentShip As IShip

Private RNG As IRandomizer

Private view As IGridViewCommands
Private WithEvents viewAdapter As GridViewAdapter
Attribute viewAdapter.VB_VarHelpID = -1

Public Sub NewGame(ByVal adapter As GridViewAdapter, ByVal randomizer As IRandomizer)
    Set viewAdapter = adapter
    Set RNG = randomizer
    Set view = adapter
    view.OnNewGame
End Sub

Private Sub viewAdapter_OnCreatePlayer(ByVal gridId As Byte, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
    If gridId = 1 And Not player1 Is Nothing Then Exit Sub
    If gridId = 2 And Not player2 Is Nothing Then Exit Sub
    
    Dim player As IPlayer
    Select Case pt
        
        Case HumanControlled
            Set player = HumanPlayer.Create(gridId)
            
        Case ComputerControlled
            Select Case difficulty
                Case AIDifficulty.RandomAI
                    Set player = AIPlayer.Create(gridId, RandomShotStrategy.Create(RNG))
                Case AIDifficulty.FairplayAI
                    Set player = AIPlayer.Create(gridId, FairPlayStrategy.Create(RNG))
                Case AIDifficulty.MercilessAI
                    Set player = AIPlayer.Create(gridId, MercilessStrategy.Create(RNG))
            End Select
    
    End Select
    
    If gridId = 1 Then
        Set player1 = player
    ElseIf gridId = 2 Then
        Set player2 = player
    End If
    
    If Not player1 Is Nothing And Not player2 Is Nothing Then
        Set currentPlayer = player1
        Set currentTarget = player2
        If player1.PlayerType = ComputerControlled Then EndCurrentPlayerTurn
        OnShipPositionStart
    End If
    
End Sub

Private Sub OnShipPositionStart()
    
    Dim kinds As Variant
    kinds = Ship.ShipKinds
    Set currentShip = Ship.Create(kinds(0), Horizontal, GridCoord.Create(1, 1))
    
    If player1.PlayerType = HumanControlled Then
        view.OnBeginShipPosition currentShip, player1
    ElseIf player2.PlayerType = HumanControlled Then
        view.OnBeginShipPosition currentShip, player2
    Else
        'AI vs AI
        Dim i As Long
        For i = LBound(kinds) To UBound(kinds)
            Set currentShip = Ship.Create(kinds(i), Horizontal, GridCoord.Create(1, 1))
            player1.PlaceShip currentShip
            player2.PlaceShip currentShip
        Next
        Set currentPlayer = player1
        Set currentTarget = player2
        PlayAIvsAI
    End If
    
End Sub

Private Sub viewAdapter_OnGameOver(ByVal winner As IPlayer)
    Teardown
End Sub

Private Sub viewAdapter_OnPreviewCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    On Error Resume Next
    Set currentShip = Ship.Create(currentShip.ShipKind, currentShip.Orientation, position)
    On Error GoTo 0
    If gridId = 1 Then
        view.OnPreviewShipPosition player1, currentShip
    Else
        view.OnPreviewShipPosition player2, currentShip
    End If
End Sub

Private Sub viewAdapter_OnRotateCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    On Error Resume Next
    Set currentShip = Ship.Create(currentShip.ShipKind, IIf(currentShip.Orientation = Horizontal, Vertical, Horizontal), position)
    On Error GoTo 0
    If gridId = 1 Then
        view.OnPreviewShipPosition player1, currentShip
    Else
        view.OnPreviewShipPosition player2, currentShip
    End If
End Sub

Private Sub viewAdapter_OnConfirmCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    Debug.Assert gridId = currentPlayer.PlayGrid.gridId
    
    If currentPlayer.PlayGrid.CanAddShip(currentShip.GridPosition, currentShip.Orientation, currentShip.Size) Then
        currentPlayer.PlayGrid.AddShip currentShip
        view.OnConfirmShipPosition currentPlayer, currentShip
    Else
        view.OnInvalidShipPosition
        Exit Sub
    End If
    
    If player1.PlayerType = ComputerControlled Then
        player1.PlaceShip currentShip
    ElseIf player2.PlayerType = ComputerControlled Then
        player2.PlaceShip currentShip
    End If
    
    Debug.Assert player1.PlayGrid.shipCount = player2.PlayGrid.shipCount
    
    Dim ships As Long
    ships = currentPlayer.PlayGrid.shipCount
    
    If ships < PlayerGrid.ShipsPerGrid Then
        Dim kind As ShipType
        kind = Ship.ShipKinds(ships)
        Set currentShip = Ship.Create(kind, Horizontal, GridCoord.Create(1, 1))
        view.OnBeginShipPosition currentShip, currentPlayer
    Else
        Set currentShip = Nothing
        view.OnBeginAttack gridId
    End If
End Sub

Private Sub viewAdapter_OnPlayerReady()
    If currentPlayer Is player2 Then EndCurrentPlayerTurn
    If player1.PlayerType = ComputerControlled Then
        PlayAI
        EndCurrentPlayerTurn
    End If
End Sub

Private Sub viewAdapter_OnAttackPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    If gridId = currentPlayer.PlayGrid.gridId Then Exit Sub
    On Error GoTo CleanFail
    
    Play gridId, position
    EndCurrentPlayerTurn
    
    If currentPlayer.PlayerType = ComputerControlled And Not currentPlayer.PlayGrid.IsAllSunken Then
        PlayAI
        EndCurrentPlayerTurn
    End If
    
    Exit Sub
CleanFail:
    With Err
        If .Number = PlayerGridErrors.KnownGridStateError Then
            view.OnKnownPositionAttack
        End If
    End With
End Sub

Private Sub PlayAIvsAI()
    Do Until currentTarget.PlayGrid.IsAllSunken Or currentPlayer.PlayGrid.IsAllSunken
        PlayAI
        EndCurrentPlayerTurn
    Loop
    Teardown
End Sub

Private Sub PlayAI()
    Debug.Assert currentPlayer.PlayerType <> HumanControlled
    view.OnBeginWaitForComputerPlayer
    Play currentTarget.PlayGrid.gridId, currentPlayer.Play(currentTarget.PlayGrid)
    view.OnEndWaitForComputerPlayer
End Sub

Private Sub Play(ByVal gridId As Byte, ByVal position As IGridCoord)
    Dim result As AttackResult, hitShip As IShip
    result = currentTarget.PlayGrid.TryHit(position, hitShip)
    
    view.OnRefreshGrid currentTarget.PlayGrid
    view.OnSelectPosition gridId, position
    
    Dim showAIfleet As Boolean
    showAIfleet = (player1.PlayerType = ComputerControlled And player2.PlayerType = ComputerControlled) Or result = Sunk
    Select Case result
    
        Case AttackResult.Miss
            view.OnMiss gridId
            
        Case AttackResult.Hit
            view.OnUpdateFleetStatus currentTarget, hitShip, showAIfleet
            view.OnHit gridId
            
        Case AttackResult.Sunk
            view.OnUpdateFleetStatus currentTarget, hitShip, showAIfleet
            If currentTarget.PlayGrid.IsAllSunken Then
                view.OnGameOver currentPlayer.PlayGrid.gridId
            Else
                view.OnSink gridId
            End If
            
    End Select
End Sub

Private Sub EndCurrentPlayerTurn()
    If currentPlayer Is player1 Then
        Set currentPlayer = player2
        Set currentTarget = player1
    Else
        Set currentPlayer = player1
        Set currentTarget = player2
    End If
End Sub

Private Sub Teardown()
    Set view = Nothing
    Set viewAdapter = Nothing
End Sub
