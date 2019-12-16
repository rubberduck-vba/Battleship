VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StandardGameController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "A game controller implementing the classic rules."
'@Folder("Battleship")
'@ModuleDescription("A game controller implementing the classic rules.")
'@PredeclaredId
Option Explicit

Private Type TController
    PlayerFactory As IPlayerFactory
    Random As IRandomizer
    
    View As IGridViewCommands
    
    Player1 As IPlayer
    Player2 As IPlayer
    CurrentPlayer As IPlayer
    CurrentTarget As IPlayer
    CurrentShip As IShip
End Type
Private WithEvents viewAdapter As GridViewAdapter
Attribute viewAdapter.VB_VarHelpID = -1
Private this As TController

Implements IGameController

Public Function Create(ByVal adapter As GridViewAdapter, ByVal randomizer As IRandomizer, ByVal players As IPlayerFactory) As IGameController
    With New StandardGameController
        Set .View = adapter
        Set .PlayerFactory = players
        Set .Random = randomizer
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IGameController
    Set Self = Me
End Property

Public Property Get PlayerFactory() As IPlayerFactory
    Set PlayerFactory = this.PlayerFactory
End Property

Public Property Set PlayerFactory(ByVal value As IPlayerFactory)
    Set this.PlayerFactory = value
End Property

Public Property Get View() As IGridViewCommands
    Set View = this.View
End Property

Public Property Set View(ByVal adapter As IGridViewCommands)
    Set this.View = adapter
    Set viewAdapter = adapter
End Property

Public Property Get Random() As IRandomizer
    Set Random = this.Random
End Property

Public Property Set Random(ByVal value As IRandomizer)
    Set this.Random = value
End Property

Private Sub IGameController_NewGame()
    this.View.OnNewGame
End Sub

Private Sub viewAdapter_OnCreatePlayer(ByVal gridId As PlayGridId, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
    If gridId = PlayerGrid1 And Not this.Player1 Is Nothing Then Exit Sub
    If gridId = PlayerGrid2 And Not this.Player2 Is Nothing Then Exit Sub
    
    Dim player As IPlayer
    Set player = this.PlayerFactory.Create(gridId, difficulty)
    
    If gridId = PlayerGrid1 Then
        Set this.Player1 = player
    ElseIf gridId = PlayerGrid2 Then
        Set this.Player2 = player
    End If
    
    If Not this.Player1 Is Nothing And Not this.Player2 Is Nothing Then
        Set this.CurrentPlayer = this.Player1
        Set this.CurrentTarget = this.Player2
        If this.Player1.PlayerType = ComputerControlled Then EndCurrentPlayerTurn
        OnShipPositionStart
    End If
    
End Sub

Private Sub OnShipPositionStart()
    
    Dim kinds As Variant
    kinds = Ship.ShipKinds
    Set this.CurrentShip = Ship.Create(kinds(0), Horizontal, GridCoord.Create(1, 1))
    
    If this.Player1.PlayerType = HumanControlled Then
        View.OnBeginShipPosition this.CurrentShip, this.Player1
    ElseIf this.Player2.PlayerType = HumanControlled Then
        View.OnBeginShipPosition this.CurrentShip, this.Player2
    Else
        'AI vs AI
        Dim i As Long
        For i = LBound(kinds) To UBound(kinds)
            Set this.CurrentShip = Ship.Create(kinds(i), Horizontal, GridCoord.Create(1, 1))
            this.Player1.PlaceShip this.CurrentShip
            this.Player2.PlaceShip this.CurrentShip
        Next
        Set this.CurrentPlayer = this.Player1
        Set this.CurrentTarget = this.Player2
        PlayAIvsAI
    End If
    
End Sub

Private Sub viewAdapter_OnGameOver(ByVal winner As IPlayer)
    Teardown
End Sub

Private Sub viewAdapter_OnPreviewCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    On Error Resume Next
    Set this.CurrentShip = Ship.Create(this.CurrentShip.ShipKind, this.CurrentShip.Orientation, position)
    On Error GoTo 0
    If gridId = 1 Then
        View.OnPreviewShipPosition this.Player1, this.CurrentShip
    Else
        View.OnPreviewShipPosition this.Player2, this.CurrentShip
    End If
End Sub

Private Sub viewAdapter_OnRotateCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    On Error Resume Next
    Set this.CurrentShip = Ship.Create(this.CurrentShip.ShipKind, IIf(this.CurrentShip.Orientation = Horizontal, Vertical, Horizontal), position)
    On Error GoTo 0
    If gridId = 1 Then
        View.OnPreviewShipPosition this.Player1, this.CurrentShip
    Else
        View.OnPreviewShipPosition this.Player2, this.CurrentShip
    End If
End Sub

Private Sub viewAdapter_OnConfirmCurrentShipPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Debug.Assert gridId = this.CurrentPlayer.PlayGrid.gridId
    
    If this.CurrentPlayer.PlayGrid.CanAddShip(this.CurrentShip.GridPosition, this.CurrentShip.Orientation, this.CurrentShip.Size) Then
        this.CurrentPlayer.PlayGrid.AddShip this.CurrentShip
        View.OnConfirmShipPosition this.CurrentPlayer, this.CurrentShip
    Else
        View.OnInvalidShipPosition
        Exit Sub
    End If
    
    If this.Player1.PlayerType = ComputerControlled Then
        this.Player1.PlaceShip this.CurrentShip
    ElseIf this.Player2.PlayerType = ComputerControlled Then
        this.Player2.PlaceShip this.CurrentShip
    End If
    
    Debug.Assert this.Player1.PlayGrid.shipCount = this.Player2.PlayGrid.shipCount
    
    Dim ships As Long
    ships = this.CurrentPlayer.PlayGrid.shipCount
    
    If ships < PlayerGrid.ShipsPerGrid Then
        Dim kind As ShipType
        kind = Ship.ShipKinds(ships)
        Set this.CurrentShip = Ship.Create(kind, Horizontal, GridCoord.Create(1, 1))
        View.OnBeginShipPosition this.CurrentShip, this.CurrentPlayer
    Else
        Set this.CurrentShip = Nothing
        View.OnBeginAttack gridId
    End If
End Sub

Private Sub viewAdapter_OnPlayerReady()
    If this.CurrentPlayer Is this.Player2 Then EndCurrentPlayerTurn
    If this.Player1.PlayerType = ComputerControlled Then
        PlayAI
        EndCurrentPlayerTurn
    End If
End Sub

Private Sub viewAdapter_OnAttackPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    If gridId = this.CurrentPlayer.PlayGrid.gridId Then Exit Sub
    On Error GoTo CleanFail
    
    Play gridId, position
    EndCurrentPlayerTurn
    
    If this.CurrentPlayer.PlayerType = ComputerControlled And Not this.CurrentPlayer.PlayGrid.IsAllSunken Then
        PlayAI
        EndCurrentPlayerTurn
    End If
    
    Exit Sub
CleanFail:
    With Err
        If .Number = PlayerGridErrors.KnownGridStateError Then
            View.OnKnownPositionAttack
        End If
    End With
End Sub

Private Sub PlayAIvsAI()
    Do Until this.CurrentTarget.PlayGrid.IsAllSunken Or this.CurrentPlayer.PlayGrid.IsAllSunken
        PlayAI
        EndCurrentPlayerTurn
    Loop
    Teardown
End Sub

Private Sub PlayAI()
    Debug.Assert this.CurrentPlayer.PlayerType <> HumanControlled
    View.OnBeginWaitForComputerPlayer
    Play this.CurrentTarget.PlayGrid.gridId, this.CurrentPlayer.Play(this.CurrentTarget.PlayGrid)
    View.OnEndWaitForComputerPlayer
End Sub

Private Sub Play(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
    Dim result As AttackResult, hitShip As IShip
    result = this.CurrentTarget.PlayGrid.TryHit(position, hitShip)
    
    View.OnRefreshGrid this.CurrentTarget.PlayGrid
    View.OnSelectPosition gridId, position
    
    Dim showAIfleet As Boolean
    showAIfleet = (this.Player1.PlayerType = ComputerControlled And this.Player2.PlayerType = ComputerControlled) Or result = Sunk
    Select Case result
    
        Case AttackResult.Miss
            View.OnMiss gridId
            
        Case AttackResult.Hit
            View.OnUpdateFleetStatus this.CurrentTarget, hitShip, showAIfleet
            View.OnHit gridId
            
        Case AttackResult.Sunk
            View.OnUpdateFleetStatus this.CurrentTarget, hitShip, showAIfleet
            If this.CurrentTarget.PlayGrid.IsAllSunken Then
                View.OnGameOver this.CurrentPlayer.PlayGrid.gridId
            Else
                View.OnSink gridId
            End If
            
    End Select
End Sub

Private Sub EndCurrentPlayerTurn()
    If this.CurrentPlayer Is this.Player1 Then
        Set this.CurrentPlayer = this.Player2
        Set this.CurrentTarget = this.Player1
    Else
        Set this.CurrentPlayer = this.Player1
        Set this.CurrentTarget = this.Player2
    End If
End Sub

Private Sub Teardown()
    Set View = Nothing
    Set viewAdapter = Nothing
End Sub
