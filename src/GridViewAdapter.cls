VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GridViewAdapter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Folder("Battleship.View")
Option Explicit
Implements IGridViewCommands
Implements IGridViewEvents

Public Enum ViewMode
    NewGame
    MessageShown
    FleetPosition
    player1
    player2
    GameOver
End Enum

Public Event OnCreatePlayer(ByVal gridId As Byte, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)

Public Event OnPreviewCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
Public Event OnRotateCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
Public Event OnConfirmCurrentShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)

Public Event OnPlayerReady()

Public Event OnAttackPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
Public Event OnHit(ByVal gridId As Byte, ByVal position As IGridCoord, ByVal hitShip As IShip)
Public Event OnMiss(ByVal gridId As Byte, ByVal position As IGridCoord)

Public Event OnGameOver(ByVal winner As IPlayer)

Private Type TAdapter
    ShipsToPosition As Byte
    GridView As IGridViewCommands
End Type
Private this As TAdapter

Private Sub Class_Initialize()
    this.ShipsToPosition = PlayerGrid.ShipsPerGrid
End Sub

Public Function Create(ByVal view As IGridViewCommands) As GridViewAdapter
    With New GridViewAdapter
        Set .GridView = view
        Set view.Events = .Self
        Set Create = .Self
    End With
End Function

Public Property Get Self() As GridViewAdapter
    Set Self = Me
End Property

'@Description("Gets/sets a reference that exposes commands to send to the view.")
Public Property Get GridView() As IGridViewCommands
    Set GridView = this.GridView
End Property

Public Property Set GridView(ByVal Value As IGridViewCommands)
    Set this.GridView = Value
End Property


':IGridViewEvents
':Messages sent from the view
':***************************

Private Sub IGridViewEvents_AttackPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    RaiseEvent OnAttackPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_ConfirmShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    RaiseEvent OnConfirmCurrentShipPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_CreatePlayer(ByVal gridId As Byte, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
    RaiseEvent OnCreatePlayer(gridId, pt, difficulty)
End Sub

Private Sub IGridViewEvents_HumanPlayerReady()
    RaiseEvent OnPlayerReady
End Sub

Private Sub IGridViewEvents_PreviewRotateShip(ByVal gridId As Byte, ByVal position As IGridCoord)
    RaiseEvent OnRotateCurrentShipPosition(gridId, position)
End Sub

Private Sub IGridViewEvents_PreviewShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    RaiseEvent OnPreviewCurrentShipPosition(gridId, position)
End Sub


':IGridViewCommands
':Messages sent from the controller
':*********************************

Private Property Set IGridViewCommands_Events(ByVal Value As IGridViewEvents)
    Err.Raise 5, TypeName(Me), "Invalid use of property"
End Property

Private Property Get IGridViewCommands_Events() As IGridViewEvents
    Set IGridViewCommands_Events = Me
End Property

Private Sub IGridViewCommands_OnBeginAttack(ByVal currentPlayerGridId As Byte)
    this.GridView.OnBeginAttack currentPlayerGridId
End Sub

Private Sub IGridViewCommands_OnBeginShipPosition(ByVal currentShip As IShip, ByVal player As IPlayer)
    this.GridView.OnLockGrid IIf(player.PlayGrid.gridId = 1, 2, 1)
    this.GridView.OnBeginShipPosition currentShip, player
End Sub

Private Sub IGridViewCommands_OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If player.PlayerType = ComputerControlled Then Exit Sub
    this.GridView.OnConfirmShipPosition player, newShip
    this.ShipsToPosition = this.ShipsToPosition - 1
End Sub

Private Sub IGridViewCommands_OnGameOver(ByVal winningGrid As Byte)
    this.GridView.OnGameOver winningGrid
    Set this.GridView.Events = Nothing
End Sub

Private Sub IGridViewCommands_OnHit(ByVal gridId As Byte)
    this.GridView.OnHit gridId
End Sub

Private Sub IGridViewCommands_OnInvalidShipPosition()
    this.GridView.OnInvalidShipPosition
End Sub

Private Sub IGridViewCommands_OnKnownPositionAttack()
    this.GridView.OnKnownPositionAttack
End Sub

Private Sub IGridViewCommands_OnLockGrid(ByVal gridId As Byte)
    this.GridView.OnLockGrid gridId
End Sub

Private Sub IGridViewCommands_OnMiss(ByVal gridId As Byte)
    this.GridView.OnMiss gridId
End Sub

Private Sub IGridViewCommands_OnNewGame()
    this.GridView.OnNewGame
End Sub

Private Sub IGridViewCommands_OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
    If player.PlayerType = ComputerControlled Then Exit Sub
    this.GridView.OnPreviewShipPosition player, newShip
End Sub

Private Sub IGridViewCommands_OnRefreshGrid(ByVal grid As PlayerGrid)
    this.GridView.OnRefreshGrid grid
End Sub

Private Sub IGridViewCommands_OnSelectPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
    this.GridView.OnSelectPosition gridId, position
End Sub

Private Sub IGridViewCommands_OnSink(ByVal gridId As Byte)
    this.GridView.OnSink gridId
End Sub

Private Sub IGridViewCommands_OnUpdateFleetStatus(ByVal player As IPlayer, ByVal hitShip As IShip, Optional ByVal showAIStatus As Boolean = False)
    this.GridView.OnUpdateFleetStatus player, hitShip, showAIStatus
End Sub

Private Sub IGridViewCommands_OnBeginWaitForComputerPlayer()
    this.GridView.OnBeginWaitForComputerPlayer
End Sub

Private Sub IGridViewCommands_OnEndWaitForComputerPlayer()
    this.GridView.OnEndWaitForComputerPlayer
End Sub


