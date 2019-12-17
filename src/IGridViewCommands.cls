VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGridViewCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Interface for commands sent from the controller to the view."
'@Exposed
'@Folder("Battleship.View.Abstract")
'@ModuleDescription("Interface for commands sent from the controller to the view.")
'@Interface
Option Explicit

'@Description("Gets/sets a weak refererence to the view events.")
Public Property Get Events() As IGridViewEvents
Attribute Events.VB_Description = "Gets/sets a weak refererence to the view events."
End Property

Public Property Set Events(ByVal value As IGridViewEvents)
End Property

'@Description("Instructs the view to report to a miss in the specified grid.")
Public Sub OnMiss(ByVal gridId As PlayGridId)
Attribute OnMiss.VB_Description = "Instructs the view to report to a miss in the specified grid."
End Sub

'@Description("Instructs the view to report a hit in the specified grid.")
Public Sub OnHit(ByVal gridId As PlayGridId)
Attribute OnHit.VB_Description = "Instructs the view to report a hit in the specified grid."
End Sub

'@Description("Instructs the view to report a sunken ship in the specified grid.")
Public Sub OnSink(ByVal gridId As PlayGridId)
Attribute OnSink.VB_Description = "Instructs the view to report a sunken ship in the specified grid."
End Sub

'@Description("Instructs the view to update the specified player's fleet status, for the specified ship.")
Public Sub OnUpdateFleetStatus(ByVal player As IPlayer, ByVal hitShip As IShip, Optional ByVal showAIStatus As Boolean = False)
Attribute OnUpdateFleetStatus.VB_Description = "Instructs the view to update the specified player's fleet status, for the specified ship."
End Sub

'@Description("Instructs the view to select the specified position in the specified grid.")
Public Sub OnSelectPosition(ByVal gridId As PlayGridId, ByVal position As IGridCoord)
Attribute OnSelectPosition.VB_Description = "Instructs the view to select the specified position in the specified grid."
End Sub

'@Description("Instructs the view to lock the specified grid, preventing user interaction.")
Public Sub OnLockGrid(ByVal gridId As PlayGridId)
Attribute OnLockGrid.VB_Description = "Instructs the view to lock the specified grid, preventing user interaction."
End Sub

'@Description("Instructs the view to begin a new game.")
Public Sub OnNewGame()
Attribute OnNewGame.VB_Description = "Instructs the view to begin a new game."
End Sub

'@Description("Instructs the view to report the end of the game.")
Public Sub OnGameOver(ByVal winningGrid As PlayGridId)
Attribute OnGameOver.VB_Description = "Instructs the view to report the end of the game."
End Sub

'@Description("Instructs the view to begin positioning the specified ship.")
Public Sub OnBeginShipPosition(ByVal CurrentShip As IShip, ByVal player As IPlayer)
Attribute OnBeginShipPosition.VB_Description = "Instructs the view to begin positioning the specified ship."
End Sub

'@Description("Instructs the view to confirm the position of the specified ship.")
Public Sub OnConfirmShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
Attribute OnConfirmShipPosition.VB_Description = "Instructs the view to confirm the position of the specified ship."
End Sub

'@Description("Instructs the view to preview the position of the specified ship.")
Public Sub OnPreviewShipPosition(ByVal player As IPlayer, ByVal newShip As IShip)
Attribute OnPreviewShipPosition.VB_Description = "Instructs the view to preview the position of the specified ship."
End Sub

'@Description("Instructs the view to react to an attempt to place the current ship at an invalid position.")
Public Sub OnInvalidShipPosition()
Attribute OnInvalidShipPosition.VB_Description = "Instructs the view to react to an attempt to place the current ship at an invalid position."
End Sub

'@Description("Instructs the view to begin attack phase.")
Public Sub OnBeginAttack(ByVal currentPlayerGridId As PlayGridId)
Attribute OnBeginAttack.VB_Description = "Instructs the view to begin attack phase."
End Sub

'@Description("Instructs the view to react to an attack attempt on a known-state position.")
Public Sub OnKnownPositionAttack()
Attribute OnKnownPositionAttack.VB_Description = "Instructs the view to react to an attack attempt on a known-state position."
End Sub

'@Description("Instructs the view to redraw the specified grid.")
Public Sub OnRefreshGrid(ByVal grid As PlayerGrid)
Attribute OnRefreshGrid.VB_Description = "Instructs the view to redraw the specified grid."
End Sub

Public Sub OnBeginWaitForComputerPlayer()
End Sub

Public Sub OnEndWaitForComputerPlayer()
End Sub
