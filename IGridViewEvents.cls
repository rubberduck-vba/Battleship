VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IGridViewEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("Battleship.View")
'@Description("Commands sent from the view to the GridViewAdapter.")
'@Interface
Option Explicit

Public Sub CreatePlayer(ByVal gridId As Byte, ByVal pt As PlayerType, ByVal difficulty As AIDifficulty)
End Sub

Public Sub PreviewRotateShip(ByVal gridId As Byte, ByVal position As IGridCoord)
End Sub

Public Sub PreviewShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
End Sub

Public Sub ConfirmShipPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
End Sub

Public Sub AttackPosition(ByVal gridId As Byte, ByVal position As IGridCoord)
End Sub

Public Sub HumanPlayerReady()
End Sub
