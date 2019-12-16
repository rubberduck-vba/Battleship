VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IPlayerFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object responsible for creating players."
'@Folder("Battleship.Model.Abstract")
'@ModuleDescription("Describes an object responsible for creating players.")
'@Interface
Option Explicit

'@Description("Creates a HumanPlayer unless an AIDifficulty argument is specified.")
Public Function Create(ByVal gridId As PlayGridId, Optional ByVal difficulty As AIDifficulty = AIDifficulty.Unspecified) As IPlayer
Attribute Create.VB_Description = "Creates a HumanPlayer unless an AIDifficulty argument is specified."
End Function
