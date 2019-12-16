Attribute VB_Name = "Resources"
Attribute VB_Description = "A module for accessing localized string resources."
'@Folder("Battleship.Resources")
'@ModuleDescription("A module for accessing localized string resources.")
Option Explicit
Option Private Module

Public Const DefaultCulture As String = "en-US"

Public Function GetString(ByVal key As String, Optional ByVal cultureKey As String) As String
    If cultureKey = vbNullString Then cultureKey = Resources.DefaultCulture
    GetString = ResourcesSheet.Resource(key, cultureKey)
End Function

