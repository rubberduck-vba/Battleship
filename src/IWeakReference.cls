VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IWeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Describes an object that holds the address of a pointer to another object."
'@Exposed
'@Folder("Win32.Abstract")
'@ModuleDescription("Describes an object that holds the address of a pointer to another object.")
'@Interface
Option Explicit

'@Description("Gets the object at the held pointer address.")
Public Property Get Object() As Object
Attribute Object.VB_Description = "Gets the object at the held pointer address."
End Property
