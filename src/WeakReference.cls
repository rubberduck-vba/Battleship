VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WeakReference"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_Description = "Default instance factory; avoid altering default instance state."
'@Folder("Battleship.Win32")
'@Description("Default instance factory; avoid altering default instance state.")
Option Explicit
Implements IWeakReference

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As LongPtr)
#Else
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
#End If

'@Description("Describes the instance-level state.")
Private Type TReference
#If VBA7 Then
    Address As LongPtr
#Else
    Address As Long
#End If
End Type

'@Description("Encapsulates the instance-level state.")
Private this As TReference

'@Description("Default instance factory method.")
Public Function Create(ByVal instance As Object) As IWeakReference
    With New WeakReference
        .Address = ObjPtr(instance)
        Set Create = .Self
    End With
End Function

Public Property Get Self() As IWeakReference
    Set Self = Me
End Property

#If VBA7 Then
Public Property Get Address() As LongPtr
#Else
Public Property Get Address() As Long
#End If
    Address = this.Address
End Property

#If VBA7 Then
Public Property Let Address(ByVal value As LongPtr)
#Else
Public Property Let Address(ByVal value As Long)
#End If
    this.Address = value
End Property

Private Property Get IWeakReference_Object() As Object
' Based on Bruce McKinney's code for getting an Object from the object pointer:
    
#If VBA7 Then
    Dim pointerSize As LongPtr
#Else
    Dim pointerSize As Long
#End If
    
    On Error GoTo CleanFail
    pointerSize = LenB(this.Address)
    
    Dim obj As Object
    CopyMemory obj, this.Address, pointerSize
    
    Set IWeakReference_Object = obj
    CopyMemory obj, 0&, pointerSize
    
CleanExit:
    Exit Property
    
CleanFail:
    Set IWeakReference_Object = Nothing
    Resume CleanExit
End Property
