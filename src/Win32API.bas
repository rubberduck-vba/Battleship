Attribute VB_Name = "Win32API"
Attribute VB_Description = "Win32 utility function imports."
'@Folder("Battleship.Win32")
'@Description("Win32 utility function imports.")
'@IgnoreModule UserMeaningfulName, HungarianNotation; Win32 parameter names are what they are
Option Explicit
Option Private Module

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type
    
Private Const WM_SETREDRAW = &HB&
Private Const WM_USER = &H400
Private Const EM_GETEVENTMASK = (WM_USER + 59)
Private Const EM_SETEVENTMASK = (WM_USER + 69)

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
    
    Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
    Private Declare PtrSafe Function InvalidateRect Lib "user32" (ByVal hwnd As LongPtr, ByVal lpRect As LongPtr, ByVal bErase As LongPtr) As LongPtr
    Private Declare PtrSafe Function UpdateWindow Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function LockWindowUpdate Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As LongPtr, ByVal nIndex As LongPtr, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As LongPtr, ByVal hwnd As LongPtr, ByVal Msg As LongPtr, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

    Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As String,ByVal lpszWindow As String) As Long
    Private Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, ByRef lpRect As Long, ByVal bErase As Long) As Long
    Private Declare Function UpdateWindow Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function LockWindowUpdate Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
#End If

Public Property Get GUIDSIZE() As Long
    Dim Value As GUID
    GUIDSIZE = LenB(Value)
End Property

Public Sub ScreenUpdate(ByVal bState As Boolean)
    
#If Mac Then
    Application.ScreenUpdating = bState
    Exit Sub
#End If
    
    Dim hwnd As LongPtr
    hwnd = GethWndWorkbook
    
    'Using SendMessage:
    ' - Turn off redraw for faster and smoother action:
    '     SendMessage hEdit, %WM_SETREDRAW, 0, 0
    ' - Turn on redraw again and refresh:
    '     SendMessage hEdit, %WM_SETREDRAW, 1, 0

    Dim lResult As LongPtr
    If bState Then
        lResult = SendMessage(hwnd, WM_SETREDRAW, 1&, 0&)
        lResult = InvalidateRect(hwnd, 0&, 1&)
        lResult = UpdateWindow(hwnd)
        DoEvents
    Else
        lResult = SendMessage(hwnd, WM_SETREDRAW, 0&, 0&)
    End If
    DoEvents
    
End Sub

Private Function GethWndWorkbook() As LongPtr

    Dim hWndXLDESK As LongPtr
    hWndXLDESK = FindWindowEx(Application.hwnd, 0, "XLDESK", vbNullString)
    
    GethWndWorkbook = FindWindowEx(hWndXLDESK, 0, vbNullString, ThisWorkbook.Name)
    
End Function

