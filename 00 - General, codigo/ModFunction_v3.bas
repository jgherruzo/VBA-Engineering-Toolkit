Attribute VB_Name = "ModFunction_v3"
'           .---.        .-----------
'          /     \  __  /    ------
'         / /     \(..)/    -----
'        //////   ' \/ `   ---
'       //// / // :    : ---
'      // /   /  /`    '--
'     // /        //..\\
'   o===|========UU====UU=====-  -==========================o
'                '//||\\`
'                       DEVELPOPED BY JGH
'
'   -=====================|===o  o===|======================-
Option Explicit
'----------------------------------------------------------------------------------------
' Module    : ModFunction_v3
' DateTime  : 17/02/2014
' Author    : Jerid
' Purpose   : This module contents function to move the mouse
' References: N/A
' Functions :
'               1-GetWindowHandle
' Procedures: N/A
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
'**Win32 API Declarations
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'**Win32 API User Defined Types
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type

'Private Sub test()
'Dim Rec As RECT
'Get Left, Right, Top and Bottom of Form1
'GetWindowRect GetWindowHandle, Rec
'Set Cursor position on X
'SetCursorPos Rec.Right - 600, Rec.Top + 400
'End Sub

Private Function GetWindowHandle() As Long

Const CLASSNAME_MSExcel = "XLMAIN"

'Gets the Apps window handle, since you can't use App.hInstance in VBA (VB Only)
GetWindowHandle = FindWindow(CLASSNAME_MSExcel, vbNullString)
End Function

