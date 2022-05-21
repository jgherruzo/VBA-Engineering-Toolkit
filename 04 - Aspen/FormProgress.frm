VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormProgress 
   Caption         =   "APE Progress Information"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   OleObjectBlob   =   "FormProgress.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "FormProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'           .---.        .-----------
'          /     \  __  /    ------
'         / /     \(..)/    -----
'        //////   ' \/ `   ---
'       //// / // :    : ---
'      // /   /  /`    '--
'     // /        //..\\
'   o===|========UU====UU=====-  -==========================o
'                '//||\\`
'                       DEVELOPED BY JGH
'
'   -=====================|===o  o===|======================-
Option Explicit

'---------------------------------------------------------------------------------------
' Event     : UserForm_QueryClose
' Author    : José García Herruzo
' Purpose   : You cannot close this form from X button
'---------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
     
End Sub
