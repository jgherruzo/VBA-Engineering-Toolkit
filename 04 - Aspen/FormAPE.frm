VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormAPE 
   Caption         =   "Aspen Properties Extractor"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6315
   OleObjectBlob   =   "FormAPE.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "FormAPE"
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
' Event     : chk_Local_Click
' Author    : José García Herruzo
' Purpose   : Select sim path source
'---------------------------------------------------------------------------------------
Private Sub chk_Local_Click()

If chk_Local.Value = True Then

    chk_Server.Value = False

Else

    chk_Server.Value = True

End If

End Sub

'---------------------------------------------------------------------------------------
' Event     : chk_Server_Click
' Author    : José García Herruzo
' Purpose   : Select sim path source
'---------------------------------------------------------------------------------------
Private Sub chk_Server_Click()

If chk_Server.Value = True Then

    chk_Local.Value = False

Else

    chk_Local.Value = True

End If

End Sub

'---------------------------------------------------------------------------------------
' Event     : cmdClose_Click
' Author    : José García Herruzo
' Purpose   : Unload the form and close this workbook
'---------------------------------------------------------------------------------------
Private Sub cmdClose_Click()

Application.Visible = True

Unload Me

Application.DisplayAlerts = False
ThisWorkbook.Close
Application.DisplayAlerts = True

End Sub

'---------------------------------------------------------------------------------------
' Event     : cmdNext_Click
' Author    : José García Herruzo
' Purpose   : Show main form
'---------------------------------------------------------------------------------------
Private Sub cmdNext_Click()

'-- Load source path --
Server_Sim = chk_Server.Value

If Server_Sim = True Then

    str_SimPath = ws_Setup.Range("B1").Value

Else

    str_SimPath = ws_Setup.Range("B2").Value

End If

Unload Me
FormSim.Show

End Sub

'---------------------------------------------------------------------------------------
' Event     : UserForm_Initialize
' Author    : José García Herruzo
' Purpose   : If sim software is not installed the only way is close this file
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

Application.Visible = False

'-- If aspen is installed you can continue, if not, you must close this app --
If IsSoftwareInstalled(ThisWorkbook.Worksheets("Installed_Software"), "A1", "Aspen Plus") = False Then
    
    L_AspenStatus.Caption = "Aspen plus software is not installed in this computer. Please close this app."
    L_AspenStatus.BackColor = RGB(254, 72, 25)
    
    Fr_SimFileLocation.Visible = False
    cmdNext.Enabled = False
    
Else

    L_AspenStatus.Caption = "Aspen plus software is installed in this computer. Please select a file source and continue."
    L_AspenStatus.BackColor = RGB(146, 208, 80)

End If

End Sub

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

