Attribute VB_Name = "ModGlobal"
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
' Module    : ModGlobal
' DateTime  : 05/20/2013
' Author    : José García Herruzo
' Purpose   : This module contents global variables
' References: N/A
' Functions : N/A
' Procedures:
'               1-InitializateVar
'----------------------------------------------------------------------------------------

Public Server_Sim As Boolean
Public str_SimPath As String
Public str_BalPath As String

Public ws_Setup As Worksheet
Public ws_Log As Worksheet

Public str_BECSPath As String
Public str_ProjectWs As String
Public str_ProjecColumn As String
Public str_Projects() As String
'---------------------------------------------------------------------------------------
' Procedure : Initializate_Var
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Initializate global variables
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Initializate_Var()

Set ws_Setup = ThisWorkbook.Worksheets("Setup")
Set ws_Log = ThisWorkbook.Worksheets("Log")

str_BalPath = ws_Setup.Range("B3").Value

str_BECSPath = ws_Setup.Range("B4").Value
str_ProjectWs = ws_Setup.Range("B5").Value
str_ProjecColumn = ws_Setup.Range("B6").Value

Call Update_Components_Range

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Close_Var
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Initializate global variables
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Close_Var()

Set ws_Setup = Nothing

End Sub
