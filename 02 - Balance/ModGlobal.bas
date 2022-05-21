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
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : This module contents global variables
' References: N/A
' Functions : N/A
' Procedures:
'               1-InitializateVar
'----------------------------------------------------------------------------------------
Public str_Design_Basis_Extract_Range As String
Public str_Design_Basis_Import_Range As String
Public str_Pressure_Drop_Extract_Range As String
Public str_Utilities_Range As String

Public str_ws_help As String

Public ws_Utilities As Worksheet
Public WS_Setup As Worksheet
Public WS_Exchangers As Worksheet
Public ws_Input As Worksheet
Public ws_Input_Destination As Worksheet

Public var_Utilities_Lower_Limit As Variant
Public var_Utilities_Upper_Limit As Variant
Public var_SimultaneityCoef As Variant

'---------------------------------------------------------------------------------------
' Procedure : Initializate_Var
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Initializate global variables
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Initializate_Var()

Set WS_Setup = ThisWorkbook.Worksheets("Setup")
Set WS_Exchangers = ThisWorkbook.Worksheets("Heater-Cooler")
'Set ws_Utilities = ThisWorkbook.Worksheets("D-07000-NT-001")
Set ws_Input = ThisWorkbook.Worksheets("Input Information")
Set ws_Input_Destination = ThisWorkbook.Worksheets("Input Streams")

str_Design_Basis_Extract_Range = "A1:C16000"
str_Pressure_Drop_Extract_Range = "A1:D16000"
str_Design_Basis_Import_Range = "A2"
'str_Utilities_Range = "E1"

str_ws_help = "Pressure Drop"

'var_Utilities_Lower_Limit = 7000
'var_Utilities_Upper_Limit = 7800
'var_SimultaneityCoef = ThisWorkbook.Worksheets("DB-07000").Range("B4").Value

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

Set WS_Setup = Nothing
Set WS_Exchangers = Nothing
'Set ws_Utilities = Nothing

End Sub
