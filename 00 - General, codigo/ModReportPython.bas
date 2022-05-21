Attribute VB_Name = "ModReportPython"
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
'----------------------------------------------------------------------------------------
' Module    : ModReport
' DateTime  : 25/06/2020
' Author    : José García Herruzo
' Purpose   : This module contents function to report with python
' References: N/A
' Functions : N/A
' Procedures:
'               1-xp_Export_myCSV
'               2-xp_Make_myReport
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Public Sub xp_Export_myCSV()

Dim ws As Worksheet
Dim ws1 As Worksheet
Dim ws2 As Worksheet
Dim ws3 As Worksheet

Dim str_File As String
Dim str_Temp As String
Dim str_Year As String
Dim str_Week As String
Dim str_Path As String

Set ws = ThisWorkbook.Worksheets("Menu")
Set ws1 = ThisWorkbook.Worksheets("SAP1")
Set ws2 = ThisWorkbook.Worksheets("SAP2")
Set ws3 = ThisWorkbook.Worksheets("SAP3")

str_Week = ws.Range("ra_Week")
str_Year = Right(ws.Range("ra_Year"), 2)

str_Path = "\\hue-4\Usuarios\Publicaciones\Acido\Fbr. Ácido\00-General\04-Reports\98 - General\GEN - 2020 - 09 - Eficiencia de Contactos\01-Datos\"

Call SeparadorPunto

Application.ScreenUpdating = False

'-- SAP1 --
str_Temp = "SAP1_" & str_Year & "W" & str_Week

str_File = str_Path & str_Temp

Call xp_Export_CSV(ws1, str_File)

'-- SAP2 --
str_Temp = "SAP2_" & str_Year & "W" & str_Week

str_File = str_Path & str_Temp

Call xp_Export_CSV(ws2, str_File)

'-- SAP3 --
str_Temp = "SAP3_" & str_Year & "W" & str_Week

str_File = str_Path & str_Temp

Call xp_Export_CSV(ws3, str_File)

Call SeparadorComa

End Sub

Public Sub xp_Make_myReport(ByVal int_Key As Integer)

Dim ws As Worksheet

Dim str_Temp As String
Dim str_Temp2 As String
Dim str_Year As String
Dim str_Week As String
Dim str_myScript As String

Set ws = ThisWorkbook.Worksheets("Menu")

str_Week = ws.Range("ra_Week")
str_Year = Right(ws.Range("ra_Year"), 2)

str_Temp = str_Year & "W" & str_Week

str_myScript = "C:\Users\jgarciah2\OneDrive - Freeport-McMoRan Inc\05 - Data Science\00-Mis Proyectos\00-Seguimiento de emisiones\JGH_Seguimiento_Emisiones.py"

str_Temp2 = "Python.exe """ & str_myScript & """ " & str_Temp & " " & int_Key
Shell str_Temp2
    
End Sub
