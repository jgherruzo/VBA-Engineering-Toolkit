Attribute VB_Name = "ModStreamList"
'    _______----__________                 __________----_______
'    \------____-------___--__---------__--___-------____------/
'     \//////// / / / / / \   _-------_   / \ \ \ \ \ \\\\\\\\/
'       \////-/-/------/_/_| /___   ___\ |_\_\------\-\-\\\\/
'         --//// / /  /  //|| (O)\ /(O) ||\\  \  \ \ \\\\--
'              ---__/  // /| \_  /V\  _/ |\ \\  \__---
'                   -//  / /\_ ------- _/\ \  \\-
'                     \_/_/ /\---------/\ \_\_/
'                         ----\---|---/----
'                              \--|--/
'          ===================(((===)))===================
'          _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-DEVELOPED BY JESG
'          ===============================================
Option Explicit
'----------------------------------------------------------------------------------------
' Module    : ModStreamList
' DateTime  : 03/05/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : This module contents funcions and procedures to update the fluids code
'
' References: N/A
' Functions :
'
' Procedures:
'               1-Service_Update
'               2-Initializate_StreamListVar
'               3-Close_StreamListVar
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Public str_WStemp As String
Public row_number As Long

Public str_PipeSpec  As String
Public str_RootPath As String
Public str_Version As String


'---------------------------------------------------------------------------------------
' Procedure : Service_Update()
' DateTime  : 03/05/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Extract Pipe Spec using ADO function
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Service_Update()

Dim LastColumn2 As Integer
Dim Lastrow1 As Integer
Dim ws As Worksheet

Call Initializate_StreamListVar
Call Update_Components_Range

str_WStemp = "Selection temp"

    '-- it creates a new temp sheet and after that, pastes pipe spec for drop list -
    If SheetExist(ThisWorkbook.Name, str_WStemp) = True Then
    Worksheets(str_WStemp).Delete
    End If
    
Call AddNewSheet(ThisWorkbook.Name, str_WStemp)
ThisWorkbook.Worksheets(str_WStemp).Visible = False
Call Extract_Data_From_Excel_WithTittle(str_PipeSpec, "Selection", "B1:C1600", ThisWorkbook, str_WStemp, "A2")
Lastrow1 = ThisWorkbook.Worksheets(str_WStemp).Cells(Cells.Rows.Count, 1).End(xlUp).Row

'-- Create a drop list for each -NT- sheet
For Each ws In ThisWorkbook.Worksheets
    
    If InStr(1, ws.Name, "-NT-", vbTextCompare) <> 0 Then
    row_number = ws.Range(RA_SERVICE).Row
    LastColumn2 = ws.Cells(1, Cells.Columns.Count).End(xlToLeft).Column
    
    ws.Activate
    ws.Range(Cells(row_number, 4), Cells(row_number, LastColumn2)).Select
        
        With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="='Selection temp'!$B$2:$B$" & Lastrow1
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = False
        .ShowError = False
        End With
        
    End If

Next ws

WS_Setup.Range("R2").Value = 1
WS_Setup.Range("R3").Value = "Updated"

Close_StreamListVar
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Service_Update()
' DateTime  : 03/011/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Load stream list variables used in this module
' Arguments : N/A
'---------------------------------------------------------------------------------------

Private Sub Initializate_StreamListVar()

Set WS_Setup = ThisWorkbook.Worksheets("Setup")

'-- Load General Path
str_RootPath = WS_Setup.Range("C4").Value
str_Version = WS_Setup.Range("Q2").Value

'-- Build PipeSpec file Path --
str_PipeSpec = str_RootPath & "\Pipes_Spec." & str_Version & ".xls"
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Service_Update()
' DateTime  : 03/011/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Close stream list variables used in this module
' Arguments : N/A
'---------------------------------------------------------------------------------------

Private Sub Close_StreamListVar()
Set WS_Setup = Nothing
End Sub













