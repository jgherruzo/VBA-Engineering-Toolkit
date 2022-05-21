Attribute VB_Name = "ModImportBalanceData"
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
' Module    : ModImportBalanceData
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for extract design basis &
'             Exchangers  pressure drop
' References: N/A
' Functions : N/A
' Procedures:
'               1-Update_Design_Basis
' Updates   :
'       DATE        USER    DESCRIPTION
'       08/30/2013  JGH     Exchangers efficiency
'       04/04/2014  JGH     Calling Ado code is modify according last ADO updated
'----------------------------------------------------------------------------------------
'-- Design Basis --
Public DBPath As String
Public DBName As String
Public neededDB() As String
Public DBVersion() As String

'-- Pressure drop --
Public PDPath As String
Public PDName As String
Public PDversion As String
'---------------------------------------------------------------------------------------
' Procedure : Update_Design_Basis
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Extract DB using ADO connection
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Update_Design_Basis()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim mySheetName As String
Dim lastrow As Integer
Dim Lastrow2 As Integer
Dim myDB As String
Dim str_DB As String

' -- Required DB are updated --
DBPath = WS_Setup.Range("C2").Value
lastrow = WS_Setup.Range("A1").End(xlDown).Row

ReDim neededDB(lastrow - 2)
ReDim DBVersion(lastrow - 2)

For i = 1 To lastrow - 1

    neededDB(i - 1) = WS_Setup.Range("A1").Offset(i, 0).Value
    DBVersion(i - 1) = WS_Setup.Range("A1").Offset(i, 1).Value
    
Next i

' -- It is check if the DB sheet exist at current WB --
For i = 0 To lastrow - 2
    
    myDB = "DB-" & neededDB(i)
    If SheetExist(ThisWorkbook.Name, myDB) = False Then
        
        Call AddNewSheet(ThisWorkbook.Name, myDB)
        
    Else
    
        Dim a As Integer
    
        a = ThisWorkbook.Worksheets(myDB).Range("A16000").End(xlUp).Row
        ThisWorkbook.Worksheets(myDB).Range("A2:D" & a & "").ClearContents
    
    End If

Next i

' -- DB are updated --
For k = 0 To lastrow - 2
    
    xlStartSettings ("Extracting design basis " & k + 1 & " to " & lastrow - 1)
    
    str_DB = DBPath & "\" & neededDB(k) & "\DB." & neededDB(k) & "." & DBVersion(k) & ".xls"
    myDB = "DB-" & neededDB(k)
       
    Call Extract_Data_From_Excel(str_DB, myDB, str_Design_Basis_Extract_Range, ThisWorkbook, myDB, str_Design_Basis_Import_Range)
    
Next k

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Update_Pressure_Drop
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : Extract DB using ADO connection
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Update_Pressure_Drop()

Dim i As Integer
Dim j As Integer
Dim Exchanger_Counter As Integer
Dim int_TotalExchanger_Counter As Integer
Dim str_PD_Complete_Path As String
Dim TotalExchanger_Info() As Variant
Dim str_myExchanger As String

'-- Reset Pressure drops values --
Exchanger_Counter = WS_Exchangers.Range("B16000").End(xlUp).Row
Exchanger_Counter = Exchanger_Counter - 2

If Exchanger_Counter < 1 Then

    Exit Sub
    
End If

WS_Exchangers.Range("F3:F" & Exchanger_Counter + 2).ClearContents
WS_Exchangers.Range("J3:J" & Exchanger_Counter + 2).ClearContents

xlStartSettings ("Extracting exchangers pressure drop")

'-- Update Import File --
PDPath = WS_Setup.Range("K2").Value
PDName = WS_Setup.Range("L2").Value
PDversion = WS_Setup.Range("M2").Value
  
str_PD_Complete_Path = PDPath & "\" & PDName & "." & PDversion & ".xls"

'-- a sheet is add in order to paste the heaters values --
ThisWorkbook.Worksheets.Add
ActiveSheet.Name = str_ws_help

Call Extract_Data_From_Excel(str_PD_Complete_Path, str_ws_help, str_Pressure_Drop_Extract_Range, ThisWorkbook, str_ws_help, str_Design_Basis_Import_Range)

int_TotalExchanger_Counter = ThisWorkbook.Worksheets(str_ws_help).Range("A16000").End(xlUp).Row

int_TotalExchanger_Counter = int_TotalExchanger_Counter - 2
ReDim TotalExchanger_Info(int_TotalExchanger_Counter, 3)

'-- Pressure drop values are updated --
For i = 0 To int_TotalExchanger_Counter

    TotalExchanger_Info(i, 0) = ThisWorkbook.Worksheets(str_ws_help).Range("A2").Offset(i, 1).Value
    TotalExchanger_Info(i, 1) = ThisWorkbook.Worksheets(str_ws_help).Range("A2").Offset(i, 2).Value
    TotalExchanger_Info(i, 2) = ThisWorkbook.Worksheets(str_ws_help).Range("A2").Offset(i, 3).Value
    TotalExchanger_Info(i, 3) = ThisWorkbook.Worksheets(str_ws_help).Range("A2").Offset(i, 4).Value
    
Next i
    
For j = 0 To Exchanger_Counter

    For i = 0 To int_TotalExchanger_Counter
    
        str_myExchanger = WS_Exchangers.Range("B3").Offset(j, 0).Value
        
        If str_myExchanger = TotalExchanger_Info(i, 0) Then
        
            WS_Exchangers.Range("B3").Offset(j, 4).Value = TotalExchanger_Info(i, 1)
            WS_Exchangers.Range("B3").Offset(j, 8).Value = TotalExchanger_Info(i, 2)
            WS_Exchangers.Range("B3").Offset(j, 11).Value = TotalExchanger_Info(i, 3)
        
        End If
        
    Next i

Next j

'-- The new sheet is deleted--
Application.DisplayAlerts = False
ThisWorkbook.Worksheets(str_ws_help).Delete
Application.DisplayAlerts = True

End Sub
