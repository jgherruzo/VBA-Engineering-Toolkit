Attribute VB_Name = "ModImportInputStream"
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
' Module    : ModImportInputStream
' DateTime  : 06/26/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for extract the input streams
' References: N/A
' Functions :
'               1-Get_Input_Stream_Worksheet_Destination
'               2-Get_Input_Stream_Column_Destination
' Procedures:
'               1-Update_Design_Basis
'               2-CheckingInputStream
' Updates   :
'       DATE        USER    DESCRIPTION
'       06/26/2013  JGH     Project code are eliminated from balance file name. Code is
'                           modified in order to use the new name code
'       06/26/2013  JGH     Reset prior streams only if stream <>1 (Deactivating
'                           automatic stream update
'       06/27/2013  JGH     Close source workbook without msg. Deactivate uploading
'                           and msg
'       07/17/2013  JGH     Problem with more than one stream. Code is modified
'       08/21/2013  JGH     Problem with opened workbook. Code is modified
'       08/28/2013  JGH     Problem with opened workbook. Code is modified
'       02/01/2014  JGH     Update_Input_Streams is modified to not close workbook which
'                           were already opened.
'----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : Update_Input_Streams
' DateTime  : 06/26/2013
' Author    : José García Herruzo
' Purpose   : This function controls input stream update process. It returns a number
'             depending on the result.
'               0--> Input stream information is uncomplete
'               1--> Is an utility area and does not need extract input streams
'               2--> The process is completed
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function Update_Input_Streams() As Integer

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim lastrow As Integer
Dim input_Info() As String
Dim wb_1 As Workbook
Dim ws_mySheet As Worksheet
Dim str_myPath As String
Dim myStream() As Variant
Dim int_myColumn As Integer
Dim ra_Destination As Range
Dim bol_wb_open As Boolean
Dim bol_wb_Already_opened As Boolean
Dim str_HelpVar As String
Dim str_myRange As String
Dim help() As String
Dim a As Integer

'-- Check if the input required information is completed --
bol_wb_open = False
lastrow = ws_Input.Range("A16000").End(xlUp).Row
str_myPath = WS_Setup.Range("C3").Value
bol_wb_Already_opened = False

If lastrow <= 1 Then

    '-- If not, the rest of the process is stopped --
    Update_Input_Streams = 0
    MsgBox "The required information about input stream is incompleted. Please, go to input worksheet and complete it", vbOKOnly + vbCritical
    Exit Function

End If

'-- Input information is updated --
lastrow = lastrow - 2
ReDim input_Info(lastrow, 3)

For i = 0 To lastrow

    input_Info(i, 0) = ws_Input.Range("A2").Offset(i, 0).Value
    input_Info(i, 1) = ws_Input.Range("A2").Offset(i, 1).Value
    input_Info(i, 2) = ws_Input.Range("A2").Offset(i, 2).Value
    
Next i

If input_Info(0, 0) = 1 Then

    '-- Uitility area --
    Update_Input_Streams = 1
    Exit Function

End If

'-- Reset prior streams--
ws_Input_Destination.Range("D1:AA16000").ClearContents

    '-- set input stream balance file complete path--
    str_HelpVar = str_myPath & "\" & input_Info(0, 1) & "\" & input_Info(0, 1) & ".02." _
            & input_Info(0, 2) & ".xls"
            '-- check if wb is opened --
            bol_wb_open = IsFileOpen(str_HelpVar)
            
            '-- if workbook is opened --
            If bol_wb_open = True Then
                '-- Active the flag --
                bol_wb_Already_opened = True
            
            End If
            
'-- Now, check if it is the first time the input streams have extracted --
For i = 0 To lastrow

    xlStartSettings ("Extracting streams " & i + 1 & " to " & lastrow + 1)
            
            If bol_wb_open = False Then
                
                Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "1", "Update Flag")
                Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "1", "Msg Flag")
                Set wb_1 = Workbooks.Open(str_HelpVar)
            
            Else
                
                help = Split(str_HelpVar, "\")
                a = UBound(help)
                Set wb_1 = Workbooks(help(a))
            
            End If
            
            For Each ws_mySheet In wb_1.Worksheets
            
                If InStr(ws_mySheet.Name, "-NT-") <> 0 Then
                
                    For j = 0 To ReturnColumn(wb_1.Name, ws_mySheet.Name, "D1")
                
                        If input_Info(i, 0) = ws_mySheet.Range("D1").Offset(0, j).Value Then
                        
                            ReDim myStream(lon_TOTAL_PAR)
                            
                            For k = 0 To lon_TOTAL_PAR
                            
                                myStream(k) = ws_mySheet.Range("D1").Offset(k, j).Value
                
                            Next k
                            
                        End If
                
                    Next j
                    
                End If
            
            Next ws_mySheet
            
            Set ra_Destination = ws_Input_Destination.Range("D1")
            int_myColumn = ReturnColumn(ThisWorkbook.Name, ws_Input_Destination.Name, "D1")
            
            '-- Values are updated --
            For k = 0 To lon_TOTAL_PAR
                            
                ra_Destination.Offset(k, int_myColumn).Value = myStream(k)
                
            Next k
            
            If i < lastrow Then
            
                If input_Info(i, 1) = input_Info(i + 1, 1) And input_Info(i, 2) = input_Info(i + 1, 2) Then
                
                    bol_wb_open = True
                    
                Else
                    
                    '-- Check if w1 had already opened --
                    If bol_wb_Already_opened = True Then
                        
                        bol_wb_Already_opened = False
                        str_HelpVar = str_myPath & "\" & input_Info(i + 1, 1) & "\" & input_Info(i + 1, 1) & ".02." _
                        & input_Info(i + 1, 2) & ".xls"
                    
                                '-- check if wb is opened --
                        bol_wb_open = IsFileOpen(str_HelpVar)
                        
                        '-- if workbook is opened --
                        If bol_wb_open = True Then
                            '-- Active the flag --
                            bol_wb_Already_opened = True
                        
                        End If
                    
                    Else
                        
                        '-- if not, it is closed --
                        Application.DisplayAlerts = False
                        wb_1.Close
                        Application.DisplayAlerts = True
                        Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "0", "Msg Flag")
                        Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "0", "Update Flag")
                        
                            '-- check the next --
                            str_HelpVar = str_myPath & "\" & input_Info(i + 1, 1) & "\" & input_Info(i + 1, 1) & ".02." _
                            & input_Info(i + 1, 2) & ".xls"
                        bol_wb_open = IsFileOpen(str_HelpVar)
                        
                        '-- if workbook is opened --
                        If bol_wb_open = True Then
                            '-- Active the flag --
                            bol_wb_Already_opened = True
                        
                        End If
                        
                    End If
                    
                End If
                
            Else
            
                If bol_wb_open = False Or bol_wb_Already_opened = False Then
                
                    Application.DisplayAlerts = False
                    wb_1.Close
                    Application.DisplayAlerts = True
                    Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "0", "Msg Flag")
                    Call Update_Data_Into_ExcelCell(str_HelpVar, "Setup", "0", "Update Flag")
                
                End If
                
            End If
            
Next i

Set wb_1 = Nothing

Update_Input_Streams = 2

End Function

'---------------------------------------------------------------------------------------
' Function  : Get_Input_Stream_Column_Destination
' DateTime  : 05/10/2013
' Author    : José García Herruzo
' Purpose   : Return the offset column of the searched stream is
' Arguments :
'             str_myStream              --> Seacrhed stream name
'---------------------------------------------------------------------------------------
Private Function Get_Input_Stream_Column_Destination(ByVal str_myStream As String) As Integer

Dim ws_mySheet As Worksheet

For Each ws_mySheet In ThisWorkbook.Worksheets

    If InStr(ws_mySheet.Name, "-NT-") <> 0 Then
    
        For i = 0 To ReturnColumn(ThisWorkbook.Path & "\" & ThisWorkbook.Name, ws_mySheet.Name, "D1")
        
            If ws_mySheet.Range("D1").Offset(0, i).Value = str_myStream Then
            
                Get_Input_Stream_Destination = i
                End Function
            
            End If
        
        Next i
        
    End If

Next ws_mySheet

End Function

'---------------------------------------------------------------------------------------
' Function  : Get_Input_Stream_Worksheet_Destination
' DateTime  : 05/10/2013
' Author    : José García Herruzo
' Purpose   : Return the name of the worksheet of the searched stream is
' Arguments :
'             str_myStream              --> Seacrhed stream name
'---------------------------------------------------------------------------------------
Private Function Get_Input_Stream_Worksheet_Destination(ByVal str_myStream As String) As String

Dim ws_mySheet As Worksheet

For Each ws_mySheet In ThisWorkbook.Worksheets

    If InStr(ws_mySheet.Name, "-NT-") <> 0 Then
    
        For i = 0 To ReturnColumn(ThisWorkbook.Path & "\" & ThisWorkbook.Name, ws_mySheet.Name, "D1")
        
            If ws_mySheet.Range("D1").Offset(0, i).Value = str_myStream Then
            
                Get_Input_Stream_Worksheet_Destination = ws_mySheet.Name
                End Function
            
            End If
        
        Next i
        
    End If

Next ws_mySheet

End Function






