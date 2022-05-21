Attribute VB_Name = "ModLiga"
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
'-----------------------------------------------------------------------------------------
' Module      : ModFluor
' DateTime    : 18/01/2021
' Author      : José García Herruzo
' Purpose     : Fluor followup
' References  : N/A
' Requirements: N/A
' Functions   :
'               01-xfGetFinHF
'               01-xfGetHginHF
' Procedures  : N/A
' Updates     :
'       DATE        USER    DESCRIPTION
'       N/A
'-----------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function  : xfGetFinHF
' DateTime  : 18/01/2021
' Author    : José García Herruzo
' Purpose   : Look for the last theoretical F
' Arguments :
'               dbl_Date            --> Date to look for
'---------------------------------------------------------------------------------------
Public Function xfGetFinHF(ByVal dt_date As Date) As Double

Dim i As Long
Dim arr_temp() As Variant
Dim lon_row As Long

Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets("Liga Teórica")

lon_row = ws.Range("A" & ws.Rows.Count).End(xlUp).Row - 2

ReDim arr_temp(lon_row, 1)

For i = 0 To lon_row

    arr_temp(i, 0) = ws.Range("A2").Offset(i, 0).Value
    arr_temp(i, 1) = ws.Range("A2").Offset(i, 15).Value
    
Next i


For i = 0 To lon_row

    If arr_temp(i, 0) > dt_date Then
    
        xfGetFinHF = arr_temp(i - 1, 1)
        Exit For
    
    End If
    
Next i

End Function

'---------------------------------------------------------------------------------------
' Function  : xfGetHginHF
' DateTime  : 20/01/2021
' Author    : José García Herruzo
' Purpose   : Look for the last theoretical Hg
' Arguments :
'               dbl_Date            --> Date to look for
'---------------------------------------------------------------------------------------
Public Function xfGetHginHF(ByVal dt_date As Date) As Double

Dim i As Long
Dim arr_temp() As Variant
Dim lon_row As Long

Dim ws As Worksheet

Set ws = ThisWorkbook.Worksheets("Liga Teórica")

lon_row = ws.Range("A" & ws.Rows.Count).End(xlUp).Row - 2

ReDim arr_temp(lon_row, 1)

For i = 0 To lon_row

    arr_temp(i, 0) = ws.Range("A2").Offset(i, 0).Value
    arr_temp(i, 1) = ws.Range("A2").Offset(i, 14).Value
    
Next i


For i = 0 To lon_row

    If arr_temp(i, 0) > dt_date Then
    
        xfGetHginHF = arr_temp(i - 1, 1)
        Exit For
    
    End If
    
Next i

End Function

