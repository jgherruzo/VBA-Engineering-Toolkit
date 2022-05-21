Attribute VB_Name = "ModTBP"
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
'----------------------------------------------------------------------------------------
' Module    : ModTBP
' DateTime  : 09/11/2016
' Author    : José García Herruzo
' Purpose   : This module contents procedures and function extract TBP curve for a given
'               FTL stream
' References: N/A
' Functions :
'               1-xfGiveTBP
' Procedures:
'               1-xpLoadVar
'               2-xpResetVar
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Dim ws_TB As Worksheet
Dim ws_Stream As Worksheet
Dim ws_Result As Worksheet

Dim arr_TB() As Variant
Dim arr_Stream() As Variant
'---------------------------------------------------------------------------------------
' Procedure : xpLoadVar
' DateTime  : 09/11/2016
' Author    : José García Herruzo
' Purpose   : Load general parameter to be used into this module
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub xpLoadVar()

Dim i As Integer
Dim j As Integer

Dim int_Counter As Integer

'-- select worksheets --
Set ws_TB = ThisWorkbook.Worksheets("TB")
Set ws_Stream = ThisWorkbook.Worksheets("Stream info")
Set ws_Result = ThisWorkbook.Worksheets("Result")

'-- Extract TB array --
int_Counter = ws_TB.Range("A" & ws_TB.Rows.Count).End(xlUp).Row - 3

ReDim arr_TB(int_Counter, 1)

For i = 0 To int_Counter

    For j = 0 To 1
    
        arr_TB(i, j) = ws_TB.Range("A3").Offset(i, j).Value
    
    Next j

Next i

'-- Extract Stream array --
int_Counter = ws_Stream.Range("A" & ws_Stream.Rows.Count).End(xlUp).Row - 2

ReDim arr_Stream(int_Counter, 1)

For i = 0 To int_Counter

    For j = 0 To 1
    
        arr_Stream(i, j) = ws_Stream.Range("A2").Offset(i, j).Value
    
    Next j

Next i

End Sub
'---------------------------------------------------------------------------------------
' Procedure : xpResetVar
' DateTime  : 09/11/2016
' Author    : José García Herruzo
' Purpose   : Reset general parameter used into this module
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub xpResetVar()

'-- select worksheets --
Set ws_TB = Nothing
Set ws_Stream = Nothing
Set ws_Result = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Function  : xfGiveTBP
' DateTime  : 09/11/2016
' Author    : José García Herruzo
' Purpose   : Calculate mass weight below the specified temperature
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function xfGiveTBP(ByVal dbl_myTemperature As Double) As Double

Dim dbl_Flow As Double
Dim dbl_Value As Double
Dim dbl_accumulate As Double

Dim i As Integer
Dim j As Integer

Dim int_limit As Integer
Dim int_limit2 As Integer

Call xpLoadVar

int_limit = UBound(arr_TB)
int_limit2 = UBound(arr_Stream)

dbl_Flow = ws_Result.Range("ra_flowrate").Value

'-- For each compound --
For i = 0 To int_limit
    '-- if it is lower TB than the specified --
    If arr_TB(i, 1) < dbl_myTemperature Then
        '-- Search into stream info the same compound --
        For j = 0 To int_limit2
        
            If arr_TB(i, 0) = arr_Stream(j, 0) Then
            
                dbl_accumulate = dbl_accumulate + arr_Stream(j, 1)
                Exit For
                
            End If
            
        Next j
        
    End If

Next i

dbl_Value = dbl_accumulate / dbl_Flow

xfGiveTBP = dbl_Value

End Function


