Attribute VB_Name = "Equilibrium"
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
' Module    : ModEquilibrium
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : This module contents procedures and function to interpolate into a
'               equilibrium table
' References: N/A
' Functions :
'               1-xlInterpolate
'               2-xlGetTemp_TXY_y
'               3-xlGetTemp_TXY_x
'               4-xlGetx_TXY_T
'               5-xlGetx_TXY_y
'               6-xlGety_TXY_T
'               7-xlGety_TXY_x
'               2-xlGetPress_PXY_y
'               3-xlGetPress_PXY_x
'               4-xlGetx_PXY_P
'               5-xlGetx_PXY_y
'               6-xlGety_PXY_P
'               7-xlGety_PXY_x
' Procedures:
'               1-xsGeneral
' Updates   :
'       DATE        USER    DESCRIPTION
'       19/01/2015  JGH     xlGety_TXY_x is modified, argument into interpolate equation
'                           were backwards
'----------------------------------------------------------------------------------------
Dim dou_MWL As Double
Dim dou_MWH As Double
Dim dou_TXYParCounter As Double
Dim dou_PXYParCounter As Double

Dim i As Long
Dim j As Long

Dim arr_Temp() As Double

Dim ws_TXY As Worksheet
Dim ws_PXY As Worksheet

Dim bol_PXY As Boolean
Dim bol_TXY As Boolean
'---------------------------------------------------------------------------------------
' Procedure : xsGeneral
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Load general parameter to be used into this module
' Arguments :
'             str_myName                --> Spreadsheet specific name
'---------------------------------------------------------------------------------------
Private Sub xsGeneral(ByVal str_myName As String)

Dim ws As Worksheet

On Error GoTo myhandler:
'-- set band --
bol_PXY = False
bol_TXY = False

'-- set source --
            '-- TXY --
If str_myName = "Default SPD name" Then

    For Each ws In ThisWorkbook.Worksheets
        
        If InStr(1, ws.Name, "TXY", vbTextCompare) <> 0 Then
        
            Set ws_TXY = ThisWorkbook.Worksheets(ws.Name)
            '-- count parameter --
            dou_TXYParCounter = ws_TXY.Range("B" & Rows.Count).End(xlUp).Row - 2
            '-- TXY spreadsheet exist --
            bol_TXY = True
            dou_MWL = ws_TXY.Range("C1").Value
            dou_MWH = ws_TXY.Range("D1").Value
            
            Exit For
            
        End If
        
    Next ws

Else

    For Each ws In ThisWorkbook.Worksheets
        
        If InStr(1, ws.Name, "TXY", vbTextCompare) <> 0 Then
        
            Set ws_TXY = ThisWorkbook.Worksheets("TXY " & str_myName)
            '-- count parameter --
            dou_TXYParCounter = ws_TXY.Range("B" & Rows.Count).End(xlUp).Row - 2
            '-- TXY spreadsheet exist --
            bol_TXY = True
            dou_MWL = ws_TXY.Range("C1").Value
            dou_MWH = ws_TXY.Range("D1").Value
            
            Exit For
            
        End If
        
    Next ws
End If

            '-- PXY --
If str_myName = "Default SPD name" Then

    For Each ws In ThisWorkbook.Worksheets
        
        If InStr(1, ws.Name, "PXY", vbTextCompare) <> 0 And InStr(1, ws.Name, "PXY", vbTextCompare) < 7 Then
        
            Set ws_PXY = ThisWorkbook.Worksheets(ws.Name)
            '-- count parameter --
            dou_PXYParCounter = ws_PXY.Range("B" & Rows.Count).End(xlUp).Row - 2
            '-- TXY spreadsheet exist --
            bol_PXY = True
            dou_MWL = ws_PXY.Range("C1").Value
            dou_MWH = ws_PXY.Range("D1").Value
            
            Exit For
            
        End If
        
    Next ws

Else

    For Each ws In ThisWorkbook.Worksheets
        
        If InStr(1, ws.Name, "PXY", vbTextCompare) <> 0 Then
        
            Set ws_PXY = ThisWorkbook.Worksheets("PXY " & str_myName)
            '-- count parameter --
            dou_PXYParCounter = ws_PXY.Range("B" & Rows.Count).End(xlUp).Row - 2
            '-- TXY spreadsheet exist --
            bol_PXY = True
            dou_MWL = ws_PXY.Range("C1").Value
            dou_MWH = ws_PXY.Range("D1").Value
            
            Exit For
            
        End If
        
    Next ws
End If

myhandler:

End Sub
'---------------------------------------------------------------------------------------
' Function  : xlInterpolate
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Interpolate between given values
' Arguments :
'             dou_val1                  --> Higher value 1
'             dou_val2                  --> lower value 1
'             dou_val3                  --> Higher value 2
'             dou_val4                  --> lower value 2
'             dou_val                   --> Reference
'---------------------------------------------------------------------------------------
Private Function xlInterpolate(ByVal dou_val1 As Double, ByVal dou_val2 As Double, ByVal dou_val3 As Double, _
                                ByVal dou_val4 As Double, ByVal dou_val As Double) As Double

If dou_val1 <> dou_val2 Then

    xlInterpolate = dou_val3 - (dou_val3 - dou_val4) * (dou_val1 - dou_val) / (dou_val1 - dou_val2)

Else

    xlInterpolate = dou_val3
    
End If

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetTemp_TXY_y
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium temperature
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetTemp_TXY_y(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGetTemp_TXY_y = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 1
int_ResPointer = 0

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetTemp_TXY_y = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetTemp_TXY_y = -1002
    Exit Function

End If

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGetTemp_TXY_y = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGetTemp_TXY_y = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value

If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)
    
Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

xlGetTemp_TXY_y = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGetTemp_TXY_x
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium temperature
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Light component liquid composition
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetTemp_TXY_x(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGetTemp_TXY_x = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 2
int_ResPointer = 0

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetTemp_TXY_x = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetTemp_TXY_x = -1002
    Exit Function

End If

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGetTemp_TXY_x = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGetTemp_TXY_x = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

xlGetTemp_TXY_x = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetx_TXY_T
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Temperature
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetx_TXY_T(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGetx_TXY_T = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 0
int_ResPointer = 2

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGetx_TXY_T = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGetx_TXY_T = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- check if --
If dou_VAR2 > ws_TXY.Range("C" & lon_myRow).Value Then

    xlGetx_TXY_T = -1005
    Exit Function

End If

If Bol_Var1 = False Then

    If dou_VAR2 < ws_TXY.Range("C" & lon_myRow - 3).Offset(0, dou_Temp).Value Then

        xlGetx_TXY_T = -1006
        Exit Function
    
    End If
    
Else

    If dou_VAR2 < ws_TXY.Range("C" & lon_myRow).Offset(0, dou_Temp).Value Then

        xlGetx_TXY_T = -1006
        Exit Function
    
    End If

End If
'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value -- WARNING, FOR TEMPERATURE CHANGE THE SIGN
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGetx_TXY_T = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGetx_TXY_y
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetx_TXY_y(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGetx_TXY_y = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 1
int_ResPointer = 2

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetx_TXY_y = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetx_TXY_y = -1002
    Exit Function

End If

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGetx_TXY_y = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGetx_TXY_y = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value
'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGetx_TXY_y = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGety_TXY_T
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Temperature
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGety_TXY_T(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGety_TXY_T = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 0
int_ResPointer = 1

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGety_TXY_T = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGety_TXY_T = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- check if --
If dou_VAR2 > ws_TXY.Range("C" & lon_myRow).Value Then

    xlGety_TXY_T = -1005
    Exit Function

End If

If Bol_Var1 = False Then

    If dou_VAR2 < ws_TXY.Range("C" & lon_myRow - 3).Offset(0, dou_Temp).Value Then

        xlGety_TXY_T = -1006
        Exit Function
    
    End If
    
Else

    If dou_VAR2 < ws_TXY.Range("C" & lon_myRow).Offset(0, dou_Temp).Value Then

        xlGety_TXY_T = -1006
        Exit Function
    
    End If

End If
'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value -- WARNING, FOR TEMPERATURE CHANGE THE SIGN
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGety_TXY_T = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGety_TXY_x
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Pressure
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> TXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGety_TXY_x(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if TXY spreadsheet exist --
If bol_TXY = False Then

    xlGety_TXY_x = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 1
int_ResPointer = 2

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGety_TXY_x = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGety_TXY_x = -1002
    Exit Function

End If

'-- search for pressure --
'-- check first pressure --
If dou_VAR1 < ws_TXY.Range("A2").Value Then
    
    xlGety_TXY_x = -1003
    Exit Function
    
End If
'-- check last pressure --
If dou_VAR1 > ws_TXY.Range("A2").Offset(dou_TXYParCounter - 2, 0).Value Then
    
    xlGety_TXY_x = -1004
    Exit Function
    
End If

For i = 0 To dou_TXYParCounter
    
    '-- check if there is any pressure with the same value --
    If ws_TXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- pressure --
arr_Temp(0, 0) = ws_TXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_TXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_TXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value
'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 1), arr_Help(1, 1), arr_Help(0, 0), arr_Help(1, 0), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_TXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)
dou_Temp = (dou_Temp / dou_MWH) / ((dou_Temp / dou_MWH) + ((1 - dou_Temp) / dou_MWL))

xlGety_TXY_x = dou_Temp

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetPress_PXY_y
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium Temperature
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetPress_PXY_y(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGetPress_PXY_y = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 1
int_ResPointer = 0

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetPress_PXY_y = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetPress_PXY_y = -1002
    Exit Function

End If

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGetTemp_PXY_y = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGetTemp_PXY_y = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value

If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)
    
Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

xlGetPress_PXY_y = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGetPress_PXY_x
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium Temperature
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Light component liquid composition
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetPress_PXY_x(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGetPress_PXY_x = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 2
int_ResPointer = 0

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetPress_PXY_x = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetPress_PXY_x = -1002
    Exit Function

End If

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGetPress_PXY_x = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGetPress_PXY_x = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

xlGetPress_PXY_x = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetx_PXY_P
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Temperature
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetx_PXY_P(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGetx_PXY_P = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 0
int_ResPointer = 2

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGetx_PXY_P = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGetx_PXY_P = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- check if --
If dou_VAR2 > ws_PXY.Range("C" & lon_myRow).Value Then

    xlGetx_PXY_P = -1005
    Exit Function

End If

If Bol_Var1 = False Then

    If dou_VAR2 < ws_PXY.Range("C" & lon_myRow - 3).Offset(0, dou_Temp).Value Then

        xlGetx_PXY_P = -1006
        Exit Function
    
    End If
    
Else

    If dou_VAR2 < ws_PXY.Range("C" & lon_myRow).Offset(0, dou_Temp).Value Then

        xlGetx_PXY_P = -1006
        Exit Function
    
    End If

End If
'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value -- WARNING, FOR Temperature CHANGE THE SIGN
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGetx_PXY_P = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGetx_PXY_y
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGetx_PXY_y(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGetx_PXY_y = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 1
int_ResPointer = 2

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGetx_PXY_y = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGetx_PXY_y = -1002
    Exit Function

End If

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGetx_PXY_y = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGetx_PXY_y = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value
'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGetx_PXY_y = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))
End Function
'---------------------------------------------------------------------------------------
' Function  : xlGety_PXY_P
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Temperature
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGety_PXY_P(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGety_PXY_P = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for Xv, 2 for Xl or 0 for T
int_ArgPointer = 0
int_ResPointer = 1

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGety_PXY_P = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGety_PXY_P = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- check if --
If dou_VAR2 > ws_PXY.Range("C" & lon_myRow).Value Then

    xlGety_PXY_P = -1005
    Exit Function

End If

If Bol_Var1 = False Then

    If dou_VAR2 < ws_PXY.Range("C" & lon_myRow - 3).Offset(0, dou_Temp).Value Then

        xlGety_PXY_P = -1006
        Exit Function
    
    End If
    
Else

    If dou_VAR2 < ws_PXY.Range("C" & lon_myRow).Offset(0, dou_Temp).Value Then

        xlGety_PXY_P = -1006
        Exit Function
    
    End If

End If
'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value -- WARNING, FOR Temperature CHANGE THE SIGN
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value

'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value < dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGety_PXY_P = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function

'---------------------------------------------------------------------------------------
' Function  : xlGety_PXY_x
' DateTime  : 26/11/2014
' Author    : José García Herruzo
' Purpose   : Return equilibrium vapor composition
' Arguments :
'             dou_VAR1                 --> Temperature
'             dou_VAR2                 --> Light component vapor composition
'             str_wsName               --> PXY specific name
'---------------------------------------------------------------------------------------
Public Function xlGety_PXY_x(ByVal dou_VAR1 As Double, ByVal dou_VAR2 As Double, Optional ByVal str_wsName As String) As Double


Dim lon_myRow As Double
Dim lon_myColumn As Double

Dim arr_Help() As Double

Dim dou_Temp As Double

Dim int_ArgPointer As Integer
Dim int_ResPointer As Integer

Dim Bol_Var1 As Boolean
Dim Bol_Var2 As Boolean

Bol_Var1 = False
Bol_Var2 = False

If str_wsName = "" Then

    str_wsName = "Default SPD name"
    
End If

Call xsGeneral(str_wsName)

'-- check if PXY spreadsheet exist --
If bol_PXY = False Then

    xlGety_PXY_x = -1000
    Exit Function

End If

ReDim arr_Temp(1, 1)
ReDim arr_Help(1, 1)


'int_ArgPointer 1 for y, 2 for x or 0 for T
int_ArgPointer = 1
int_ResPointer = 2

' -- Convert composition --
dou_VAR2 = (dou_VAR2 / dou_MWL) / ((dou_VAR2 / dou_MWL) + ((1 - dou_VAR2) / dou_MWH))

'-- check if --
If dou_VAR2 < 0 Then

    xlGety_PXY_x = -1001
    Exit Function
    
ElseIf dou_VAR2 > 1 Then

    xlGety_PXY_x = -1002
    Exit Function

End If

'-- search for Temperature --
'-- check first Temperature --
If dou_VAR1 < ws_PXY.Range("A2").Value Then
    
    xlGety_PXY_x = -1003
    Exit Function
    
End If
'-- check last Temperature --
If dou_VAR1 > ws_PXY.Range("A2").Offset(dou_PXYParCounter - 2, 0).Value Then
    
    xlGety_PXY_x = -1004
    Exit Function
    
End If

For i = 0 To dou_PXYParCounter
    
    '-- check if there is any Temperature with the same value --
    If ws_PXY.Range("A2").Offset(i, 0).Value = dou_VAR1 Then
    
        lon_myRow = i + 2
        Bol_Var1 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("A2").Offset(i, 0).Value > dou_VAR1 Then
    
        lon_myRow = i + 2
        Exit For
        
    End If

Next i

'-- Temperature --
arr_Temp(0, 0) = ws_PXY.Range("A" & lon_myRow).Value
If Bol_Var1 = True Then

    arr_Temp(1, 0) = arr_Temp(0, 0)
    
Else

    arr_Temp(1, 0) = ws_PXY.Range("A" & lon_myRow - 3).Value

End If

'-- count the number of the point
dou_Temp = ws_PXY.Cells(2, Cells.Columns.Count).End(xlToLeft).Column - 3

For i = 0 To dou_Temp
    '-- check if there is any concentration with the same value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
    
        lon_myColumn = i
        Bol_Var2 = True
        Exit For
        
    End If
    '-- if not, check for the first higher value --
    If ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
    
        lon_myColumn = i
        Exit For
        
    End If

Next i

'-- Comp, higher --
arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn).Value
'-- Comp, lower --
arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow).Offset(int_ArgPointer, lon_myColumn - 1).Value
'-- Temp, higher --
arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn).Value
'-- Temp, lower --
arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow).Offset(int_ResPointer, lon_myColumn - 1).Value
'-- var 2 --
If Bol_Var2 = True Then

    arr_Temp(0, 1) = arr_Help(0, 1)

Else

    arr_Temp(0, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

End If

If Bol_Var1 = False Then

    '-- check second Temperature
    For i = 0 To dou_Temp
        '-- check if there is any concentration with the same value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value = dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
        '-- if not, check for the first higher value --
        If ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, i).Value > dou_VAR2 Then
        
            lon_myColumn = i
            Exit For
            
        End If
    
    Next i
    '-- Comp, higher --
    arr_Help(0, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn).Value
    '-- Comp, lower --
    arr_Help(1, 0) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ArgPointer, lon_myColumn - 1).Value
    '-- Temp, higher --
    arr_Help(0, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn).Value
    '-- Temp, lower --
    arr_Help(1, 1) = ws_PXY.Range("C" & lon_myRow - 3).Offset(int_ResPointer, lon_myColumn - 1).Value
    
    '-- Temperature --
    arr_Temp(1, 1) = xlInterpolate(arr_Help(0, 0), arr_Help(1, 0), arr_Help(0, 1), arr_Help(1, 1), dou_VAR2)

Else

    arr_Temp(1, 1) = arr_Temp(0, 1)
    
End If

dou_Temp = xlInterpolate(arr_Temp(0, 0), arr_Temp(1, 0), arr_Temp(0, 1), arr_Temp(1, 1), dou_VAR1)

' -- Convert composition --
xlGety_PXY_x = (dou_Temp * dou_MWL) / ((dou_Temp * dou_MWL) + ((1 - dou_Temp) * dou_MWH))

End Function


