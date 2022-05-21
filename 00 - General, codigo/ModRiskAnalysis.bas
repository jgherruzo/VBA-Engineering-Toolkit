Attribute VB_Name = "ModRiskAnalysis"
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
' Module    : ModRiskAnalysis
' DateTime  : 13/06/2018
' Author    : José García Herruzo
' Purpose   : This module contents procedures and function in order to build up a risk
'               analysis
' References: ModEnhacements
' Functions : N/A
' Procedures:
'               1-xpGetClasification
'               1-xpGetLikelyHood
'               1-xpGetConsecuences
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : xpGetClasification
' DateTime  : 13/06/2018
' Author    : José García Herruzo
' Purpose   : Paint cell depending on the value
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub xpGetClasification()

Dim lon_row As Integer
Dim ra_myRange As Range
Dim int_Value As Integer
Dim i As Long

Set ra_myRange = ThisWorkbook.Worksheets("Analysis").Range("ra_Cla")

lon_row = ra_myRange.Offset(16000, 0).End(xlUp).Row - ra_myRange.Row

For i = 1 To lon_row

    int_Value = ra_myRange.Offset(i, 0).Value
    
    If int_Value <= 3 Then
    
        Call PaintGreen(ra_myRange.Offset(i, 0))
        
    ElseIf int_Value > 3 And int_Value < 8 Then
    
        Call PaintYellow(ra_myRange.Offset(i, 0))
    
    ElseIf int_Value >= 8 And int_Value < 15 Then
    
        Call PaintOrange(ra_myRange.Offset(i, 0))
    
    Else
    
        Call PaintRed(ra_myRange.Offset(i, 0))
        
    End If
    
Next i


End Sub
'---------------------------------------------------------------------------------------
' Procedure : xpGetLikelyHood
' DateTime  : 13/06/2018
' Author    : José García Herruzo
' Purpose   : Paint cell depending on the value
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub xpGetLikelyHood()

Dim lon_row As Integer
Dim ra_myRange As Range
Dim int_Value As Integer
Dim i As Long

Set ra_myRange = ThisWorkbook.Worksheets("Analysis").Range("ra_lik")

lon_row = ra_myRange.Offset(16000, 0).End(xlUp).Row - ra_myRange.Row

For i = 1 To lon_row

    int_Value = ra_myRange.Offset(i, 0).Value
    
    Select Case int_Value
        
        Case 1
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(221, 235, 247)
            
        Case 2
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(189, 215, 238)
        
        Case 3
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(155, 194, 230)
        
        Case 4
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(47, 117, 181)
            
        Case 5
            
            ra_myRange.Offset(i, 0).Interior.Color = RGB(31, 78, 120)
            
    End Select
    
Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xpGetConsecuences
' DateTime  : 13/06/2018
' Author    : José García Herruzo
' Purpose   : Paint cell depending on the value
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub xpGetConsecuences()

Dim lon_row As Integer
Dim ra_myRange As Range
Dim int_Value As Integer
Dim i As Long

Set ra_myRange = ThisWorkbook.Worksheets("Analysis").Range("ra_Con")

lon_row = ra_myRange.Offset(16000, 0).End(xlUp).Row - ra_myRange.Row

For i = 1 To lon_row

    int_Value = ra_myRange.Offset(i, 0).Value
    
    Select Case int_Value
        
        Case 1
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(217, 225, 242)
            
        Case 2
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(180, 198, 231)
        
        Case 3
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(142, 169, 219)
        
        Case 4
        
            ra_myRange.Offset(i, 0).Interior.Color = RGB(48, 84, 150)
            
        Case 5
            
            ra_myRange.Offset(i, 0).Interior.Color = RGB(32, 55, 100)
            
    End Select
    
Next i


End Sub
