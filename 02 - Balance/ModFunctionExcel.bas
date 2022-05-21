Attribute VB_Name = "ModFunctionExcel"
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
' DateTime  : 04/08/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : This module contents funcions for calculating some balance parameters
'
' References: N/A
' Functions :
'               1-TSS
'               2-TIS
' Procedures:
'               1-Update_Parameters
'Updates   :
'       DATE        USER    DESCRIPTION
'       04/25/2014  JESG    Function update
'       05/26/2014  JESG    Added cell "S2" in setup to update formula
'----------------------------------------------------------------------------------------

'-- Design variables --
Dim SolubleSolid As String
Dim InsolubleSolid As String

'---------------------------------------------------------------------------------------
' Procedure : Update_Parameters
' DateTime  : 04/08/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Update component to calcute with function
' Arguments : N/A
'---------------------------------------------------------------------------------------

Private Sub Update_Parameters()

SolubleSolid = "ss"
InsolubleSolid = "is"

End Sub

'---------------------------------------------------------------------------------------
' Procedure : TSS
' DateTime  : 04/08/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Calculate Total Soluble Solids in a stream.
' Arguments :
'               rg_Stream ----> Stream range. Include all cells involve in TSS calculation
'---------------------------------------------------------------------------------------
Public Function TSS(ByVal rg_Stream As Range, Optional rg_ParameterType As Range, Optional flag As Range) As Variant

Dim Column As Long
Dim TotalFlow As Variant
Dim LastRow As Long
Dim Row As Long
Dim Param() As Variant
Dim ComponentType As String
Dim ComponetAdded As Variant
Dim ComponentFlow As Variant
Dim rg_FirstParameter As Range
Dim i As Integer

If ThisWorkbook.Worksheets("Setup").Range("S2").Value = 0 Then

Call Update_Parameters

ComponentType = SolubleSolid
Column = rg_Stream.Column
Set rg_FirstParameter = rg_Stream.Offset(0, -Column + 2)

LastRow = rg_FirstParameter.Offset(1000, 0).End(xlUp).Row - 1

'-- Load parameters value in an matrix --
ReDim Param(LastRow)

    For i = 1 To LastRow
    Param(i) = rg_FirstParameter.Offset(i - 1, 0).Value
    Next i


    '-- If balance is MEB --
    If Left(rg_Stream.Formula, 1) = "=" Then
    
        For i = 1 To LastRow
                
            If InStr(1, Param(i), ComponentType, vbTextCompare) <> 0 And InStr(1, Param(i), "asp", vbTextCompare) = 0 Then
            ComponetAdded = rg_Stream.Offset(i - 1, 0).Value
                
                If IsNumeric(ComponetAdded) = True Then
                ComponentFlow = ComponentFlow + ComponetAdded
                End If
            
            End If
        
        Next i
    
    '-- If balance is AMEB --
    Else
    
        For i = 1 To LastRow
        
            If InStr(1, Param(i), ComponentType, vbTextCompare) <> 0 And InStr(1, Param(i), "exc", vbTextCompare) = 0 Then
            ComponetAdded = rg_Stream.Offset(i - 1, 0).Value
                
                If IsNumeric(ComponetAdded) = True Then
                ComponentFlow = ComponentFlow + ComponetAdded
                End If
            
            End If
        
        Next i
        
    End If
    
TotalFlow = rg_Stream.Value

    If IsNumeric(TotalFlow) = True And TotalFlow > 0 Then
    TSS = ComponentFlow / TotalFlow
    Else
    TSS = 0
    End If

End If

End Function

'---------------------------------------------------------------------------------------
' Procedure : TIS
' DateTime  : 04/08/2014
' Author    : José Enrique Sarmiento Garrido
' Purpose   : Calculate Total Insoluble Solids in a stream.
' Arguments :
'               rg_Stream ----> Stream range. Include all cells involve in TIS calculation
'---------------------------------------------------------------------------------------

Public Function TIS(ByVal rg_Stream As Range, Optional rg_ParameterType As Range, Optional flag As Range) As Variant

Dim Column As Long
Dim TotalFlow As Variant
Dim LastRow As Long
Dim Row As Long
Dim Param() As Variant
Dim ComponentType As String
Dim ComponetAdded As Variant
Dim ComponentFlow As Variant
Dim rg_FirstParameter As Range
Dim i As Integer


If ThisWorkbook.Worksheets("Setup").Range("S2").Value = 0 Then

Call Update_Parameters

ComponentType = InsolubleSolid
Column = rg_Stream.Column
Set rg_FirstParameter = rg_Stream.Offset(0, -Column + 2)


LastRow = rg_FirstParameter.Offset(1000, 0).End(xlUp).Row - 1

'-- Load parameters value in an matrix --
ReDim Param(LastRow)

    For i = 1 To LastRow
    Param(i) = rg_FirstParameter.Offset(i - 1, 0).Value
    Next i


    '-- If balance is MEB --
    If Left(rg_Stream.Formula, 1) = "=" Then
    
        For i = 1 To LastRow
                
            If InStr(1, Param(i), ComponentType, vbTextCompare) <> 0 And InStr(1, Param(i), "asp", vbTextCompare) = 0 Then
            ComponetAdded = rg_Stream.Offset(i - 1, 0).Value
                
                If IsNumeric(ComponetAdded) = True Then
                ComponentFlow = ComponentFlow + ComponetAdded
                End If
            
            End If
        
        Next i
    
    '-- If balance is AMEB --
    Else
    
        For i = 1 To LastRow
        
            If InStr(1, Param(i), ComponentType, vbTextCompare) <> 0 And InStr(1, Param(i), "exc", vbTextCompare) = 0 Then
            ComponetAdded = rg_Stream.Offset(i - 1, 0).Value
                
                If IsNumeric(ComponetAdded) = True Then
                ComponentFlow = ComponentFlow + ComponetAdded
                End If
            
            End If
        
        Next i
        
    End If

TotalFlow = rg_Stream.Value

    If IsNumeric(TotalFlow) = True And TotalFlow > 0 Then
    TIS = ComponentFlow / TotalFlow
    Else
    TIS = 0
    End If
    
End If

End Function








