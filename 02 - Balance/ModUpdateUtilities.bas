Attribute VB_Name = "ModUpdateUtilities"
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
' Module    : ModUpdateUtilities
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for update the utilities
' References: N/A
' Functions : N/A
' Procedures:
'               1-Import_Utility_Concumptions
'               2-Update_Utility_Requirements
'               3-Update_Chemicals_Requirements
'               4-Reset_Old_Chemicals_Values
'               5-Special_Steam_Procedure
'               6-xlRawWater
'               7-Update_Water_Requirements
'               8-xlRawWaterSpecialW2BSeville
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/01/2013  JGH     Update code into new kind of file nomenclature
'       07/02/2013  JGH     Floculant is added to chemical code
'       07/03/2013  JGH     Special steam code
'       07/03/2013  JGH     Floculant is eliminated
'       08/20/2013  JGH     Error when file are opened is solved
'       09/30/2013  JGH     xlRawWater is added
'       10/31/2013  JGH     xlRawWater is modified to sent an effluent to WWT in case of
'                           distillation stream was larger than consumption
'       30/12/2013 JGH      xlRawWaterSpecialW2BSeville is added
'----------------------------------------------------------------------------------------
Dim MBPath As String
Public int_Error_Pointer As Integer ' It controls the error
Dim lon_File_Counter As Long
Dim Source_File() As String
Dim myWB As Workbook
Dim myWorksheets As Worksheet
Dim lon_Stream_Counter As Long
Dim Stream_Name As Integer
Dim i As Long
Dim j As Integer
Dim k As Integer
Dim Stream_Data() As Variant '
Dim band_IsFileOpen As Boolean
'---------------------------------------------------------------------------------------
' Procedure : Import_Utility_Concumptions
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : Search each cooling water consumption
' Arguments :
'             str_Sheet_Name            --> Sheet were extracted streams will be
'                                           written
'             str_First_Range           --> Range were extracted streams will be
'                                           started to be written
'             var_Stream_Lower_limit    --> Number were utility starts to be called
'             var_Stream_Upper_limit    --> Number were utility ends to be called
'---------------------------------------------------------------------------------------
Public Sub Import_Utility_Concumptions(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                    ByVal var_Stream_Lower_limit As Integer, ByVal var_Stream_Upper_limit As Integer)

Dim str_help As String
Dim help() As String
Dim a As Integer

int_Error_Pointer = 0
lon_Stream_Counter = 0
band_IsFileOpen = False
ReDim Stream_Data(1000, lon_TOTAL_PAR)

'-- reset old lines --
If var_Utilities_Upper_Limit <> 1999 Then

    For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
    
        For k = 0 To lon_TOTAL_PAR
        
            ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(k, i).Value = ""
        
        Next k
    
    Next i

End If

'-- Update balance root path --
MBPath = WS_Setup.Range("C3").Value

If MBPath = "" Then

    int_Error_Pointer = 1
    Exit Sub

End If

'-- Update balances name --
lon_File_Counter = WS_Setup.Range("I16000").End(xlUp).Row

If lon_File_Counter <= 1 Then

    int_Error_Pointer = 2
    Exit Sub
    
End If

lon_File_Counter = lon_File_Counter - 2
ReDim Source_File(1, lon_File_Counter)

For i = 0 To lon_File_Counter

    Source_File(0, i) = WS_Setup.Range("I2").Offset(i, 0).Value
    Source_File(1, i) = WS_Setup.Range("I2").Offset(i, 1).Value

Next i

'-- Per each balance file --
For i = 0 To lon_File_Counter

    xlStartSettings ("Searching utilities streams in Area " & i + 1 & " to " & lon_File_Counter + 1)
    
    str_help = MBPath & "\" & Source_File(0, i) & "\" & Source_File(0, i) & ".02." & Source_File(1, i) & ".xls"
    
    If IsFileOpen(str_help) = True Then

        band_IsFileOpen = True
        help = Split(str_help, "\")
        a = UBound(help)
        Set myWB = Workbooks(help(a))
        
    Else
        
        band_IsFileOpen = False
        Call Update_Data_Into_ExcelCell(str_help, "Setup", "1", "Update Flag")
        Call Update_Data_Into_ExcelCell(str_help, "Setup", "1", "Msg Flag")
        Set myWB = Workbooks.Open(str_help)
    
    End If
    
    For Each myWorksheets In myWB.Worksheets
        
        '-- Search NT worksheets --
        If InStr(1, myWorksheets.Name, "-NT-", vbTextCompare) <> 0 Then
            '-- Per each stream --
            For j = 0 To ReturnColumn(myWB.Name, myWorksheets.Name, "D1") - 1
                
                If IsNumeric(myWorksheets.Range("D1").Offset(0, j).Value) = True Then
            
                    Stream_Name = myWorksheets.Range("D1").Offset(0, j).Value
                    '-- Search if the stream is cooling water --
                    If Stream_Name > var_Stream_Lower_limit And Stream_Name < var_Stream_Upper_limit Then
                        
                        For k = 0 To lon_TOTAL_PAR
                        
                            Stream_Data(lon_Stream_Counter, k) = myWorksheets.Range("D1").Offset(k, j).Value
                        
                        Next k
                        
                        lon_Stream_Counter = lon_Stream_Counter + 1
                        
                    End If
            
                End If
            
            Next j
            
        End If
    
    Next myWorksheets
    
    If band_IsFileOpen = False Then
    
        Application.DisplayAlerts = False
        myWB.Close
        Application.DisplayAlerts = True
        Call Update_Data_Into_ExcelCell(str_help, "Setup", "0", "Msg Flag")
        Call Update_Data_Into_ExcelCell(str_help, "Setup", "0", "Update Flag")
        
    End If
    
Next i

lon_Stream_Counter = lon_Stream_Counter - 1

'-- Cooling water streams are discharger over the cooling file --
If var_Utilities_Upper_Limit = 1999 Then

    Call Update_Chemicals_Requirements

Else

    For i = 0 To lon_Stream_Counter
    
        For k = 0 To lon_TOTAL_PAR
        
            ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(k, i).Value = Stream_Data(i, k)
            
        Next k
    
    Next i

End If

Set myWB = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Update_Utility_Requirements
' DateTime  : 05/08/2013
' Author    : José García Herruzo
' Purpose   : Calculate the total utility requirements
' Arguments :
'             str_Sheet_Name            --> Sheet where extracted streams will be
'                                           written
'             str_First_Range           --> Range where extracted streams will be
'                                           started to be written
'             var_mySimultaneityCoef    --> From Design Basis
'             bol_OnlySypply            --> True if Area only has a distribution lines.
'                                           False if Area has supply and return lines
'---------------------------------------------------------------------------------------
Public Sub Update_Utility_Requirements(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                        ByVal var_mySimultaneityCoef As Variant, bol_OnlySypply As Boolean)

Dim myValue As Variant
Dim myValue1 As Variant

'-- Update supply line --
For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
        
    myValue = myValue + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
    myValue1 = myValue1 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value

Next i

If bol_OnlySypply = False Then
    
    myValue = myValue / 2
    myValue1 = myValue1 / 2
    
End If

ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Value = myValue
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Value = myValue1 * var_mySimultaneityCoef

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Update_Chemicals_Requirements
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : write the extracted chemical streams
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Update_Chemicals_Requirements()

Dim int_Written_Lines As Integer
Dim i As Integer

Call Reset_Old_Chemicals_Values

'--PFD1--
int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1900 And Stream_Data(i, 0) <= 1915 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-001").Range("N1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i

'--PFD2--
int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1915 And Stream_Data(i, 0) <= 1930 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-002").Range("H1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i
   
'--PFD3--
int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1930 And Stream_Data(i, 0) <= 1945 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-003").Range("H1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i
    
 '--PFD4--
 int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1945 And Stream_Data(i, 0) <= 1960 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-004").Range("K1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i
    
 '--PFD5--
 int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1960 And Stream_Data(i, 0) <= 1975 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-005").Range("H1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i
    
'--PFD6--
int_Written_Lines = 0
    For i = 0 To lon_Stream_Counter
    
        If Stream_Data(i, 0) > 1975 And Stream_Data(i, 0) <= 1990 Then
        
            For k = 0 To lon_TOTAL_PAR
        
                ThisWorkbook.Worksheets("D-01900-NT-006").Range("f1").Offset(k, int_Written_Lines).Value = Stream_Data(i, k)
            
            Next k
            
            int_Written_Lines = int_Written_Lines + 1
            
        End If
        
    Next i
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : Reset_Old_Chemicals_Values
' DateTime  : 05/28/2013
' Author    : José García Herruzo
' Purpose   : write the extracted chemical streams
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Reset_Old_Chemicals_Values()

Dim int_Written_Lines As Integer
Dim i As Integer
Dim str_help As String


'--PFD1--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-001").Range("N1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-001", "N1")))
        
ThisWorkbook.Worksheets("D-01900-NT-001").Range("N1:" & str_help & "180").ClearContents

'--PFD2--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-002").Range("H1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-002", "H1")))
        
ThisWorkbook.Worksheets("D-01900-NT-002").Range("H1:" & str_help & "180").ClearContents
   
'--PFD3--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-003").Range("H1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-003", "H1")))
        
ThisWorkbook.Worksheets("D-01900-NT-003").Range("H1:" & str_help & "180").ClearContents
    
 '--PFD4--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-004").Range("K1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-004", "K1")))
        
ThisWorkbook.Worksheets("D-01900-NT-004").Range("K1:" & str_help & "180").ClearContents
    
 '--PFD5--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-005").Range("H1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-005", "H1")))
        
ThisWorkbook.Worksheets("D-01900-NT-005").Range("H1:" & str_help & "180").ClearContents
    
'--PFD6--
str_help = Get_Column_Letter(ThisWorkbook.Worksheets("D-01900-NT-006").Range("F1").Offset _
    (0, ReturnColumn(ThisWorkbook.Name, "D-01900-NT-006", "F1")))
        
ThisWorkbook.Worksheets("D-01900-NT-006").Range("F1:" & str_help & "180").ClearContents

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Special_Steam_Procedure
' DateTime  : 07/03/2013
' Author    : José García Herruzo
' Purpose   : Calculate the total utility requirements
' Arguments :
'             str_Sheet_Name            --> Sheet where extracted streams will be
'                                           written
'             str_First_Range           --> Range where extracted streams will be
'                                           started to be written
'             var_mySimultaneityCoef    --> From Design Basis
'---------------------------------------------------------------------------------------
Public Sub Special_Steam_Procedure(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                        ByVal var_mySimultaneityCoef As Variant)

Dim LP As Variant
Dim LPDesign As Variant

Dim HP As Variant
Dim HPDesign As Variant

'-- Update supply line --
For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
        
    If ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value > 5300 And _
        ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value < 5800 Then
    
        LP = LP + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
        LPDesign = LPDesign + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value
    
    ElseIf ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value > 5800 And _
        ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value < 5900 Then
    
        HP = HP + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
        HPDesign = HPDesign + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value
    
    End If
    
Next i

ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Value = LP / 2
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Value = (LPDesign / 2) * var_mySimultaneityCoef

ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 1).Value = HP
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Offset(0, 1).Value = HPDesign * var_mySimultaneityCoef

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlRawWater
' DateTime  : 09/30/2013
' Author    : José García Herruzo
' Purpose   : Calculate the total raw water requirement. It is a general code
' Arguments :
'             str_Sheet_Name            --> Sheet where extracted streams will be
'                                           written
'             str_First_Range           --> Range where extracted streams will be
'                                           started to be written
'             var_mySimultaneityCoef    --> From Design Basis
'             bol_OnlySypply            --> True if Area only has a distribution lines.
'                                           False if Area has supply and return lines
'---------------------------------------------------------------------------------------
Public Sub xlRawWater(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                        ByVal var_mySimultaneityCoef As Variant)

Dim myValue As Variant
Dim myValue1 As Variant

'-- Update supply line --
For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
        
    myValue = myValue + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
    myValue1 = myValue1 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value

Next i

If myValue - ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 1).Value < 0 Then

    ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Value = 0
    ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Value = 0
    
Else

    ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 7).Value = myValue - ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 1).Value
    ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Offset(0, 7).Value = (myValue1 - ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Offset(0, 1).Value) * var_mySimultaneityCoef

End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlRawWaterSpecialW2BSeville
' DateTime  : 30/12/2013
' Author    : José García Herruzo
' Purpose   : Calculate the total raw water requirement. It is applied to W2B Seville
' Arguments :
'             str_Sheet_Name            --> Sheet where extracted streams will be
'                                           written
'             str_First_Range           --> Range where extracted streams will be
'                                           started to be written
'             var_mySimultaneityCoef    --> From Design Basis
'             bol_OnlySypply            --> True if Area only has a distribution lines.
'                                           False if Area has supply and return lines
'---------------------------------------------------------------------------------------
Public Sub xlRawWaterSpecialW2BSeville(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                        ByVal var_mySimultaneityCoef As Variant)

Dim myValue As Variant
Dim myValue1 As Variant

'-- Update supply line --
For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
        
    '-- only scrubbers --
    If ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value <> "7353" Then
    
        myValue = myValue + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
        myValue1 = myValue1 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value
    
    End If
    
Next i

'-- Calculates total consumption --
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 7).Value = myValue
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Offset(0, 7).Value = myValue1 * var_mySimultaneityCoef

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Update_Water_Requirements
' DateTime  : 07/12/2013
' Author    : José García Herruzo
' Purpose   : Calculate the total water requirements and supplies
' Arguments :
'             str_Sheet_Name            --> Sheet where extracted streams will be
'                                           written
'             str_First_Range           --> Range where extracted streams will be
'                                           started to be written
'             var_mySimultaneityCoef    --> From Design Basis
'             bol_OnlySypply            --> True if Area only has a distribution lines.
'                                           False if Area has supply and return lines
'---------------------------------------------------------------------------------------
Public Sub Update_Water_Requirements(ByVal str_Sheet_Name As String, ByVal str_First_Range As String, _
                                        ByVal var_mySimultaneityCoef As Variant)

Dim myValue As Variant
Dim myValue1 As Variant
Dim Composition() As Variant

Dim myTemp As Variant

Dim myValue2 As Variant
Dim myValue3 As Variant

Dim j As Integer

ReDim Composition(110 - 12)

'-- Update supply line --
For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)
        
    If ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value >= 9250 Then
    
        myValue = myValue + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
        myValue1 = myValue1 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value
        myTemp = myTemp + (ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value * ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(8, i).Value)
        
        For j = 0 To 110 - 12
        
            Composition(j) = Composition(j) + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(11 + j, i).Value
        
        Next j
    
    End If
    
Next i

ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_TEMP).Value = myTemp / myValue
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Value = myValue
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Value = myValue1 * var_mySimultaneityCoef

For i = 0 To 110 - 12

    ThisWorkbook.Worksheets(str_Sheet_Name).Range("D12").Offset(i, 0).Value = Composition(i)

Next i

For i = 0 To ReturnColumn(ThisWorkbook.Name, str_Sheet_Name, str_First_Range)

    If ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(0, i).Value < 9250 Then
    
        myValue2 = myValue2 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(1, i).Value
        myValue3 = myValue3 + ThisWorkbook.Worksheets(str_Sheet_Name).Range(str_First_Range).Offset(5, i).Value
    
    End If

Next i

ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 2).Value = ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_ACF_MASS).Offset(0, 1).Value - myValue2
ThisWorkbook.Worksheets(str_Sheet_Name).Range(RA_DF_MASS).Offset(0, 1).Value = myValue3 * var_mySimultaneityCoef

End Sub


