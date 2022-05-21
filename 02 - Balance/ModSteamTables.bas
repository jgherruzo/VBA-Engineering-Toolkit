Attribute VB_Name = "ModSteamTables"
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
' Module    : ModSteamTables
' DateTime  : 09/24/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for extract steam tables value
' References: N/A
' Functions :
'               1-GetLanda
'               2-GetPhase
' Procedures:
'               1-InitializateSheets
'               2-SaturatedConditionbyPressure
'               3-CloseSheets
'               4-ExtractConditions
' Updates   :
'       DATE        USER    DESCRIPTION
'       10/15/2013  JGH     GetPhase is added
'       10/15/2013  JGH     Add boolean argument to the procedures string in order to
'                           only delete Tables worksheet at the last parameter
'       26/02/2014  JGH     3-ExtractConditions is added
'----------------------------------------------------------------------------------------
'****************************************************************************************
'*  MANUAL:                                                                             *
'*  if you need a steam condition you only have to call SaturatedConditionbyPressure    *
'*  and SaturatedInfo() will be completed. If you need landa you only have to call it   *
'*  Getphase return 0,1 or 0.5 if the stream is inside the equilibrium                  *
'****************************************************************************************
Dim ws_SteamTables As Worksheet
Dim Saturated1() As Variant
Dim Saturated2() As Variant
Public SaturatedInfo() As Variant
'---------------------------------------------------------------------------------------
' Procedure : InitializateSheets
' DateTime  : 09/24/2013
' Author    : José García Herruzo
' Purpose   : This procedure creates sheet, import it from original source and set the
'               names of the sheets
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub InitializateSheets()

Dim str_Path As String
Call Initializate_Var
'-- Check if sheets exist, if not, create it --
If SheetExist(ThisWorkbook.Name, "Tables") = False Then

    Set ws_SteamTables = ThisWorkbook.Worksheets.Add
    ws_SteamTables.Name = "Tables"
    
Else
    
    Set ws_SteamTables = ThisWorkbook.Worksheets("Tables")
    Exit Sub
    
    ws_SteamTables.Range("A1:AT16000").ClearContents
    
End If

'-- Data are extracted --
str_Path = WS_Setup.Range("K2").Value
str_Path = str_Path & "\SteamTables.xls"
    
Call Extract_Data_From_Excel_WithTittle(str_Path, ws_SteamTables.Name, "A1:AT16000", ws_SteamTables.Name, "A2")
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CloseSheets
' DateTime  : 09/24/2013
' Author    : José García Herruzo
' Purpose   : Delete sheet and close the variable
' Arguments :
'             bol_Last                 --> True if it is the last parameter and delete
'                                           the tables worksheet
'---------------------------------------------------------------------------------------
Private Sub CloseSheets(ByVal bol_Last As Boolean)

Application.DisplayAlerts = False

If bol_Last = True Then

    ws_SteamTables.Delete
    
End If

Set ws_SteamTables = Nothing
Application.DisplayAlerts = True

End Sub
'---------------------------------------------------------------------------------------
' Function  : GetLanda
' DateTime  : 09/24/2013
' Author    : José García Herruzo
' Purpose   : Return landa at pressure
' Arguments :
'             myPressure               --> Pressure
'             bol_Last                 --> True if it is the last parameter and delete
'                                           the tables worksheet
'---------------------------------------------------------------------------------------
Public Function GetLanda(ByVal myPressure As Variant, ByVal bol_Last As Boolean) As Variant

Call SaturatedConditionbyPressure(myPressure, bol_Last)

GetLanda = SaturatedInfo(3, 1) - SaturatedInfo(3, 0)

End Function
'---------------------------------------------------------------------------------------
' Function  : GetPhase
' DateTime  : 10/15/2013
' Author    : José García Herruzo
' Purpose   : Return stream phase. Return 0.5 to biphase conditions
' Arguments :
'             myPressure               --> Pressure
'             myTemperature            --> Temperature
'             bol_Last                 --> True if it is the last parameter and delete
'                                           the tables worksheet
'---------------------------------------------------------------------------------------
Public Function GetPhase(ByVal myPressure As Variant, ByVal myTemperature As Variant, ByVal bol_Last As Boolean) As Variant

'-- Fill conditions array --
Call SaturatedConditionbyPressure(myPressure, bol_Last)

' -- if temperature is bigger, is a vapor stream --
If myTemperature > SaturatedInfo(0, 1) Then

    GetPhase = 1
    
' -- if temperature is smaller, is a liquid stream --
ElseIf myTemperature < SaturatedInfo(0, 1) Then

    GetPhase = 0

' -- else, is inside the equilibrium --
Else

    GetPhase = 0.5

End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : SaturatedConditionbyPressure
' DateTime  : 09/24/2013
' Author    : José García Herruzo
' Purpose   : Calculates steam saturated properties
' Arguments :
'             myPressure               --> Pressure
'             bol_Last                 --> True if it is the last parameter and delete
'                                           the tables worksheet
'---------------------------------------------------------------------------------------
Public Sub SaturatedConditionbyPressure(ByVal myPressure As Variant, ByVal bol_Last As Boolean)

Dim lon_Counter As Long
Dim i As Long
Dim j As Long
Dim lon_Offset As Long

Call InitializateSheets

'-- convert pressure from bara to kPa --
myPressure = myPressure * 100

'-- Check the limit --
If myPressure > 60000 Or myPressure < 1 Then
    Call CloseSheets(bol_Last)
    Exit Sub

End If

'-- Count the las pressure --
lon_Counter = ws_SteamTables.Range("A16000").End(xlUp).Row - 3
'-- Search the first value which is bigger than pressure --
For i = 0 To lon_Counter

    If ws_SteamTables.Range("A3").Offset(i, 0).Value > myPressure Then
        
        lon_Offset = i
        Exit For
        
    End If

Next i
'-- Dim the matrix --
    '-- First for liquid. 0,0 -> Pressure, 0,1-> Temperature
ReDim Saturated1(4, 1)
ReDim Saturated2(4, 1)

ReDim SaturatedInfo(4, 1)

Saturated2(0, 0) = ws_SteamTables.Range("A3").Offset(lon_Offset, 0).Value
Saturated2(0, 1) = ws_SteamTables.Range("A3").Offset(lon_Offset, 4).Value

Saturated1(0, 0) = ws_SteamTables.Range("A3").Offset(lon_Offset - 4, 0).Value
Saturated1(0, 1) = ws_SteamTables.Range("A3").Offset(lon_Offset - 4, 4).Value

For i = 1 To 4

    For j = 0 To 1
    
        Saturated2(i, j) = ws_SteamTables.Range("A3").Offset(lon_Offset + i - 1, 2 + j).Value
        Saturated1(i, j) = ws_SteamTables.Range("A3").Offset(lon_Offset + i - 1 - 4, 2 + j).Value
        
    Next j

Next i

'-- Calculate --
SaturatedInfo(0, 0) = myPressure
'-- Temperature --
SaturatedInfo(0, 1) = Saturated1(0, 1) - (((Saturated1(0, 0) - myPressure) / _
                (Saturated1(0, 0) - Saturated2(0, 0))) * (Saturated1(0, 1) - Saturated2(0, 1)))

For i = 1 To 4
    
    For j = 0 To 1
    
        SaturatedInfo(i, j) = Saturated1(i, j) - (((Saturated1(0, 0) - myPressure) / _
                        (Saturated1(0, 0) - Saturated2(0, 0))) * (Saturated1(i, j) - Saturated2(i, j)))

    Next j
    
Next i

Call CloseSheets(bol_Last)

End Sub
'---------------------------------------------------------------------------------------
' Procedure : ExtractConditions
' DateTime  : 08/02/2014
' Author    : José García Herruzo
' Purpose   : Read pressures and download the properties
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub ExtractConditions()

Dim PInfo() As Variant
Dim limit As Integer
Dim ws As Worksheet
Dim i As Integer
Dim j As Integer

Set ws = ThisWorkbook.Worksheets("Steam Properties")

limit = ws.Range("HH1").End(xlToLeft).Column - 2

If limit < 0 Then

    Exit Sub

End If

ReDim PInfo(limit, 3)

For i = 0 To UBound(PInfo)

    PInfo(i, 0) = ws.Range("B1").Offset(0, i).Value

Next i

For i = 0 To UBound(PInfo)

    If i = UBound(PInfo) Then
    
        Call SaturatedConditionbyPressure(PInfo(i, 0), True)
        
    Else
    
        Call SaturatedConditionbyPressure(PInfo(i, 0), False)
    
    End If

    PInfo(i, 1) = SaturatedInfo(0, 1)
    PInfo(i, 2) = SaturatedInfo(3, 0)
    PInfo(i, 3) = SaturatedInfo(3, 1)
    
Next i

For i = 0 To UBound(PInfo)

    For j = 0 To 3
    
        ws.Range("B2").Offset(j, i).Value = PInfo(i, j)
    
    Next j
    
Next i

Set ws = Nothing

End Sub



