Attribute VB_Name = "ModConvergence"
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
' Module    : ModConvergence
' DateTime  : 04/30/2013
' Author    : José García Herruzo
' Purpose   : This module contents procedures and functions for makes a convergence
' References: N/A
' Functions : N/A
' Procedures:
'               1-General_Convergence_Procedure
'               2-Distillation_Convergence
'               2-Stri_Mol_Sieve_Convergence
'               4-C18140_Convergence
'               5-Scrubber_Convergence
'               6-Water_Treatment_Convergence
'               7-Aerobic_Convergence
'               8-Anaerobic_Convergence
'               9-Scrubber_Convergence
'               10-PretII_Convergence
'               11-Fermentor_Convergence
'               12-Prop2_Convergence
'               13-Prop1_Convergence
'               14-PretIII_Convergence
'               15-Acid_Convergence
'               16-WaterToAcid_Convergence
' Updates   :
'       DATE        USER    DESCRIPTION
'       06/26/2013  JGH     Acid consumption by CaCO3 is added to Soaking procedure
'       06/27/2013  JGH     Fermentor. New tank volume file
'       07/08/2013  JGH     Distillation. General_Convergence_Procedure
'       07/16/2013  JGH     PretIII
'----------------------------------------------------------------------------------------
Public IterationCounter As Double
Public IterationCounterO As Double
Public ExtraIterationCounterO As Double
Public limit As Double
Public LimitO As Double
Public ExtraLimitO As Double

Public IterationCounter1 As Double
Public IterationCounter2 As Double
Public IterationCounter3 As Double
Public IterationCounter4 As Double

Public Dist1Error As Variant
Public Dist2Error As Variant
Public Dist3Error As Variant
Public Dist4Error As Variant

Public Tolerance As Variant

Dim help1 As Variant
Dim help2 As Variant
Dim help3 As Variant

Dim var_water1 As Variant
Dim var_water2 As Variant
Dim var_water_Error As Variant

Dim ws_Treatment As Worksheet

'---------------------------------------------------------------------------------------
' Procedure : General_Convergence_Procedure
' DateTime  : 06/25/2013
' Author    : José García Herruzo
' Purpose   : Load  Global Variables. Select the convergence procedure
' Arguments :
'             int_myPointer             --> Integer which indicates the convergence
'                                           process
'                                           * 1--> Distillation
'                                           * 2--> Water treatment
'                                           * 3--> Pretreatment II
'                                           * 4--> Fermentor
'                                           * 4--> Pretreatment III
'---------------------------------------------------------------------------------------
Public Sub General_Convergence_Procedure(ByVal int_myPointer As Integer)

Dist3Error = 100
Dist2Error = 100
Dist1Error = 100
Dist4Error = 100

limit = WS_Setup.Range("E2").Value
LimitO = WS_Setup.Range("D2").Value
ExtraLimitO = WS_Setup.Range("F2").Value

Tolerance = WS_Setup.Range("G2").Value

xlStartSettings ("Converging...")

If int_myPointer = 1 Then

    Call Distillation_Convergence
    Call Distillation_Convergence
    Call Distillation_Convergence
    
ElseIf int_myPointer = 2 Then

    Call Water_Treatment_Convergence
    
ElseIf int_myPointer = 3 Then

    Call PretII_Convergence
    
ElseIf int_myPointer = 4 Then

    Call Fermentor_Convergence
    
ElseIf int_myPointer = 5 Then

    Call PretIII_Convergence
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PretIII_Convergence
' DateTime  : 06/25/2013
' Author    : José García Herruzo
' Purpose   : Fermentor convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub PretIII_Convergence()

Dim ws As Worksheet
Dim ws4 As Worksheet
Dim i As Integer

Dim a() As Variant
Dim b() As Variant

Dim Solid_Error As Variant

Set ws = ThisWorkbook.Worksheets("D-12000-NT-001")
Set ws4 = ThisWorkbook.Worksheets("D-12000-NT-004")

Solid_Error = 0
IterationCounterO = 0

ReDim a(13)
ReDim b(13)

'--update values--
For i = 0 To 11

    a(i) = ws.Range(RA_CELLULOSE).Offset(i, 1).Value
    b(i) = ws4.Range(RA_CELLULOSE).Offset(i, 8).Value
    
Next i

    a(12) = ws.Range(RA_CACO3).Offset(i, 1).Value
    b(12) = ws4.Range(RA_CACO3).Offset(i, 8).Value
    
    a(13) = ws.Range(RA_GYPSUM).Offset(i, 1).Value
    b(13) = ws4.Range(RA_GYPSUM).Offset(i, 8).Value
    
For i = 0 To 13

    If a(i) = 0 And b(i) <> 0 Then
        
        a(i) = 1
        Solid_Error = Solid_Error + ((a(i) - b(i)) / a(i))
        a(i) = 0
       
    ElseIf a(i) <> 0 Then
    
    
        Solid_Error = Solid_Error + ((a(i) - b(i)) / a(i))
    
    End If

Next i

Solid_Error = Solid_Error / 14

Do Until Solid_Error = "0" Or Solid_Error = "0.00"
        
    For i = 0 To 11
    
        a(i) = b(i)
        ws.Range(RA_CELLULOSE).Offset(i, 1).Value = a(i)
        b(i) = ws4.Range(RA_CELLULOSE).Offset(i, 8).Value
        
    Next i
    
    a(12) = b(12)
    ws.Range(RA_CACO3).Offset(i, 1).Value = a(12)
    b(12) = ws4.Range(RA_CACO3).Offset(i, 8).Value
    
    a(13) = b(13)
    ws.Range(RA_GYPSUM).Offset(i, 1).Value = a(13)
    b(13) = ws4.Range(RA_GYPSUM).Offset(i, 8).Value
    
    For i = 0 To 13

        If a(i) = 0 And b(i) <> 0 Then
            
            a(i) = 1
            Solid_Error = Solid_Error + ((a(i) - b(i)) / a(i))
            a(i) = 0
           
        ElseIf a(i) <> 0 Then
        
        
            Solid_Error = Solid_Error + ((a(i) - b(i)) / a(i))
        
        End If

    Next i

    Call Acid_Convergence
    Call WaterToAcid_Convergence
    '--update values--
    Solid_Error = Solid_Error / 14
        
    IterationCounterO = IterationCounterO + 1
            
        If IterationCounterO = LimitO Then
            
            Exit Do
            
        End If
            
Loop

Set ws = Nothing
Set ws4 = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Acid_Convergence
' DateTime  : 07/16/2013
' Author    : José García Herruzo
' Purpose   : Acid procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Acid_Convergence()

Dim ws1 As Worksheet
Dim ws4 As Worksheet

Dim Current_Acid As Variant
Dim Desired_Acid As Variant

Dim Acid_Error As Variant
Dim AcidFlow As Variant

Set ws1 = ThisWorkbook.Worksheets("D-12000-NT-001")
Set ws4 = ThisWorkbook.Worksheets("D-12000-NT-004")

'--update values--
Current_Acid = ws4.Range(RA_ACID_SUL).Offset(0, 8).Value
Desired_Acid = ws1.Range(RA_ACID_SUL).Offset(0, 1).Value
AcidFlow = ws4.Range(RA_ACID_SUL).Offset(0, 3).Value

IterationCounter = 0

Acid_Error = (Desired_Acid - Current_Acid) / Desired_Acid

Do Until Acid_Error = "0" Or Acid_Error = "0.00"
        
    '--update values--
    AcidFlow = AcidFlow + AcidFlow * Acid_Error
    ws4.Range(RA_ACID_SUL).Offset(0, 3).Value = AcidFlow
    
    Current_Acid = ws4.Range(RA_ACID_SUL).Offset(0, 8).Value
    
    Acid_Error = (Desired_Acid - Current_Acid) / Desired_Acid
        
    IterationCounter = IterationCounter + 1
            
        If IterationCounter = limit Then
            
            Exit Do
            
        End If
            
Loop

Set ws1 = Nothing
Set ws4 = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : WaterToAcid_Convergence
' DateTime  : 06/25/2013
' Author    : José García Herruzo
' Purpose   : Soaking input water
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub WaterToAcid_Convergence()

Dim ws1 As Worksheet
Dim ws4 As Worksheet

Dim Current_Water As Variant
Dim Desired_Water As Variant

Dim Water_Error As Variant
Dim WaterFlow As Variant
Dim int_Counter As Variant

Set ws1 = ThisWorkbook.Worksheets("D-12000-NT-001")
Set ws4 = ThisWorkbook.Worksheets("D-12000-NT-004")

'--update values--
Current_Water = ws4.Range(RA_WATER).Offset(0, 8).Value
Desired_Water = ws1.Range(RA_WATER).Offset(0, 1).Value
WaterFlow = ws4.Range(RA_WATER).Offset(0, 2).Value

int_Counter = 0

Water_Error = (Desired_Water - Current_Water) / Desired_Water

Do Until Water_Error = "0" Or Water_Error = "0.00"
        
    '--update values--
    WaterFlow = WaterFlow + WaterFlow * Water_Error
    ws4.Range(RA_WATER).Offset(0, 2).Value = WaterFlow
    
    Current_Water = ws4.Range(RA_WATER).Offset(0, 8).Value

    Water_Error = (Desired_Water - Current_Water) / Desired_Water
        
    int_Counter = int_Counter + 1
            
        If int_Counter = limit Then
            
            Exit Do
            
        End If
            
Loop

Set ws1 = Nothing
Set ws4 = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Fermentor_Convergence
' DateTime  : 06/19/2013
' Author    : José García Herruzo
' Purpose   : Fermentor convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Fermentor_Convergence()

Dim ws As Worksheet

Dim Ferm_Error As Variant
Dim Prop_Error As Variant

Dim YeastFlow As Variant
Dim Ratio As Variant

Dim Desired_Ferm As Variant
Dim Current_Ferm As Variant
Dim Desired_Prop As Variant
Dim Current_Prop As Variant

Dim totalerror As Variant

Set ws = ThisWorkbook.Worksheets("Propagation&Fermentation")

'--update values--
totalerror = 1
Desired_Ferm = ws.Range("AD23").Value
IterationCounterO = 0
IterationCounter = 0

Current_Ferm = ws.Range("AD16").Value
YeastFlow = ws.Range(RA_ORGANISMS).Offset(0, 15).Value

Ferm_Error = (Desired_Ferm - Current_Ferm) / Desired_Ferm
Ferm_Error = Ferm_Error / 100

Do Until totalerror = "0" Or totalerror = "0.00"

    Do Until Ferm_Error < 0.000000001 And Ferm_Error > -0.000000001
        
        YeastFlow = ws.Range(RA_ORGANISMS).Offset(0, 15).Value
        YeastFlow = YeastFlow + YeastFlow * Ferm_Error
        ws.Range(RA_ORGANISMS).Offset(0, 15).Value = YeastFlow
        
        '--update values--
        Desired_Ferm = ws.Range("AD23").Value
        
        Current_Ferm = ws.Range("AD16").Value
        
        Ferm_Error = (Desired_Ferm - Current_Ferm) / Desired_Ferm
        Ferm_Error = Ferm_Error / 100
        
        IterationCounter = IterationCounter + 1
            
            If IterationCounter = limit Then
            
                Exit Do
            
            End If
            
    Loop
    
    Call Prop2_Convergence
    Call Prop1_Convergence
    Call Close_Convergence
    
    ThisWorkbook.Worksheets("Tanks-Vessels").Range("C6").Value = ThisWorkbook.Worksheets("Tanks-Vessels").Range("H10").Value
    ThisWorkbook.Worksheets("Tanks-Vessels").Range("C8").Value = ThisWorkbook.Worksheets("Tanks-Vessels").Range("I10").Value
    ThisWorkbook.Worksheets("Tanks-Vessels").Range("C10").Value = ThisWorkbook.Worksheets("Tanks-Vessels").Range("K10").Value
    
    IterationCounter = 0
    
    Desired_Ferm = ws.Range("AD23").Value
        
    Current_Ferm = ws.Range("AD16").Value
        
    Ferm_Error = (Desired_Ferm - Current_Ferm) / Desired_Ferm
    totalerror = Ferm_Error
    
            IterationCounterO = IterationCounterO + 1
            
            If IterationCounterO = LimitO Then
            
                Exit Do
            
            End If
            
Loop

Set ws = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Prop1_Convergence
' DateTime  : 06/19/2013
' Author    : José García Herruzo
' Purpose   : Propagator 2 convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Prop1_Convergence()

Dim ws As Worksheet

Dim Prop_Error As Variant

Dim YeastFlow As Variant

Dim Desired_Prop As Variant
Dim Current_Prop As Variant

Set ws = ThisWorkbook.Worksheets("Propagation&Fermentation")
        '--update values--
        Desired_Prop = ws.Range("AF23").Value
        
        Current_Prop = ws.Range("AF16").Value
        
        Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop
        Prop_Error = Prop_Error / 100
        
Do Until Prop_Error < 0.000000001 And Prop_Error > -0.000000001
        
        YeastFlow = ws.Range(RA_ORGANISMS).Offset(0, 7).Value
        YeastFlow = YeastFlow + YeastFlow * Prop_Error
        ws.Range(RA_ORGANISMS).Offset(0, 7).Value = YeastFlow
        
        '--update values--
        Desired_Prop = ws.Range("AF23").Value
        
        Current_Prop = ws.Range("AF16").Value
        
        Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop
        Prop_Error = Prop_Error / 100
        
        IterationCounter = IterationCounter + 1
            
            If IterationCounter = limit Then
            
                Exit Do
            
            End If
            
    Loop

Set ws = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Prop2_Convergence
' DateTime  : 06/19/2013
' Author    : José García Herruzo
' Purpose   : Propagator 2 convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Prop2_Convergence()

Dim ws As Worksheet

Dim Prop_Error As Variant

Dim Ratio As Variant

Dim Desired_Prop As Variant
Dim Current_Prop As Variant

Set ws = ThisWorkbook.Worksheets("Propagation&Fermentation")

'--update values--
Desired_Prop = ws.Range("Ae23").Value
IterationCounter = 0

Ratio = ws.Range("AC1").Value

Current_Prop = ws.Range("AE16").Value

Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop

Do Until Prop_Error < 0.000000001 And Prop_Error > -0.000000001

    Ratio = Ratio - Ratio * Prop_Error
    ws.Range("AC1").Value = Ratio
    
    '--update values--
    Desired_Prop = ws.Range("AE23").Value
    
    Current_Prop = ws.Range("AE16").Value
    
    Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop
    Ratio = ws.Range("AC1").Value
    
    IterationCounter = IterationCounter + 1
    
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
        
Loop

ws.Range(RA_ACF_MASS).Offset(0, 5).Value = ws.Range(RA_ACF_MASS).Offset(0, 7).Value / ws.Range(RA_AVG_DEN).Offset(0, 7).Value
ws.Range(RA_ACF_MASS).Offset(0, 5).Value = ws.Range(RA_ACF_MASS).Offset(0, 5).Value * ThisWorkbook.Worksheets("DB-16000").Range("B27").Value
ws.Range(RA_ACF_MASS).Offset(0, 5).Value = ws.Range(RA_ACF_MASS).Offset(0, 5).Value / ws.Range(RA_AVG_DEN).Offset(0, 5).Value

Set ws = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : Close_Convergence
' DateTime  : 06/19/2013
' Author    : José García Herruzo
' Purpose   : Propagator 2 convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Close_Convergence()

Dim ws As Worksheet

Dim Prop_Error As Variant

Dim Ratio As Variant

Dim Desired_Prop As Variant
Dim Current_Prop As Variant

Set ws = ThisWorkbook.Worksheets("Propagation&Fermentation")

'--update values--
Desired_Prop = ws.Range("AG23").Value
IterationCounter = 0

Ratio = ws.Range(RA_WATER).Offset(0, 4).Value

Current_Prop = ws.Range("Ag16").Value

Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop

Do Until Prop_Error < 0.000000001 And Prop_Error > -0.000000001

    Ratio = Ratio - Ratio * Prop_Error
    ws.Range(RA_WATER).Offset(0, 4).Value = Ratio
    
    '--update values--
    Desired_Prop = ws.Range("Ag23").Value
    
    Current_Prop = ws.Range("Ag16").Value
    
    Prop_Error = (Desired_Prop - Current_Prop) / Desired_Prop
    Ratio = ws.Range(RA_WATER).Offset(0, 4).Value
    
    IterationCounter = IterationCounter + 1
    
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
        
Loop

Set ws = Nothing

End Sub
'---------------------------------------------------------------------------------------
' Procedure : PretII_Convergence
' DateTime  : 06/17/2013
' Author    : Jose Manuel Tapia Jurado; Modified by José García Herruzo
' Purpose   : Pretreatment II convergence procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub PretII_Convergence()

Dim Hoja As Worksheet
Dim error2 As Variant
Dim error1 As Variant

Set Hoja = ThisWorkbook.Worksheets("Iteration")

Hoja.Range("c4:c22") = 0
Hoja.Range("e4:e22") = 0
Hoja.Range("g4:g22") = 0
Hoja.Range("i4:i22") = 0
Hoja.Range("k4:k22") = 0
Hoja.Range("m4:m22") = 0
Hoja.Range("o4:o22") = 0
Hoja.Range("q4:q22") = 0
Hoja.Range("s4:s22") = 0
Hoja.Range("u4:u22") = 0
Hoja.Range("w4:w22") = 0

'-- tomamos el error global de ambos bucles --

error1 = Hoja.Range("B24").Value
error2 = Hoja.Range("B25").Value

'Cálculos previos al do. Necesario mantener este orden para converger a valores reales.
IterationCounter = 0

Do Until error1 = "0" Or error1 = "0.00"
'1
Hoja.Range("b4:b22").Copy
Hoja.Range("c4:c22").PasteSpecial Paste:=xlValues

'2
Hoja.Range("d4:d22").Copy
Hoja.Range("e4:e22").PasteSpecial Paste:=xlValues

'3
Hoja.Range("f4:f22").Copy
Hoja.Range("g4:g22").PasteSpecial Paste:=xlValues

'4
Hoja.Range("h4:h22").Copy
Hoja.Range("i4:i22").PasteSpecial Paste:=xlValues

'5
Hoja.Range("j4:j22").Copy
Hoja.Range("k4:k22").PasteSpecial Paste:=xlValues

'6
Hoja.Range("l4:l22").Copy
Hoja.Range("m4:m22").PasteSpecial Paste:=xlValues

'11
Hoja.Range("v4:v22").Copy
Hoja.Range("w4:w22").PasteSpecial Paste:=xlValues

'-- update values--
error1 = Hoja.Range("B24").Value

        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
        
Loop

IterationCounter = 0

Do Until error2 = "0" Or error2 = "0.00"

'1
Hoja.Range("b4:b22").Copy
Hoja.Range("c4:c22").PasteSpecial Paste:=xlValues

'2
Hoja.Range("d4:d22").Copy
Hoja.Range("e4:e22").PasteSpecial Paste:=xlValues

'3
Hoja.Range("f4:f22").Copy
Hoja.Range("g4:g22").PasteSpecial Paste:=xlValues

'4
Hoja.Range("h4:h22").Copy
Hoja.Range("i4:i22").PasteSpecial Paste:=xlValues

'5
Hoja.Range("j4:j22").Copy
Hoja.Range("k4:k22").PasteSpecial Paste:=xlValues

'6
Hoja.Range("l4:l22").Copy
Hoja.Range("m4:m22").PasteSpecial Paste:=xlValues

'11
Hoja.Range("v4:v22").Copy
Hoja.Range("w4:w22").PasteSpecial Paste:=xlValues

'7
Hoja.Range("n4:n22").Copy
Hoja.Range("o4:o22").PasteSpecial Paste:=xlValues

'8
Hoja.Range("p4:p22").Copy
Hoja.Range("q4:q22").PasteSpecial Paste:=xlValues

'9
Hoja.Range("r4:r22").Copy
Hoja.Range("s4:s22").PasteSpecial Paste:=xlValues

'10
Hoja.Range("t4:t22").Copy
Hoja.Range("u4:u22").PasteSpecial Paste:=xlValues

'-- update values--
error2 = Hoja.Range("B25").Value

        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
Loop

Application.ScreenUpdating = True
Application.CutCopyMode = False

Set Hoja = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Water_Treatment_Convergence
' DateTime  : 04/09/2013
' Author    : José García Herruzo
' Purpose   : Launch Water treatment Convergence Procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Water_Treatment_Convergence()

Set ws_Treatment = ThisWorkbook.Worksheets("D-09000-NT-001")
    
Call Anaerobic_Convergence
Call Aerobic_Convergence
    
Set ws_Treatment = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Aerobic_Convergence
' DateTime  : 05/13/2013
' Author    : José García Herruzo
' Purpose   : Specific convergence procedure applied to aerobic reactor
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Aerobic_Convergence()

'-- Initializate var --
    IterationCounter = 0
    
    var_water1 = ws_Treatment.Range(RA_WATER).Offset(0, 17).Value
    var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 24).Value
    var_water1 = var_water1 + ws_Treatment.Range(RA_WATER).Offset(0, 23).Value
    var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 21).Value
    
    var_water2 = ws_Treatment.Range(RA_WATER).Offset(0, 18).Value

    var_water_Error = (var_water2 - var_water1) / var_water2
    
    '-- Iteration loop 1--
    Do Until var_water_Error = "0.00" Or var_water_Error = "0"
        
        '-- Calculate and write new value --
        var_water2 = var_water1
        ws_Treatment.Range(RA_WATER).Offset(0, 18).Value = var_water2
        
        '-- Update iteration value and error value --
        var_water1 = ws_Treatment.Range(RA_WATER).Offset(0, 17).Value
        var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 24).Value
        var_water1 = var_water1 + ws_Treatment.Range(RA_WATER).Offset(0, 23).Value
        var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 21).Value
        
        var_water2 = ws_Treatment.Range(RA_WATER).Offset(0, 18).Value

        var_water_Error = (var_water2 - var_water1) / var_water2
    
        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
    
    Loop
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Anaerobic_Convergence
' DateTime  : 05/13/2013
' Author    : José García Herruzo
' Purpose   : Specific convergence procedure applied to anaerobic reactor
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Anaerobic_Convergence()

'-- Initializate var --
    IterationCounter = 0
    
    var_water1 = ws_Treatment.Range(RA_WATER).Offset(0, 10).Value
    var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 13).Value
    
    var_water2 = ws_Treatment.Range(RA_WATER).Offset(0, 11).Value

    var_water_Error = (var_water2 - var_water1) / var_water2
    
    '-- Iteration loop 1--
    Do Until var_water_Error = "0.00" Or var_water_Error = "0"
        
        '-- Calculate and write new value --
        var_water2 = var_water1
        ws_Treatment.Range(RA_WATER).Offset(0, 11).Value = var_water2
        
        '-- Update iteration value and error value --
        var_water1 = ws_Treatment.Range(RA_WATER).Offset(0, 10).Value
        var_water1 = var_water1 - ws_Treatment.Range(RA_WATER).Offset(0, 13).Value
        
        var_water2 = ws_Treatment.Range(RA_WATER).Offset(0, 11).Value
    
        var_water_Error = (var_water2 - var_water1) / var_water2
    
        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
        
            Exit Do
        
        End If
    
    Loop
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Distillation_Convergence
' DateTime  : 04/09/2013
' Author    : José García Herruzo
' Purpose   : Launch Distillation Convergence Procedure
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Distillation_Convergence()

Dim CO21 As Variant
Dim CO22 As Variant
Dim CO2Result As Variant

Dim Dist2Error As Variant
Dim Ethanol As Variant
Dim WaterDist As Variant

Dim a As Variant
Dim b As Variant
Dim c As Variant
Dim d As Variant

Dim i As Integer

Dim SumA As Variant
Dim SumB As Variant

xlStartSettings ("Converging")

    '-- Initializate var --
    IterationCounterO = 0
    IterationCounter = 0
    IterationCounter1 = 0
    IterationCounter2 = 0
    IterationCounter3 = 0
    ExtraIterationCounterO = 0
    
'-- Iteration loop --

    Do Until Dist4Error = "0.00" Or Dist4Error = "0"
    
        a = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_OXYGEN).Offset(0, 3).Value
        b = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_NITROGEN).Offset(0, 3).Value
        c = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_CARB_DIO).Offset(0, 3).Value
        d = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ETHANOL).Offset(0, 3).Value
        SumA = a + b + c + d
        a = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_OXYGEN).Offset(0, 3).Value
        b = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_NITROGEN).Offset(0, 3).Value
        c = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_CARB_DIO).Offset(0, 3).Value
        d = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_ETHANOL).Offset(0, 3).Value
        SumB = a + b + c + d
        Dist1Error = (SumA - SumB) / SumA
        
        '-- Iteration loop --
        Do Until Dist1Error = "0.00" Or Dist1Error = "0"
            
            '-- Calculate and write new CO2 value --
            ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_OXYGEN).Offset(0, 3).Value = a
            ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_NITROGEN).Offset(0, 3).Value = b
            ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_CARB_DIO).Offset(0, 3).Value = c
            ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ETHANOL).Offset(0, 3).Value = d
        
            '-- Update iteration value and error value --
            a = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_OXYGEN).Offset(0, 3).Value
            b = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_NITROGEN).Offset(0, 3).Value
            c = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_CARB_DIO).Offset(0, 3).Value
            d = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ETHANOL).Offset(0, 3).Value
            SumA = a + b + c + d
            a = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_OXYGEN).Offset(0, 3).Value
            b = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_NITROGEN).Offset(0, 3).Value
            c = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_CARB_DIO).Offset(0, 3).Value
            d = ThisWorkbook.Worksheets("D-18000-NT-003").Range(RA_ETHANOL).Offset(0, 3).Value
        SumB = a + b + c + d
            SumB = a + b + c + d
            Dist1Error = (SumA - SumB) / SumA
        
            IterationCounter = IterationCounter + 1
            
            If IterationCounter = limit Then
            IterationCounter1 = IterationCounter1 + IterationCounter
                'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations")
            
                Exit Do
            
            End If
        
        Loop
        
        Do Until CheckingLimit(Dist2Error, Tolerance) = True And CheckingLimit(Dist3Error, Tolerance) = True
        
            Call C18140_Convergence
            IterationCounter2 = help1 + IterationCounter2
            Call Stri_Mol_Sieve_Convergence
            IterationCounter3 = help2 + IterationCounter3
            
            IterationCounterO = IterationCounterO + 1
            If IterationCounterO = LimitO Then
                'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations")
            
                Exit Do
            
            End If
            
        Loop
        
        Call Scrubber_Convergence
        IterationCounter4 = help3 + IterationCounter4
        
        ExtraIterationCounterO = ExtraIterationCounterO + 1
        If ExtraIterationCounterO = ExtraLimitO + 100 Then
                'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations"
            Exit Do
            
        End If
            
    Loop

End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : C18140_Convergence
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Specific convergence procedure applied to C-18140
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub C18140_Convergence()
    
    Dim BottomsWater As Variant
    Dim BottomsWaterResult As Variant
    Dim Water18 As Variant
    Dim water24 As Variant
    Dim Ethanol18 As Variant
    
    '-- Initializate var --
    IterationCounter = 0
    BottomsWater = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 4).Value
    Water18 = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 1).Value
    water24 = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 3).Value
    BottomsWaterResult = Water18 + water24
    Dist2Error = (BottomsWater - BottomsWaterResult) / BottomsWater
    Ethanol18 = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_ETHANOL).Offset(0, 1).Value
    '-- Iteration loop --
    Do Until Dist2Error = "0.00" Or Dist2Error = "0" ' Or CheckingLimit(Dist2Error) = True
        
        '-- Calculate and write new Ethanol value value --
        Ethanol18 = Ethanol18 + Ethanol18 * Dist2Error
        ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_ETHANOL).Offset(0, 1).Value = Ethanol18
        
        '-- Update iteration value and error value --
    BottomsWater = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 4).Value
    Water18 = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 1).Value
    water24 = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 3).Value
        BottomsWaterResult = water24 - Water18
        Dist2Error = (BottomsWater - BottomsWaterResult) / BottomsWater
         
        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
            'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations")
        
            Exit Do
        
        End If
    
    Loop
    
    help1 = IterationCounter
    
End Sub
    
'---------------------------------------------------------------------------------------
' Procedure : Stri_Mol_Sieve_Convergence
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Specific convergence procedure applied between stripping column and
'             molecular sieve PFDs
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Stri_Mol_Sieve_Convergence()
    
    Dim WaterA As Variant
    Dim EthanolA As Variant
    Dim TotalA As Variant
    Dim TotalB As Variant
    
    '-- Initializate var --
    IterationCounter = 0
    TotalA = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ACF_MASS).Offset(0, 11).Value
    TotalB = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_ACF_MASS).Offset(0, 8).Value
    Dist3Error = (TotalA - TotalB) / TotalA
    
    '-- Iteration loop --
    Do Until Dist3Error = "0.00" Or Dist3Error = "0" 'CheckingLimit(Dist3Error) = True
        
        '-- Calculate and write new Ethanol value value --
        WaterA = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_WATER).Offset(0, 11).Value
        EthanolA = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ETHANOL).Offset(0, 11).Value
        ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_WATER).Offset(0, 8).Value = WaterA
        ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_ETHANOL).Offset(0, 8).Value = EthanolA
         
        '-- Update iteration value and error value --
        TotalA = ThisWorkbook.Worksheets("D-18000-NT-002").Range(RA_ACF_MASS).Offset(0, 11).Value
        TotalB = ThisWorkbook.Worksheets("D-18000-NT-004").Range(RA_ACF_MASS).Offset(0, 8).Value
        Dist3Error = (TotalA - TotalB) / TotalA
         
        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
    
            'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations")
        
            Exit Do
        
        End If
    
    Loop
    
    help2 = IterationCounter

End Sub

'---------------------------------------------------------------------------------------
' Procedure : Scrubber_Convergence
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Specific convergence procedure applied between beer well and final
'             scrubber
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub Scrubber_Convergence()
    
Dim TotalA As Variant
Dim TotalB As Variant
Dim Stream18007() As Variant
Dim i As Integer

    '-- Initializate var --
    ReDim Stream18007(170)
    
    IterationCounter = 0
    For i = 0 To 170
    
        Stream18007(i) = ThisWorkbook.Worksheets("D-18000-NT-006").Range("J12").Offset(i, 0).Value
    
    Next i
    
    TotalA = ThisWorkbook.Worksheets("D-18000-NT-001").Range(RA_ACF_MASS).Offset(0, 3).Value
    TotalB = ThisWorkbook.Worksheets("D-18000-NT-006").Range(RA_ACF_MASS).Offset(0, 6).Value
    Dist4Error = (TotalA - TotalB) / TotalA
    
    '-- Iteration loop --
    Do Until Dist4Error = "0.00" Or Dist4Error = "0" 'CheckingLimit(Dist3Error) = True
        
        '-- Calculate and write new Ethanol value value --
        For i = 0 To 170
        
            Stream18007(i) = ThisWorkbook.Worksheets("D-18000-NT-006").Range("J12").Offset(i, 0).Value
        
        Next i
    
        For i = 0 To 170
        
            ThisWorkbook.Worksheets("D-18000-NT-001").Range("G12").Offset(i, 0).Value = Stream18007(i)
        
        Next i
         
        '-- Update iteration value and error value --
        TotalA = ThisWorkbook.Worksheets("D-18000-NT-001").Range(RA_ACF_MASS).Offset(0, 3).Value
        TotalB = ThisWorkbook.Worksheets("D-18000-NT-006").Range(RA_ACF_MASS).Offset(0, 6).Value
        Dist4Error = (TotalA - TotalB) / TotalA
         
        IterationCounter = IterationCounter + 1
        
        If IterationCounter = limit Then
    
            'MsgBox ("Solution have not been found in  " & IterationCounter & " iterations")
        
            Exit Do
        
        End If
    
    Loop
    
    help3 = IterationCounter

End Sub
