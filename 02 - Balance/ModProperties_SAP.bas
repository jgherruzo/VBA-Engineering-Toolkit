Attribute VB_Name = "ModProperties_SAP"
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
' Module      : ModProperties_SAP
' DateTime    : 11/10/2017
' Author      : José García Herruzo
' Purpose     : This module contents properties data and calculation specific for SAP
' References  : ModProperties_ThermoChemical
' Requirements: N/A
' Functions   :
'               01-xfTEquilSO2_SO3_PSO2O2X
'               02-Kp_X
'               03-Equil_Ref1_Kp
'               04-Equil_Ref2_Kp
'               05-Log10
'               06-Equil_Ref3_Kp
'               07-Equil_Ref4_Kp
'               08-Equil_Ref5_Kp
'               09-Equil_Ref6_Kp
'               10-Equil_Ref7_Kp
'               11-Equil_Ref7_T
'               12-X_Kp
'               13-xfXEquilSO2_SO3_PSO2O2T
'               14-xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key
'               15-xfBedX_CO2SO2O2N2SO3TinTout
'               16-xfBedT_CO2SO2O2N2SO3TinConv
'               17-xfXEquilAIATSO2_SO3_PSO2O2TPaSO2aO2a
'               19-xfTEquilAIATSO2_SO3_PSO2O2ConvPaSO2aO2a
'               18-xfTXEquilaIATSO2_SO3_PSO2iO2iCO2SO2O2N2SO3TiPaSO2aO2a_Key
'               20-xfCpDil_T_Conc
'               21-xfRhoDil_T_Conc
' Procedures  : N/A
' Updates     :
'       DATE        USER    DESCRIPTION
'       16/11/2017  JGH     Name of function 1 is modified to add the arguments
'       11/12/2017  JGH     Function 01 is modified to work with diferent equations
'       11/12/2017  JGH     Functions 02 to 10 are added to obtain the equilirbium Temp.
'       13/12/2017  JGH     Functions 11 to 14 are added to obtain the equilibrium comp.
'       14/12/2017  JGH     xfBedX_CO2SO2O2N2SO3TinTout and xfBedT_CO2SO2O2N2SO3TinConv
'                           are added
'       19/12/2017  JGH     xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key is modified to
'                           make heat balance using bed conversion instead of global
'       19/12/2017  JGH     Functions 17 to 19 are added to work after IAT
'       19/12/2017  JGH     Function 15 and 16 are modified to improve their accuracy for
'                           last bed
'       26/07/2019  JGH     Functions 20 to 21 are added
'-----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function  : xfTEquilSO2_SO3_PSO2O2X
' DateTime  : 11/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in ºC
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function xfTEquilSO2_SO3_PSO2O2X(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Conv As Double) As Variant

Dim dbl_Kp_X As Double
Dim dbl_Temp As Double
Dim int_Sel As Integer

dbl_Kp_X = Kp_X(dbl_Pressure, dbl_SO2, dbl_O2, dbl_Conv)

'-- Ref 7 is selected as the better --
int_Sel = 7

Select Case int_Sel

    Case 1
    
        dbl_Temp = Equil_Ref1_Kp(dbl_Kp_X)
    
    Case 2
    
        dbl_Temp = Equil_Ref2_Kp(dbl_Kp_X)
        
    Case 3
    
        dbl_Temp = Equil_Ref3_Kp(dbl_Kp_X)
        
    Case 4
    
        dbl_Temp = Equil_Ref4_Kp(dbl_Kp_X)
        
    Case 5
    
        dbl_Temp = Equil_Ref5_Kp(dbl_Kp_X)
        
    Case 6
    
        dbl_Temp = Equil_Ref6_Kp(dbl_Kp_X)
        
    Case 7
    
        dbl_Temp = Equil_Ref7_Kp(dbl_Kp_X)
        
End Select

xfTEquilSO2_SO3_PSO2O2X = dbl_Temp - 273.15

End Function

'---------------------------------------------------------------------------------------
' Function  : Kp_X
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns Kp at the selected conversion point
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2             --> SO2 feed composition (t.p.u)
'               dbl_O2              --> O2 feed composition (t.p.u)
'               dbl_Conv            --> Conversion point
'---------------------------------------------------------------------------------------
Private Function Kp_X(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Conv As Double) As Variant

Dim dbl_Term1 As Double
Dim dbl_Term2 As Double
Dim dbl_Term3 As Double
Dim dbl_Term4 As Double
Dim dbl_Term5 As Double

dbl_Term1 = dbl_Conv / (1 - dbl_Conv)

dbl_Term2 = 1 - dbl_SO2 * dbl_Conv / 2

dbl_Term3 = dbl_O2 - dbl_SO2 * dbl_Conv / 2

dbl_Term4 = (dbl_Term2 / dbl_Term3) ^ (1 / 2)

dbl_Term5 = (dbl_Pressure) ^ (-1 / 2)

Kp_X = dbl_Term1 * dbl_Term4 * dbl_Term5

End Function
'---------------------------------------------------------------------------------------
' Function  : Equil_Ref1_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Sulfuric Acid Manufacture" handbook
'               written by M.J.King & W.G.Davenport
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref1_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double
Dim R As Double

a = 0.09357
b = -98.41
R = 0.008134

Equil_Ref1_Kp = (-b / (a + R * Log(dbl_myKp)))

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref2_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "Elements of
'               Chemicals Reaction Engineering" - Folger, H.S.
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref2_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double

a = 4956
b = 4.678

Equil_Ref2_Kp = a / (Log10(dbl_myKp) + b)

End Function

'---------------------------------------------------------------------------------------
' Function  : Log10
' DateTime  : 11/12/2017
' Author    : Microsoft support
' Purpose   : Return logartihm base 10
' Arguments :
'               dbl_myln            --> Natural logarithm
'---------------------------------------------------------------------------------------
Private Function Log10(ByVal dbl_myln As Double) As Double

Log10 = Log(dbl_myln) / Log(10#)

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref3_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "Mechanism and
'               kinetics of SO2 Oxidation on "k-V" and "K-Na-V" Catalyst Series Kinetocs"
'               Ge, H.X., Han, Z.H. and Xie, K.C
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref3_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double

a = 4905.5
b = 4.6455

Equil_Ref3 = a / (Log10(dbl_myKp) + b)

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref4_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "Computer Simulation
'               of Sulfuric Acid Plant" Hannon, P.T, Johnson, A.I., Crowe, C.M.,
'               Hoffman, T.W., Hamielec, A.E., and Woods, D.R.
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref4_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double

Dim dbl_Error As Double
Dim dbl_Value As Double
Dim dbl_Temp As Double
Dim dbl_Goal As Double

Dim lon_Counter As Long

a = 12127
b = 11.423
c = 0.1309
d = 8.5 * (10 ^ -4)
e = 3.774 * (10 ^ 4)
lon_Counter = 0

dbl_Temp = 150
dbl_Error = 1
dbl_Goal = Log(dbl_myKp)

Do Until dbl_Error < 0.00001 And dbl_Error > -0.00001
    
    dbl_Temp = dbl_Temp - dbl_Temp * dbl_Error * 0.01
    
    dbl_Value = (a / dbl_Temp) - b - (c * Log(dbl_Temp)) + (d * dbl_Temp) - (e / (dbl_Temp ^ 2))
    dbl_Error = (dbl_Goal - dbl_Value) / dbl_Goal
    
    If dbl_Error > 1 Then
    
        dbl_Error = 1
        
    ElseIf dbl_Error < -1 Then
    
        dbl_Error = -1
        
    End If
    
    lon_Counter = lon_Counter + 1
    
    If lon_Counter > 10000 Then
    
        Exit Do
        
    End If
            
Loop

Equil_Ref4_Kp = dbl_Temp

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref5_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "MResearch Nat'l
'               Bur" Evans and Wagman, J.
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref5_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double

a = 5063
b = 4.82

Equil_Ref5_Kp = a / (Log10(dbl_myKp) + b)

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref6_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "Sulfuric Acid
'               Manufacture" Fairlie, A.M.
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref6_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double

a = 4760
b = 4.473

Equil_Ref6_Kp = a / (Log10(dbl_myKp) + b)

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref7_Kp
' DateTime  : 11/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature in K calculated based on
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "H2SO4 Atlas" Lurgi
'               GmbH
' Arguments :
'               dbl_myKp            --> Equilibrium constant at the selected conversion
'                                       point
'---------------------------------------------------------------------------------------
Private Function Equil_Ref7_Kp(ByVal dbl_myKp As Double) As Variant

Dim a As Double
Dim b As Double
Dim c As Double
Dim dbl_Error As Double
Dim dbl_Value As Double
Dim dbl_Temp As Double
Dim dbl_Goal As Double

Dim lon_Counter As Long

a = 5186.5
b = 6.75
c = 0.611
lon_Counter = 0

dbl_Temp = 150
dbl_Error = 1
dbl_Goal = Log10(dbl_myKp)


Do Until dbl_Error < 0.00001 And dbl_Error > -0.00001
    
    dbl_Temp = dbl_Temp - dbl_Temp * dbl_Error * 0.01
    dbl_Value = (a / dbl_Temp) + c * Log10(dbl_Temp) - b
    dbl_Error = (dbl_Goal - dbl_Value) / dbl_Goal

    If dbl_Error > 1 Then
    
        dbl_Error = 1
        
    ElseIf dbl_Error < -1 Then
    
        dbl_Error = -1
        
    End If
    
    lon_Counter = lon_Counter + 1
    
    If lon_Counter > 10000 Then
    
        Exit Do
        
    End If
            
Loop

Equil_Ref7_Kp = dbl_Temp

End Function

'---------------------------------------------------------------------------------------
' Function  : Equil_Ref7_T
' DateTime  : 13/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium constant value for the selected
'               temperature in K
'               the relationship extracted from "Handbook of Sulfuric Acid Manufacturing"
'               written by Douglas K. Louie. Function extracted from "H2SO4 Atlas" Lurgi
'               GmbH
' Arguments :
'               dbl_myT             --> Equilibrium temperature in K
'---------------------------------------------------------------------------------------
Private Function Equil_Ref7_T(ByVal dbl_myT As Double) As Variant

Dim a As Double
Dim b As Double
Dim c As Double

Dim dbl_Value As Double
Dim dbl_Kp As Double

a = 5186.5
b = 6.75
c = 0.611

dbl_Value = (a / dbl_myT) + c * Log10(dbl_myT) - b

dbl_Kp = 10 ^ dbl_Value

Equil_Ref7_T = dbl_Kp

End Function

'---------------------------------------------------------------------------------------
' Function  : X_Kp
' DateTime  : 13/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns Kp at the selected conversion point
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2             --> SO2 feed composition (t.p.u)
'               dbl_O2              --> O2 feed composition (t.p.u)
'               dbl_Kp              --> Equilibrium constant at the given T
'---------------------------------------------------------------------------------------
Private Function X_Kp(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Kp As Double) As Variant

Dim dbl_Term1 As Double
Dim dbl_Term2 As Double
Dim dbl_Term3 As Double
Dim dbl_Term4 As Double
Dim dbl_Term5 As Double
Dim dbl_Conv As Double

Dim dbl_Factor As Double

Dim dbl_Error As Double
Dim dbl_Value As Double

Dim lon_Counter As Long

lon_Counter = 0

dbl_Conv = 0.4
dbl_Error = 1

dbl_Factor = 0.01

Do Until dbl_Error < 0.00001 And dbl_Error > -0.00001
    
    dbl_Conv = dbl_Conv + dbl_Conv * dbl_Error * dbl_Factor
    
    If dbl_Conv > 1 Then
    
        dbl_Conv = 0.999999999
    
    End If
    
    dbl_Term1 = dbl_Conv / (1 - dbl_Conv)
    dbl_Term2 = 1 - dbl_SO2 * dbl_Conv / 2
    dbl_Term3 = dbl_O2 - dbl_SO2 * dbl_Conv / 2
    dbl_Term4 = (dbl_Term2 / dbl_Term3) ^ (1 / 2)
    dbl_Term5 = (dbl_Pressure) ^ (-1 / 2)
    
    dbl_Value = dbl_Term1 * dbl_Term4 * dbl_Term5
    
    dbl_Error = (dbl_Kp - dbl_Value) / dbl_Kp

    If dbl_Error > 1 Then
    
        dbl_Error = 1
        
    ElseIf dbl_Error < -1 Then
    
        dbl_Error = -1
        
    End If
    
    lon_Counter = lon_Counter + 1
    
    If dbl_Conv > 0.996 Then
    
        dbl_Factor = 0.0001
    
    ElseIf dbl_Conv > 0.998 Then
    
        dbl_Factor = 0.00001
        
    End If
        
    If lon_Counter > 10000 Then
        
        Exit Do
        
    End If
            
Loop

X_Kp = dbl_Conv

End Function
'---------------------------------------------------------------------------------------
' Function  : xfXEquilSO2_SO3_PSO2O2T
' DateTime  : 13/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium conversion
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2             --> SO2 feed composition (t.p.u)
'               dbl_O2              --> O2 feed composition (t.p.u)
'               dbl_Temp            --> Equilibrium temperature in ºC
'---------------------------------------------------------------------------------------
Public Function xfXEquilSO2_SO3_PSO2O2T(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Temp As Double) As Variant

Dim dbl_Kp_T As Double
Dim dbl_Value As Double

dbl_Kp_T = Equil_Ref7_T(dbl_Temp + 273.15)

dbl_Value = X_Kp(dbl_Pressure, dbl_SO2, dbl_O2, dbl_Kp_T)

xfXEquilSO2_SO3_PSO2O2T = dbl_Value

End Function
'---------------------------------------------------------------------------------------
' Function  : xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key
' DateTime  : 13/10/2017
' Author    : José García Herruzo
' Purpose   : This function calculates the intercept point in the equilibrium
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2i            --> SO2 feed composition (t.p.u)
'               dbl_O2i             --> O2 feed composition (t.p.u)
'               dbl_SO3             --> SO3 fed to the bed composition (t.p.u)
'               dbl_N2              --> N2 fed to the bed composition (t.p.u)
'               dbl_SO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_O2              --> O2 fed to the bed composition (t.p.u)
'               dbl_CO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_Temp            --> Bed inlet temperature (ºC)
'               dbl_Key             --> 0 to return Temp and 1 for conversion
'---------------------------------------------------------------------------------------
Public Function xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key(ByVal dbl_Pressure As Double, ByVal dbl_SO2i As Double, ByVal dbl_O2i As Double, ByVal dbl_CO2 As Double, _
                                                ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_N2 As Double, ByVal dbl_SO3 As Double, _
                                                ByVal dbl_Temp As Double, ByVal int_Key As Integer) As Variant

Dim dbl_Hin As Double
Dim dbl_HO2in As Double
Dim dbl_HSO2in As Double
Dim dbl_HSO3in As Double
Dim dbl_HCO2in As Double
Dim dbl_HN2in As Double

Dim dbl_Hout As Double
Dim dbl_HO2out As Double
Dim dbl_HSO2out As Double
Dim dbl_HSO3out As Double
Dim dbl_HCO2out As Double
Dim dbl_HN2out As Double

Dim dbl_O2out As Double
Dim dbl_SO2out As Double
Dim dbl_SO3out As Double
Dim dbl_CO2out As Double
Dim dbl_N2out As Double

Dim dbl_Tout As Double
Dim dbl_Conv As Double

Dim dbl_Convi As Double

Dim dbl_Error As Double
Dim int_Counter As Integer

'-- Calculate inlet enthalphy --
dbl_HO2in = xfHº_CompT("O2", dbl_Temp)
dbl_HSO2in = xfHº_CompT("SO2", dbl_Temp)
dbl_HSO3in = xfHº_CompT("SO3", dbl_Temp)
dbl_HCO2in = xfHº_CompT("CO2", dbl_Temp)
dbl_HN2in = xfHº_CompT("N2", dbl_Temp)

dbl_Hin = dbl_O2 * dbl_HO2in + dbl_SO2 * dbl_HSO2in + dbl_SO3 * dbl_HSO3in + dbl_CO2 * dbl_HCO2in + dbl_N2 * dbl_HN2in

'-- Set an initial conversion --
If dbl_Hin > 0 Then

    dbl_Convi = 0.9

Else

    dbl_Convi = 0.5
    
End If

int_Counter = 0
dbl_Error = 0.5

Do Until dbl_Error < 0.0001 And dbl_Error > -0.0001
    
    If int_Counter <> 0 Then
        '-- Conversion is updated --
        
        If dbl_Hin > 0 Then
        
            dbl_Convi = dbl_Convi - dbl_Convi * dbl_Error * 0.1
        
        Else
        
            dbl_Convi = dbl_Convi + dbl_Convi * dbl_Error * 0.1
            
        End If
        
        If dbl_Convi > 1 Then
        
            dbl_Convi = 1
        
        End If
        
    End If

    '-- Calculate molar composition based on bed conversion--
    dbl_O2out = dbl_O2 - (dbl_SO2 * dbl_Convi) / 2
    dbl_SO2out = dbl_SO2 - (dbl_SO2 * dbl_Convi)
    dbl_SO3out = dbl_SO3 + (dbl_SO2 * dbl_Convi)
    dbl_CO2out = dbl_CO2
    dbl_N2out = dbl_N2
    
    '-- Calculate whole process conversion
    dbl_Conv = 1 - dbl_SO2out / dbl_SO2i
    
    '-- Determinate the temperature of the equilibrium --
    dbl_Tout = xfTEquilSO2_SO3_PSO2O2X(dbl_Pressure, dbl_SO2i, dbl_O2i, dbl_Conv)
    
    '-- Once Temperature is determinated, extract outlet enthalphies --
    dbl_HO2out = xfHº_CompT("O2", dbl_Tout)
    dbl_HSO2out = xfHº_CompT("SO2", dbl_Tout)
    dbl_HSO3out = xfHº_CompT("SO3", dbl_Tout)
    dbl_HCO2out = xfHº_CompT("CO2", dbl_Tout)
    dbl_HN2out = xfHº_CompT("N2", dbl_Tout)
    
    dbl_Hout = dbl_O2out * dbl_HO2out + dbl_SO2out * dbl_HSO2out + dbl_SO3out * dbl_HSO3out + dbl_CO2out * dbl_HCO2out + dbl_N2out * dbl_HN2out
    
    '-- Enthalpies are compared --
    dbl_Error = (dbl_Hin - dbl_Hout) / dbl_Hin
    
    int_Counter = int_Counter + 1
    
    '-- Update loop values --
    If int_Counter > 1000 Then
        
        Exit Do
    
    End If
    
Loop

Select Case int_Key

    Case 0
    
        xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key = dbl_Tout
        
    Case 1
    
        xfTXEquilSO2_SO3_PSO2iO2iCO2SO2O2N2SO3Ti_Key = dbl_Conv
        
End Select

End Function

'---------------------------------------------------------------------------------------
' Function  : xfBedX_CO2SO2O2N2SO3TinTout
' DateTime  : 14/10/2017
' Author    : José García Herruzo
' Purpose   : This function calculates bed equilibrium depending on its outlet Temp.
' Arguments :
'               dbl_SO3             --> SO3 fed to the bed composition (t.p.u)
'               dbl_N2              --> N2 fed to the bed composition (t.p.u)
'               dbl_SO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_O2              --> O2 fed to the bed composition (t.p.u)
'               dbl_CO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_Tin             --> Bed inlet temperature (ºC)
'               dbl_Tout            --> Bed outlet temperature (ºC)
'---------------------------------------------------------------------------------------
Public Function xfBedX_CO2SO2O2N2SO3TinTout(ByVal dbl_CO2 As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_N2 As Double, ByVal dbl_SO3 As Double, _
                                                ByVal dbl_Tin As Double, ByVal dbl_Tout As Double) As Variant

Dim dbl_Hin As Double
Dim dbl_HO2in As Double
Dim dbl_HSO2in As Double
Dim dbl_HSO3in As Double
Dim dbl_HCO2in As Double
Dim dbl_HN2in As Double

Dim dbl_Hout As Double
Dim dbl_HO2out As Double
Dim dbl_HSO2out As Double
Dim dbl_HSO3out As Double
Dim dbl_HCO2out As Double
Dim dbl_HN2out As Double

Dim dbl_O2out As Double
Dim dbl_SO2out As Double
Dim dbl_SO3out As Double
Dim dbl_CO2out As Double
Dim dbl_N2out As Double

Dim dbl_Conv As Double

Dim dbl_Error As Double
Dim int_Counter As Integer

Dim dbl_Factor As Double
Dim dbl_Limit As Double

'-- Calculate inlet enthalphy --
dbl_HO2in = xfHº_CompT("O2", dbl_Tin)
dbl_HSO2in = xfHº_CompT("SO2", dbl_Tin)
dbl_HSO3in = xfHº_CompT("SO3", dbl_Tin)
dbl_HCO2in = xfHº_CompT("CO2", dbl_Tin)
dbl_HN2in = xfHº_CompT("N2", dbl_Tin)

dbl_Hin = dbl_O2 * dbl_HO2in + dbl_SO2 * dbl_HSO2in + dbl_SO3 * dbl_HSO3in + dbl_CO2 * dbl_HCO2in + dbl_N2 * dbl_HN2in

'-- Set an initial conversion --
If dbl_Hin > 0 Then

    dbl_Conv = 0.8

Else

    dbl_Conv = 0.5
    
End If

'-- Modify convergence process depending on the bed --
If dbl_Tout - dbl_Tin < 2 Then

    dbl_Factor = 1
    dbl_Limit = 0.00001
    dbl_Conv = 0.1
    
Else

    dbl_Factor = 0.1
    dbl_Limit = 0.0001

End If


int_Counter = 0
dbl_Error = 0.5

Do Until dbl_Error < dbl_Limit And dbl_Error > -dbl_Limit
    
    If int_Counter <> 0 Then
        '-- Conversion is updated --
        
        If dbl_Hin > 0 Then
        
            dbl_Conv = dbl_Conv - dbl_Conv * dbl_Factor * dbl_Error
        
        Else
        
            dbl_Conv = dbl_Conv + dbl_Conv * dbl_Factor * dbl_Error
            
        End If
        
        If dbl_Conv > 1 Then
        
            dbl_Conv = 1
        
        End If
        
    End If

    '-- Calculate molar composition --
    dbl_O2out = dbl_O2 - (dbl_SO2 * dbl_Conv) / 2
    dbl_SO2out = dbl_SO2 - (dbl_SO2 * dbl_Conv)
    dbl_SO3out = dbl_SO3 + (dbl_SO2 * dbl_Conv)
    dbl_CO2out = dbl_CO2
    dbl_N2out = dbl_N2
    
    '-- Once Temperature is determinated, extract outlet enthalphies --
    dbl_HO2out = xfHº_CompT("O2", dbl_Tout)
    dbl_HSO2out = xfHº_CompT("SO2", dbl_Tout)
    dbl_HSO3out = xfHº_CompT("SO3", dbl_Tout)
    dbl_HCO2out = xfHº_CompT("CO2", dbl_Tout)
    dbl_HN2out = xfHº_CompT("N2", dbl_Tout)
    
    dbl_Hout = dbl_O2out * dbl_HO2out + dbl_SO2out * dbl_HSO2out + dbl_SO3out * dbl_HSO3out + dbl_CO2out * dbl_HCO2out + dbl_N2out * dbl_HN2out
    
    '-- Enthalpies are compared --
    dbl_Error = (dbl_Hin - dbl_Hout) / dbl_Hin
    
    int_Counter = int_Counter + 1
    
    '-- Update loop values --
    If int_Counter > 1000 Then
        
        Exit Do
    
    End If
    
Loop
    
    xfBedX_CO2SO2O2N2SO3TinTout = dbl_Conv

End Function

'---------------------------------------------------------------------------------------
' Function  : xfBedT_CO2SO2O2N2SO3TinConv
' DateTime  : 14/10/2017
' Author    : José García Herruzo
' Purpose   : This function calculates bed equilibrium depending on its outlet Temp.
' Arguments :
'               dbl_SO3             --> SO3 fed to the bed composition (t.p.u)
'               dbl_N2              --> N2 fed to the bed composition (t.p.u)
'               dbl_SO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_O2              --> O2 fed to the bed composition (t.p.u)
'               dbl_CO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_Tin             --> Bed inlet temperature (ºC)
'               dbl_Tout            --> Bed outlet temperature (ºC)
'---------------------------------------------------------------------------------------
Public Function xfBedT_CO2SO2O2N2SO3TinConv(ByVal dbl_CO2 As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_N2 As Double, ByVal dbl_SO3 As Double, _
                                                ByVal dbl_Tin As Double, ByVal dbl_Conv As Double) As Variant

Dim dbl_Hin As Double
Dim dbl_HO2in As Double
Dim dbl_HSO2in As Double
Dim dbl_HSO3in As Double
Dim dbl_HCO2in As Double
Dim dbl_HN2in As Double

Dim dbl_Hout As Double
Dim dbl_HO2out As Double
Dim dbl_HSO2out As Double
Dim dbl_HSO3out As Double
Dim dbl_HCO2out As Double
Dim dbl_HN2out As Double

Dim dbl_O2out As Double
Dim dbl_SO2out As Double
Dim dbl_SO3out As Double
Dim dbl_CO2out As Double
Dim dbl_N2out As Double

Dim dbl_Tout As Double

Dim dbl_Error As Double
Dim int_Counter As Integer


'-- Calculate inlet enthalphy --
dbl_HO2in = xfHº_CompT("O2", dbl_Tin)
dbl_HSO2in = xfHº_CompT("SO2", dbl_Tin)
dbl_HSO3in = xfHº_CompT("SO3", dbl_Tin)
dbl_HCO2in = xfHº_CompT("CO2", dbl_Tin)
dbl_HN2in = xfHº_CompT("N2", dbl_Tin)

dbl_Hin = dbl_O2 * dbl_HO2in + dbl_SO2 * dbl_HSO2in + dbl_SO3 * dbl_HSO3in + dbl_CO2 * dbl_HCO2in + dbl_N2 * dbl_HN2in

'-- Set an initial conversion --
dbl_Tout = 500

int_Counter = 0
dbl_Error = 0.5

Do Until dbl_Error < 0.0001 And dbl_Error > -0.0001
    
    If int_Counter <> 0 Then
        '-- Conversion is updated --
        
        If dbl_Hin > 0 Then
        
            dbl_Tout = dbl_Tout + dbl_Tout * dbl_Error * 0.1
        
        Else
        
            dbl_Tout = dbl_Tout - dbl_Tout * dbl_Error * 0.1
            
        End If
        
    End If

    '-- Calculate molar composition --
    dbl_O2out = dbl_O2 - (dbl_SO2 * dbl_Conv) / 2
    dbl_SO2out = dbl_SO2 - (dbl_SO2 * dbl_Conv)
    dbl_SO3out = dbl_SO3 + (dbl_SO2 * dbl_Conv)
    dbl_CO2out = dbl_CO2
    dbl_N2out = dbl_N2
    
    '-- Once Temperature is determinated, extract outlet enthalphies --
    dbl_HO2out = xfHº_CompT("O2", dbl_Tout)
    dbl_HSO2out = xfHº_CompT("SO2", dbl_Tout)
    dbl_HSO3out = xfHº_CompT("SO3", dbl_Tout)
    dbl_HCO2out = xfHº_CompT("CO2", dbl_Tout)
    dbl_HN2out = xfHº_CompT("N2", dbl_Tout)
    
    dbl_Hout = dbl_O2out * dbl_HO2out + dbl_SO2out * dbl_HSO2out + dbl_SO3out * dbl_HSO3out + dbl_CO2out * dbl_HCO2out + dbl_N2out * dbl_HN2out
    
    '-- Enthalpies are compared --
    dbl_Error = (dbl_Hin - dbl_Hout) / dbl_Hin
    
    int_Counter = int_Counter + 1
    
    '-- Update loop values --
    If int_Counter > 1000 Then
        
        Exit Do
    
    End If
    
Loop
    
    xfBedT_CO2SO2O2N2SO3TinConv = dbl_Tout

End Function

'---------------------------------------------------------------------------------------
' Function  : xfXEquilAIATSO2_SO3_PSO2O2TPaSO2aO2a
' DateTime  : 19/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium conversion after IAT
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2             --> SO2 feed composition (t.p.u)
'               dbl_O2              --> O2 feed composition (t.p.u)
'               dbl_Temp            --> Equilibrium temperature in ºC
'               dbl_PressureA       --> System pressure after IAT(bara)
'               dbl_SO2A            --> SO2 after IAT composition (t.p.u)
'               dbl_O2A             --> O2 after IAT composition (t.p.u)
'---------------------------------------------------------------------------------------
Public Function xfXEquilAIATSO2_SO3_PSO2O2TPaSO2aO2a(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Temp As Double, _
                                            ByVal dbl_PressureA As Double, ByVal dbl_SO2A As Double, ByVal dbl_O2A As Double) As Variant

Dim dbl_Kp_T As Double
Dim dbl_e1 As Double

Dim dbl_e2 As Double

Dim dbl_Value As Double

dbl_Kp_T = Equil_Ref7_T(dbl_Temp + 273.15)

dbl_e1 = X_Kp(dbl_Pressure, dbl_SO2, dbl_O2, dbl_Kp_T)
dbl_e2 = X_Kp(dbl_PressureA, dbl_SO2A, dbl_O2A, dbl_Kp_T)

dbl_Value = 1 - (1 - dbl_e1) * (1 - dbl_e2)

xfXEquilAIATSO2_SO3_PSO2O2TPaSO2aO2a = dbl_Value

End Function

'---------------------------------------------------------------------------------------
' Function  : xfTEquilAIATSO2_SO3_PSO2O2ConvPaSO2aO2a
' DateTime  : 19/12/2017
' Author    : José García Herruzo
' Purpose   : This function returns the equilibrium temperature after IAT in ºC
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2             --> SO2 feed composition (t.p.u)
'               dbl_O2              --> O2 feed composition (t.p.u)
'               dbl_Conv            --> Equilibrium conversion
'               dbl_PressureA       --> System pressure after IAT(bara)
'               dbl_SO2A            --> SO2 after IAT composition (t.p.u)
'               dbl_O2A             --> O2 after IAT composition (t.p.u)
'---------------------------------------------------------------------------------------
Public Function xfTEquilAIATSO2_SO3_PSO2O2ConvPaSO2aO2a(ByVal dbl_Pressure As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_Conv As Double, _
                                            ByVal dbl_PressureA As Double, ByVal dbl_SO2A As Double, ByVal dbl_O2A As Double) As Variant

Dim dbl_Kp_T As Double
Dim dbl_e1 As Double
Dim dbl_Temp As Double
Dim dbl_e2 As Double

Dim dbl_Value As Double
Dim dbl_Error As Double
Dim int_Counter As Integer

dbl_Temp = 410
int_Counter = 0
dbl_Error = 0.5

Do Until dbl_Error < 0.0001 And dbl_Error > -0.0001

    If int_Counter <> 0 Then
        '-- Temperature is updated --
        dbl_Temp = dbl_Temp - dbl_Temp * dbl_Error
        
    End If
    
    dbl_Kp_T = Equil_Ref7_T(dbl_Temp + 273.15)
    
    dbl_e1 = X_Kp(dbl_Pressure, dbl_SO2, dbl_O2, dbl_Kp_T)
    dbl_e2 = X_Kp(dbl_PressureA, dbl_SO2A, dbl_O2A, dbl_Kp_T)
    
    dbl_Value = 1 - (1 - dbl_e1) * (1 - dbl_e2)
    
    dbl_Error = (dbl_Conv - dbl_Value) / dbl_Conv
    
    int_Counter = int_Counter + 1
    
    '-- Update loop values --
    If int_Counter > 300 Then
        
        Exit Do
    
    End If
    
Loop

xfTEquilAIATSO2_SO3_PSO2O2ConvPaSO2aO2a = dbl_Temp

End Function

'---------------------------------------------------------------------------------------
' Function  : xfTXEquilaIATSO2_SO3_PSO2iO2iCO2SO2O2N2SO3TiPaSO2aO2a_Key
' DateTime  : 19/12/2017
' Author    : José García Herruzo
' Purpose   : This function calculates the intercept point in the equilibrium after IAT
' Arguments :
'               dbl_Pressure        --> System pressure (bara)
'               dbl_SO2i            --> SO2 feed composition (t.p.u)
'               dbl_O2i             --> O2 feed composition (t.p.u)
'               dbl_SO3             --> SO3 fed to the bed composition (t.p.u)
'               dbl_N2              --> N2 fed to the bed composition (t.p.u)
'               dbl_SO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_O2              --> O2 fed to the bed composition (t.p.u)
'               dbl_CO2             --> SO2 fed to the bed composition (t.p.u)
'               dbl_Temp            --> Bed inlet temperature (ºC)
'               dbl_PressureA       --> System pressure after IAT(bara)
'               dbl_SO2A            --> SO2 after IAT composition (t.p.u)
'               dbl_O2A             --> O2 after IAT composition (t.p.u)
'               dbl_Key             --> 0 to return Temp and 1 for conversion
'---------------------------------------------------------------------------------------
Public Function xfTXEquilaIATSO2_SO3_PSO2iO2iCO2SO2O2N2SO3TiPaSO2aO2a_Key(ByVal dbl_Pressure As Double, ByVal dbl_SO2i As Double, ByVal dbl_O2i As Double, ByVal dbl_CO2 As Double, _
                                                ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_N2 As Double, ByVal dbl_SO3 As Double, _
                                                ByVal dbl_Temp As Double, ByVal dbl_PressureA As Double, ByVal dbl_SO2A As Double, ByVal dbl_O2A As Double, ByVal int_Key As Integer) As Variant

Dim dbl_Hin As Double
Dim dbl_HO2in As Double
Dim dbl_HSO2in As Double
Dim dbl_HSO3in As Double
Dim dbl_HCO2in As Double
Dim dbl_HN2in As Double

Dim dbl_Hout As Double
Dim dbl_HO2out As Double
Dim dbl_HSO2out As Double
Dim dbl_HSO3out As Double
Dim dbl_HCO2out As Double
Dim dbl_HN2out As Double

Dim dbl_O2out As Double
Dim dbl_SO2out As Double
Dim dbl_SO3out As Double
Dim dbl_CO2out As Double
Dim dbl_N2out As Double

Dim dbl_Tout As Double
Dim dbl_Conv As Double

Dim dbl_Convi As Double

Dim dbl_Error As Double
Dim int_Counter As Integer
Dim dbl_Factor As Double

'-- Calculate inlet enthalphy --
dbl_HO2in = xfHº_CompT("O2", dbl_Temp)
dbl_HSO2in = xfHº_CompT("SO2", dbl_Temp)
dbl_HSO3in = xfHº_CompT("SO3", dbl_Temp)
dbl_HCO2in = xfHº_CompT("CO2", dbl_Temp)
dbl_HN2in = xfHº_CompT("N2", dbl_Temp)

dbl_Hin = dbl_O2 * dbl_HO2in + dbl_SO2 * dbl_HSO2in + dbl_SO3 * dbl_HSO3in + dbl_CO2 * dbl_HCO2in + dbl_N2 * dbl_HN2in

'-- Set an initial conversion --
If dbl_Hin > 0 Then

    dbl_Convi = 0.9

Else

    dbl_Convi = 0.5
    
End If

int_Counter = 0
dbl_Error = 0.5

Do Until dbl_Error < 0.0001 And dbl_Error > -0.0001
    
    If int_Counter <> 0 Then
        '-- Conversion is updated --
        
        If dbl_Convi > 0.98 Then
        
            dbl_Factor = dbl_Error * 0.01
        
        Else
        
            dbl_Factor = dbl_Error * 0.1
            
        End If
        
        If dbl_Hin > 0 Then
        
            dbl_Convi = dbl_Convi - dbl_Convi * dbl_Factor
        
        Else
        
            dbl_Convi = dbl_Convi + dbl_Convi * dbl_Factor
            
        End If
        
        If dbl_Convi > 1 Then
        
            dbl_Convi = 1
        
        End If
        
    End If

    '-- Calculate molar composition based on bed conversion--
    dbl_O2out = dbl_O2 - (dbl_SO2 * dbl_Convi) / 2
    dbl_SO2out = dbl_SO2 - (dbl_SO2 * dbl_Convi)
    dbl_SO3out = dbl_SO3 + (dbl_SO2 * dbl_Convi)
    dbl_CO2out = dbl_CO2
    dbl_N2out = dbl_N2
    
    '-- Calculate whole process conversion
    dbl_Conv = 1 - dbl_SO2out / dbl_SO2i
    
    '-- Determinate the temperature of the equilibrium --
    dbl_Tout = xfTEquilAIATSO2_SO3_PSO2O2ConvPaSO2aO2a(dbl_Pressure, dbl_SO2i, dbl_O2i, dbl_Conv, dbl_PressureA, dbl_SO2A, dbl_O2A)
    
    '-- Once Temperature is determinated, extract outlet enthalphies --
    dbl_HO2out = xfHº_CompT("O2", dbl_Tout)
    dbl_HSO2out = xfHº_CompT("SO2", dbl_Tout)
    dbl_HSO3out = xfHº_CompT("SO3", dbl_Tout)
    dbl_HCO2out = xfHº_CompT("CO2", dbl_Tout)
    dbl_HN2out = xfHº_CompT("N2", dbl_Tout)
    
    dbl_Hout = dbl_O2out * dbl_HO2out + dbl_SO2out * dbl_HSO2out + dbl_SO3out * dbl_HSO3out + dbl_CO2out * dbl_HCO2out + dbl_N2out * dbl_HN2out
    
    '-- Enthalpies are compared --
    dbl_Error = (dbl_Hin - dbl_Hout) / dbl_Hin
    
    int_Counter = int_Counter + 1
    
    '-- Update loop values --
    If int_Counter > 1000 Then
        
        Exit Do
    
    End If
    
Loop

Select Case int_Key

    Case 0
    
        xfTXEquilaIATSO2_SO3_PSO2iO2iCO2SO2O2N2SO3TiPaSO2aO2a_Key = dbl_Tout
        
    Case 1
    
        xfTXEquilaIATSO2_SO3_PSO2iO2iCO2SO2O2N2SO3TiPaSO2aO2a_Key = dbl_Conv
        
End Select

End Function

'---------------------------------------------------------------------------------------
' Function  : xfCpDil_T_Conc
' DateTime  : 26/07/2019
' Author    : José García Herruzo
' Purpose   : This function returns Cp for dilute acid between 0 and 10%
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function xfCpDil_T_Conc(ByVal dbl_T As Double, ByVal dbl_Conc As Double) As Variant

Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_C As Double
Dim dbl_D As Double
Dim dbl_E As Double
Dim dbl_F As Double

Dim dbl_Value As Double

dbl_A = 4.23054
dbl_B = 0.002305
dbl_C = 2.1469
dbl_D = 0.000024
dbl_E = 6.193
dbl_F = 0.001604

dbl_Value = dbl_A - dbl_B * dbl_T - dbl_C * dbl_Conc + dbl_D * dbl_T * dbl_T - dbl_E * dbl_Conc * dbl_Conc - dbl_F * dbl_T * dbl_Conc

xfCpDil_T_Conc = dbl_Value

End Function

'---------------------------------------------------------------------------------------
' Function  : xfRhoDil_T_Conc
' DateTime  : 26/07/2019
' Author    : José García Herruzo
' Purpose   : This function returns density for dilute acid between 0 and 10%
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function xfRhoDil_T_Conc(ByVal dbl_T As Double, ByVal dbl_Conc As Double) As Variant

Dim dbl_A As Double
Dim dbl_B As Double
Dim dbl_C As Double
Dim dbl_D As Double
Dim dbl_E As Double
Dim dbl_F As Double

Dim dbl_Value As Double

dbl_A = 1003.63
dbl_B = 0.14223
dbl_C = 942.687
dbl_D = 0.003099
dbl_E = 414.5
dbl_F = 2.31391

dbl_Value = dbl_A - dbl_B * dbl_T + dbl_C * dbl_Conc - dbl_D * dbl_T * dbl_T + dbl_E * dbl_Conc * dbl_Conc - dbl_F * dbl_T * dbl_Conc

xfRhoDil_T_Conc = dbl_Value

End Function
