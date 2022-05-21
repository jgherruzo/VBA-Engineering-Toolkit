Attribute VB_Name = "ModProperties_ThermoChemical"
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
' Module      : ModProperties_ThermoChemical
' DateTime    : 10/10/2017
' Author      : José García Herruzo
' Purpose     : This module contents properties data and calculation
' References  : N/A
' Requirements: N/A
' Functions   :
'               01-xfCp_CompT
'               02-xfHº_CompT (kJ/kmol)
' Procedures  :
'               01-xlLoadCompoundsArray
'               02-xlLoadThermoChemData
'               03-xlLoadFormationData
'               04-xpLookForThermoChemIndex
'               05-xpLookForHfIndex
' Updates     :
'       DATE        USER    DESCRIPTION
'       11/10/2017  JGH     Function 02 and procedure 05 are added
'       13/12/2017  JGH     Function 01 and 02 are modified to show the argument
'-----------------------------------------------------------------------------------------
                                '<< Compounds Data reference    >>
                                '== NIST                        ==
                                '<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>

Dim arr_ThermoChemData As Variant
Dim arr_FormationData As Variant
Dim lon_ElementsNumber As Long
Dim lon_PropertyNumber As Long
'---------------------------------------------------------------------------------------
' Procedure : xlLoadCompoundsArray
' DateTime  : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This procedure loads an array containing the compounds data
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub xlLoadCompoundsArray()

'-- First, number of elements is specified to dimensionate the array --
lon_ElementsNumber = 12
lon_PropertyNumber = 10

'-- Generate the array --
ReDim arr_ThermoChemData(lon_ElementsNumber, lon_PropertyNumber)
ReDim arr_FormationData(lon_ElementsNumber, lon_PropertyNumber)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlLoadThermoChemData
' DateTime    : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This procedure loads an array containing the compounds Termo chemical data
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub xlLoadThermoChemData()

Dim i As Integer

'-- Index --
i = 0
arr_ThermoChemData(i, 0) = "Compound"
arr_ThermoChemData(i, 1) = "Tmin, K"
arr_ThermoChemData(i, 2) = "Tmax, K"
arr_ThermoChemData(i, 3) = "A"
arr_ThermoChemData(i, 4) = "B"
arr_ThermoChemData(i, 5) = "C"
arr_ThermoChemData(i, 6) = "D"
arr_ThermoChemData(i, 7) = "E"
arr_ThermoChemData(i, 8) = "F"
arr_ThermoChemData(i, 9) = "G"
arr_ThermoChemData(i, 10) = "H"

'-- SO2 Range 298-1200 --
i = 1
arr_ThermoChemData(i, 0) = "SO2"
arr_ThermoChemData(i, 1) = 298
arr_ThermoChemData(i, 2) = 1200
arr_ThermoChemData(i, 3) = 21.43049
arr_ThermoChemData(i, 4) = 74.35094
arr_ThermoChemData(i, 5) = -57.75217
arr_ThermoChemData(i, 6) = 16.35534
arr_ThermoChemData(i, 7) = 0.086731
arr_ThermoChemData(i, 8) = -305.7688
arr_ThermoChemData(i, 9) = 254.8872
arr_ThermoChemData(i, 10) = -296.8422

'-- SO2 Range 1200-1600 --
i = 2
arr_ThermoChemData(i, 0) = "SO2"
arr_ThermoChemData(i, 1) = 1200
arr_ThermoChemData(i, 2) = 1600
arr_ThermoChemData(i, 3) = 57.48188
arr_ThermoChemData(i, 4) = 1.009328
arr_ThermoChemData(i, 5) = -0.07629
arr_ThermoChemData(i, 6) = 0.005174
arr_ThermoChemData(i, 7) = -4.045401
arr_ThermoChemData(i, 8) = -324.414
arr_ThermoChemData(i, 9) = 302.7798
arr_ThermoChemData(i, 10) = -296.8422

'-- SO3 Range 298-1200 --
i = 3
arr_ThermoChemData(i, 0) = "SO3"
arr_ThermoChemData(i, 1) = 298
arr_ThermoChemData(i, 2) = 1200
arr_ThermoChemData(i, 3) = 24.02503
arr_ThermoChemData(i, 4) = 119.4607
arr_ThermoChemData(i, 5) = -94.38686
arr_ThermoChemData(i, 6) = 26.96237
arr_ThermoChemData(i, 7) = -0.117517
arr_ThermoChemData(i, 8) = -407.8526
arr_ThermoChemData(i, 9) = 253.5186
arr_ThermoChemData(i, 10) = -395.7654

'-- SO3 Range 1200-1600 --
i = 4
arr_ThermoChemData(i, 0) = "SO3"
arr_ThermoChemData(i, 1) = 1200
arr_ThermoChemData(i, 2) = 1600
arr_ThermoChemData(i, 3) = 81.99008
arr_ThermoChemData(i, 4) = 0.622236
arr_ThermoChemData(i, 5) = -0.12244
arr_ThermoChemData(i, 6) = 0.008294
arr_ThermoChemData(i, 7) = -6.703688
arr_ThermoChemData(i, 8) = -437.659
arr_ThermoChemData(i, 9) = 330.9264
arr_ThermoChemData(i, 10) = -395.7654

'-- O2 Range 100-700 --
i = 5
arr_ThermoChemData(i, 0) = "O2"
arr_ThermoChemData(i, 1) = 100
arr_ThermoChemData(i, 2) = 700
arr_ThermoChemData(i, 3) = 31.3223
arr_ThermoChemData(i, 4) = -20.23531
arr_ThermoChemData(i, 5) = 57.86644
arr_ThermoChemData(i, 6) = -36.50624
arr_ThermoChemData(i, 7) = -0.007374
arr_ThermoChemData(i, 8) = -8.903471
arr_ThermoChemData(i, 9) = 246.7945
arr_ThermoChemData(i, 10) = 0

'-- O2 Range 700-2000 --
i = 6
arr_ThermoChemData(i, 0) = "O2"
arr_ThermoChemData(i, 1) = 700
arr_ThermoChemData(i, 2) = 2000
arr_ThermoChemData(i, 3) = 30.03235
arr_ThermoChemData(i, 4) = 8.772972
arr_ThermoChemData(i, 5) = -3.988133
arr_ThermoChemData(i, 6) = 0.788313
arr_ThermoChemData(i, 7) = -0.741599
arr_ThermoChemData(i, 8) = -11.32468
arr_ThermoChemData(i, 9) = 236.1663
arr_ThermoChemData(i, 10) = 0

'-- O2 Range 2000-6000 --
i = 7
arr_ThermoChemData(i, 0) = "O2"
arr_ThermoChemData(i, 1) = 2000
arr_ThermoChemData(i, 2) = 6000
arr_ThermoChemData(i, 3) = 20.91111
arr_ThermoChemData(i, 4) = 10.72071
arr_ThermoChemData(i, 5) = -2.020498
arr_ThermoChemData(i, 6) = 0.146449
arr_ThermoChemData(i, 7) = 9.245722
arr_ThermoChemData(i, 8) = 5.337651
arr_ThermoChemData(i, 9) = 237.185
arr_ThermoChemData(i, 10) = 0

'-- N2 Range 100-500 --
i = 8
arr_ThermoChemData(i, 0) = "N2"
arr_ThermoChemData(i, 1) = 100
arr_ThermoChemData(i, 2) = 500
arr_ThermoChemData(i, 3) = 28.98641
arr_ThermoChemData(i, 4) = 1.853978
arr_ThermoChemData(i, 5) = -9.647459
arr_ThermoChemData(i, 6) = 16.63537
arr_ThermoChemData(i, 7) = 0.000117
arr_ThermoChemData(i, 8) = -8.671914
arr_ThermoChemData(i, 9) = 226.4168
arr_ThermoChemData(i, 10) = 0

'-- N2 Range 500-2000 --
i = 9
arr_ThermoChemData(i, 0) = "N2"
arr_ThermoChemData(i, 1) = 500
arr_ThermoChemData(i, 2) = 2000
arr_ThermoChemData(i, 3) = 19.50583
arr_ThermoChemData(i, 4) = 19.88705
arr_ThermoChemData(i, 5) = -8.598535
arr_ThermoChemData(i, 6) = 1.369784
arr_ThermoChemData(i, 7) = 0.527601
arr_ThermoChemData(i, 8) = -4.935202
arr_ThermoChemData(i, 9) = 212.39
arr_ThermoChemData(i, 10) = 0

'-- N2 Range 2000-6000 --
i = 10
arr_ThermoChemData(i, 0) = "N2"
arr_ThermoChemData(i, 1) = 2000
arr_ThermoChemData(i, 2) = 6000
arr_ThermoChemData(i, 3) = 35.51872
arr_ThermoChemData(i, 4) = 1.128728
arr_ThermoChemData(i, 5) = -0.196103
arr_ThermoChemData(i, 6) = 0.014662
arr_ThermoChemData(i, 7) = -4.55376
arr_ThermoChemData(i, 8) = -18.97091
arr_ThermoChemData(i, 9) = 224.981
arr_ThermoChemData(i, 10) = 0

'-- CO2 Range 298-1200 --
i = 11
arr_ThermoChemData(i, 0) = "CO2"
arr_ThermoChemData(i, 1) = 298
arr_ThermoChemData(i, 2) = 1200
arr_ThermoChemData(i, 3) = 24.99735
arr_ThermoChemData(i, 4) = 55.18696
arr_ThermoChemData(i, 5) = -33.69137
arr_ThermoChemData(i, 6) = 7.948387
arr_ThermoChemData(i, 7) = -0.136638
arr_ThermoChemData(i, 8) = -403.6075
arr_ThermoChemData(i, 9) = 228.2431
arr_ThermoChemData(i, 10) = -393.5224

'-- CO2 Range 1200-1600 --
i = 12
arr_ThermoChemData(i, 0) = "CO2"
arr_ThermoChemData(i, 1) = 1200
arr_ThermoChemData(i, 2) = 1600
arr_ThermoChemData(i, 3) = 58.16639
arr_ThermoChemData(i, 4) = 2.720074
arr_ThermoChemData(i, 5) = -0.492289
arr_ThermoChemData(i, 6) = 0.038844
arr_ThermoChemData(i, 7) = -6.447293
arr_ThermoChemData(i, 8) = -425.9186
arr_ThermoChemData(i, 9) = 263.6125
arr_ThermoChemData(i, 10) = -393.5224

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlLoadFormationData
' DateTime    : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This procedure loads an array containing the compounds Termo chemical data
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Sub xlLoadFormationData()

Dim i As Integer

'-- Index --
i = 0
arr_FormationData(i, 0) = "Compound"
arr_FormationData(i, 1) = "Hfº, kJ/kmol"
arr_FormationData(i, 2) = "Sº,1 bar, J/kmol K"


'-- SO2 --
i = 1
arr_FormationData(i, 0) = "SO2"
arr_FormationData(i, 1) = -296.84 * 1000
arr_FormationData(i, 2) = 248.21 / 1000

'-- SO3 --
i = 2
arr_FormationData(i, 0) = "SO3"
arr_FormationData(i, 1) = -395.77 * 1000
arr_FormationData(i, 2) = 256.77 / 1000

'-- O2 Range --
i = 3
arr_FormationData(i, 0) = "O2"
arr_FormationData(i, 1) = 0
arr_FormationData(i, 2) = 0

'-- N2 Range --
i = 4
arr_FormationData(i, 0) = "N2"
arr_FormationData(i, 1) = 0
arr_FormationData(i, 2) = 191.61 / 1000

'-- CO2 Range 298-1200 --
i = 5
arr_FormationData(i, 0) = "CO2"
arr_FormationData(i, 1) = -393.52 * 1000
arr_FormationData(i, 2) = 213.79 / 1000

End Sub
'---------------------------------------------------------------------------------------
' Procedure : xpLookForThermoChemIndex
' DateTime  : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the array index for Thermochem calculation
' Arguments :
'               str_Component       --> Compound formula
'               str_Temperature     --> System temperature
'---------------------------------------------------------------------------------------
Private Function xpLookForThermoChemIndex(ByVal str_myComponent As String, ByVal dbl_myTemperature As Double) As Variant

Dim i As Long
Dim j As Integer

Dim lon_myIndex As Long

lon_myIndex = 0
For i = 0 To lon_ElementsNumber

    If str_myComponent = arr_ThermoChemData(i, 0) Then
    
        If dbl_myTemperature >= arr_ThermoChemData(i, 1) And dbl_myTemperature < arr_ThermoChemData(i, 2) Then
        
            lon_myIndex = i
            Exit For
        End If
    
    End If

Next i

If lon_myIndex = 0 Then
    
    '-- Compound not found --
    xpLookForThermoChemIndex = -1

Else

    xpLookForThermoChemIndex = lon_myIndex
        
End If

End Function
'---------------------------------------------------------------------------------------
' Procedure : xpLookForHfIndex
' DateTime  : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the array index for formation enthalpy array
' Arguments :
'               str_Compound        --> Compound formula
'---------------------------------------------------------------------------------------
Private Function xpLookForHfIndex(ByVal str_myCompound As String) As Variant

Dim i As Long
Dim j As Integer

Dim lon_myIndex As Long

lon_myIndex = 0
For i = 0 To lon_ElementsNumber

    If str_myCompound = arr_FormationData(i, 0) Then
        
        lon_myIndex = i
        Exit For
    
    End If

Next i

If lon_myIndex = 0 Then
    
    '-- Compound not found --
    xpLookForHfIndex = -1

Else

    xpLookForHfIndex = lon_myIndex
        
End If

End Function
'---------------------------------------------------------------------------------------
' Function  : xfCp_CompT
' DateTime  : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the gas heat capacity (kJ/kmol K) of the selected
'               compound at the selected temperature
' Arguments :
'               str_Component       --> Compound formula
'               str_Temperature     --> System temperature in ºC
'---------------------------------------------------------------------------------------
Public Function xfCp_CompT(ByVal str_Component As String, ByVal dbl_Temperature As Double) As Variant

Dim lon_Index As Integer
Dim dbl_Value As Double

Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim T As Double

'-- General variables are initializated --
Call xlLoadCompoundsArray

'-- Thermochem data are loaded --
Call xlLoadThermoChemData

T = dbl_Temperature + 273.15

'-- Look for Compound array index --
lon_Index = xpLookForThermoChemIndex(str_Component, T)

If lon_Index = -1 Then

    xfCp_CompT = "Element -" & str_Component & "- not found in databank"
    Exit Function
        
End If

a = arr_ThermoChemData(lon_Index, 3)
b = arr_ThermoChemData(lon_Index, 4)
c = arr_ThermoChemData(lon_Index, 5)
d = arr_ThermoChemData(lon_Index, 6)
e = arr_ThermoChemData(lon_Index, 7)
T = T / 1000

dbl_Value = a + b * T + c * (T ^ 2) + d * (T ^ 3) + e / (T ^ 2)


xfCp_CompT = dbl_Value

End Function
'---------------------------------------------------------------------------------------
' Function  : xfHº_CompT
' DateTime  : 10/10/2017
' Author    : José García Herruzo
' Purpose   : This function returns the enthalpy (kJ/kmol) of the selected
'               compound at the selected temperature
' Arguments :
'               str_Component       --> Compound formula
'               str_Temperature     --> System temperature in ºC
'---------------------------------------------------------------------------------------
Public Function xfHº_CompT(ByVal str_Component As String, ByVal dbl_Temperature As Double) As Variant

Dim lon_Index As Integer
Dim lon_HIndex As Integer
Dim dbl_Value As Double

Dim a As Double
Dim b As Double
Dim c As Double
Dim d As Double
Dim e As Double
Dim f As Double
Dim g As Double
Dim h As Double
Dim T As Double

Dim Hf As Double

'-- General variables are initializated --
Call xlLoadCompoundsArray

'-- Databank are loaded --
Call xlLoadThermoChemData
Call xlLoadFormationData

T = dbl_Temperature + 273.15

'-- Look for Compound array index --
lon_Index = xpLookForThermoChemIndex(str_Component, T)

If lon_Index = -1 Then

    xfHº_CompT = "Element -" & str_Component & "- not found in databank"
    Exit Function
        
End If

'-- Look for enthalpy of formation value array index --
lon_HIndex = xpLookForHfIndex(str_Component)

If lon_HIndex = -1 Then

    xfHº_CompT = "Element -" & str_Component & "- not found in databank"
    Exit Function
        
End If

a = arr_ThermoChemData(lon_Index, 3)
b = arr_ThermoChemData(lon_Index, 4)
c = arr_ThermoChemData(lon_Index, 5)
d = arr_ThermoChemData(lon_Index, 6)
e = arr_ThermoChemData(lon_Index, 7)
f = arr_ThermoChemData(lon_Index, 8)
g = arr_ThermoChemData(lon_Index, 9)
h = arr_ThermoChemData(lon_Index, 10)
T = T / 1000

Hf = arr_FormationData(lon_HIndex, 1)

dbl_Value = Hf + (a * T + b * (T ^ 2) / 2 + c * (T ^ 3) / 3 + d * (T ^ 4) / 4 - e / T + f - h) * 1000

'-- Convert kJ/mol to kJ/kmol --
xfHº_CompT = dbl_Value

End Function


