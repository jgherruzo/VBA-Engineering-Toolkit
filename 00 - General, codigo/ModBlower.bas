Attribute VB_Name = "ModBlower"
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
' Module      : ModBlower
' DateTime    : 16/04/2019
' Author      : José García Herruzo, based on Jose Maria Tejera excel:
'                    "Calculos caudal y rendimiento Soplante Acido.xls"
' Purpose     : This module contents functions and procedures for blower efficiencies
'                   Calculation
' References  : N/A
' Requirements:
'               01-ModConvert
'               01-ModProperties_Element
'               01-ModProperties_ThermoChemical
'               01-ModWaterSteamTables
' Functions   :
'               01-xfAvgCpCalculation
'               02-xfAvgRhoCalculation
'               03-xfMotorEfficiency
'               04-xfBlowerEfficiency_PiTiPoToPSO2O2CO2H2O
'               05-xfBlowerFlowrate_PiTiPoToPSO2O2CO2H2O
' Procedures  :
'               01-xpCoreCalculation
' Updates     :
'       DATE        USER    DESCRIPTION
'       N/A
'-----------------------------------------------------------------------------------------

Public Const R1 = 8.134472          ' J/mol K
Public Const R2 = 0.08205746        ' atm l / mol K
Public Const Tnormal = 273.15       ' K
Public Const Pnormal = 0.98692327   ' atm

Dim dbl_IsoEff As Double
Dim dbl_Flowrate As Double
'---------------------------------------------------------------------------------------
' Procedure : xpCoreCalculation
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This procedure develop core calculation to determinate blower efficiencies
' Arguments :
'               dbl_Pin             --> Suction presure
'               dbl_Tin             --> Suction Temperature
'               dbl_Pout            --> Discharge presure
'               dbl_Tout            --> Discharge Temperature
'               dbl_Power           --> Power consumption
'               dbl_SO2             --> Composition
'               dbl_O2              --> Composition
'               dbl_CO2             --> Composition
'               dbl_H2O             --> Composition
'---------------------------------------------------------------------------------------
Private Sub xpCoreCalculation(ByVal dbl_Pin As Double, ByVal dbl_Tin As Double, ByVal dbl_Pout As Double, ByVal dbl_Tout As Double, _
                                ByVal dbl_Power As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_CO2 As Double, ByVal dbl_H2O As Double)

Dim dbl_N2 As Double

Dim dbl_AvgCv As Double
Dim dbl_AvgCp As Double             ' kJ/kmol K
Dim dbl_k As Double
Dim dbl_AvgMW As Double             ' kg/kmol
Dim dbl_AvgNRho As Double           ' Kg/Nm3
Dim dbl_AvgRho As Double            ' Kg/m3

Dim dbl_Temp As Double
Dim dbl_Y As Double

Dim dbl_Losses1 As Double
Dim dbl_Losses2 As Double

Dim dbl_RealPower As Double

'-- Average Cp for suction conditions is determined --
dbl_N2 = 1 - dbl_SO2 - dbl_O2 - dbl_CO2 - dbl_H2O

dbl_AvgCp = xfAvgCpCalculation(dbl_Tin, dbl_Tout, dbl_SO2, dbl_O2, dbl_CO2, dbl_H2O, dbl_N2)

'-- Average MW is determined --
dbl_AvgMW = xfMW("SO2") * dbl_SO2 + xfMW("O2") * dbl_O2 + xfMW("CO2") * dbl_CO2 + xfMW("H2O") * dbl_H2O + xfMW("N2") * dbl_N2

'-- Average Rho for suction conditions is determined --
dbl_AvgNRho = xfAvgRhoCalculation(dbl_SO2, dbl_O2, dbl_CO2, dbl_H2O, dbl_N2)

'-- Cv and k are determined --
dbl_AvgCv = dbl_AvgCp - (R1 / dbl_AvgMW)
dbl_k = dbl_AvgCp / dbl_AvgCv

'-- operation condition factor is determined --
dbl_Temp = ((dbl_Pout / dbl_Pin) ^ ((dbl_k - 1) / dbl_k)) - 1
dbl_Y = dbl_AvgCp * 1000 * (Tnormal + dbl_Tin) * dbl_Temp

'-- Isoentrophic efficiency is determined --
dbl_IsoEff = dbl_Y / dbl_AvgCp / 1000 / (dbl_Tout - dbl_Tin)

'-- Power calculations --
dbl_Losses1 = 15 + 0.0015 * dbl_Power
dbl_Losses2 = dbl_Power * (1 - xfMotorEfficiency(dbl_Power))
dbl_RealPower = dbl_Power - dbl_Losses1 - dbl_Losses2

'-- Flowrate calculation --
dbl_AvgRho = dbl_AvgNRho * Tnormal * ((dbl_Pin / 1000)) / (Tnormal + dbl_Tin)
dbl_Flowrate = (dbl_RealPower * 1000 * dbl_IsoEff / dbl_Y / dbl_AvgRho) * 3600 * (Tnormal * ((dbl_Pin / 1000)) / (Tnormal + dbl_Tin))

End Sub
'---------------------------------------------------------------------------------------
' Procedure : xfAvgCpCalculation
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This Function returns Avg Cp for given condition and composition
' Arguments :
'               dbl_Tin             --> Suction Temperature
'               dbl_Tout            --> Discharge Temperature
'               dbl_SO2             --> Composition
'               dbl_O2              --> Composition
'               dbl_CO2             --> Composition
'               dbl_H2O             --> Composition
'---------------------------------------------------------------------------------------
Private Function xfAvgCpCalculation(ByVal dbl_Tin As Double, ByVal dbl_Tout As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_CO2 As Double, ByVal dbl_H2O As Double, ByVal dbl_N2 As Double) As Double

Dim dbl_CpSO2 As Double
Dim dbl_CpO2 As Double
Dim dbl_CpCO2 As Double
Dim dbl_CpH2O As Double
Dim dbl_CpN2 As Double

Dim dbl_Temp As Double

'-- Average Cp for suction conditions is determined --

dbl_CpSO2 = xfCp_CompT("SO2", dbl_Tin) / xfMW("SO2")
dbl_CpO2 = xfCp_CompT("O2", dbl_Tin) / xfMW("O2")
dbl_CpCO2 = xfCp_CompT("CO2", dbl_Tin) / xfMW("CO2")
dbl_CpH2O = Cp_pT(1, dbl_Tin)
dbl_CpN2 = xfCp_CompT("N2", dbl_Tin) / xfMW("N2")

dbl_Temp = dbl_CpSO2 * dbl_SO2 + dbl_CpO2 * dbl_O2 + dbl_CpCO2 * dbl_CO2 + dbl_CpH2O * dbl_H2O + dbl_CpN2 * dbl_N2

dbl_CpSO2 = xfCp_CompT("SO2", dbl_Tout) / xfMW("SO2")
dbl_CpO2 = xfCp_CompT("O2", dbl_Tout) / xfMW("O2")
dbl_CpCO2 = xfCp_CompT("CO2", dbl_Tout) / xfMW("CO2")
dbl_CpH2O = Cp_pT(1, dbl_Tout)
dbl_CpN2 = xfCp_CompT("N2", dbl_Tout) / xfMW("N2")

dbl_Temp = (dbl_Temp + dbl_CpSO2 * dbl_SO2 + dbl_CpO2 * dbl_O2 + dbl_CpCO2 * dbl_CO2 + dbl_CpH2O * dbl_H2O + dbl_CpN2 * dbl_N2) / 2

xfAvgCpCalculation = dbl_Temp

End Function
'---------------------------------------------------------------------------------------
' Procedure : xfAvgRhoCalculation
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This Function returns Avg Rho for given condition and composition
' Arguments :
'               dbl_SO2             --> Composition
'               dbl_O2              --> Composition
'               dbl_CO2             --> Composition
'               dbl_H2O             --> Composition
'---------------------------------------------------------------------------------------
Private Function xfAvgRhoCalculation(ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_CO2 As Double, ByVal dbl_H2O As Double, ByVal dbl_N2 As Double) As Double

Dim dbl_RhoSO2 As Double
Dim dbl_RhoO2 As Double
Dim dbl_RhoCO2 As Double
Dim dbl_RhoH2O As Double
Dim dbl_RhoN2 As Double

Dim dbl_ConvFactor As Double

'-- Average Rho for suction conditions is determined --
dbl_ConvFactor = R2 * Tnormal / Pnormal ' Nl/mol ~ Nm3/kmol

dbl_RhoSO2 = xfMW("SO2") / dbl_ConvFactor
dbl_RhoO2 = xfMW("O2") / dbl_ConvFactor
dbl_RhoCO2 = xfMW("CO2") / dbl_ConvFactor
dbl_RhoH2O = xfMW("H2O") / dbl_ConvFactor
dbl_RhoN2 = xfMW("N2") / dbl_ConvFactor

xfAvgRhoCalculation = dbl_RhoSO2 * dbl_SO2 + dbl_RhoO2 * dbl_O2 + dbl_RhoCO2 * dbl_CO2 + dbl_RhoH2O * dbl_H2O + dbl_RhoN2 * dbl_N2

End Function

'---------------------------------------------------------------------------------------
' Procedure : xfMotorEfficiency
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This Function returns motor efficiency for a given consumption
' Arguments :
'               dbl_Power           --> Consumption
'---------------------------------------------------------------------------------------
Private Function xfMotorEfficiency(ByVal dbl_Power As Double) As Double

Const eA = 0.9761
Const cA = 3200
Const eB = 0.9777
Const cB = 2000
Const eC = 0.9755
Const cC = 2500

Dim dbl_Value As Double

If dbl_Power > cB Then

    dbl_Value = eA + (eB - eA) * (dbl_Power - cA) / (cB - cA)

Else

    dbl_Value = eB + (eC - eB) * (dbl_Power - cB) / (cC - cB)
    
End If

xfMotorEfficiency = dbl_Value

End Function

'---------------------------------------------------------------------------------------
' Procedure : xfBlowerEfficiency_PiTiPoToPSO2O2CO2H2O
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This function get blower efficiencies
' Arguments :
'               dbl_Pin             --> Suction presure
'               dbl_Tin             --> Suction Temperature
'               dbl_Pout            --> Discharge presure
'               dbl_Tout            --> Discharge Temperature
'               dbl_Power           --> Power consumption
'               dbl_SO2             --> Composition
'               dbl_O2              --> Composition
'               dbl_CO2             --> Composition
'               dbl_H2O             --> Composition
'---------------------------------------------------------------------------------------
Public Function xfBlowerEfficiency_PiTiPoToPSO2O2CO2H2O(ByVal dbl_Pin As Double, ByVal dbl_Tin As Double, ByVal dbl_Pout As Double, ByVal dbl_Tout As Double, _
                                ByVal dbl_Power As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_CO2 As Double, ByVal dbl_H2O As Double) As Double
                                
Call xpCoreCalculation(dbl_Pin, dbl_Tin, dbl_Pout, dbl_Tout, dbl_Power, dbl_SO2, dbl_O2, dbl_CO2, dbl_H2O)

xfBlowerEfficiency_PiTiPoToPSO2O2CO2H2O = dbl_IsoEff

End Function

'---------------------------------------------------------------------------------------
' Procedure : xfBlowerFlowrate_PiTiPoToPSO2O2CO2H2O
' DateTime  : 16/04/2019
' Author    : José García Herruzo
' Purpose   : This function get blower flowrate
' Arguments :
'               dbl_Pin             --> Suction presure
'               dbl_Tin             --> Suction Temperature
'               dbl_Pout            --> Discharge presure
'               dbl_Tout            --> Discharge Temperature
'               dbl_Power           --> Power consumption
'               dbl_SO2             --> Composition
'               dbl_O2              --> Composition
'               dbl_CO2             --> Composition
'               dbl_H2O             --> Composition
'---------------------------------------------------------------------------------------
Public Function xfBlowerFlowrate_PiTiPoToPSO2O2CO2H2O(ByVal dbl_Pin As Double, ByVal dbl_Tin As Double, ByVal dbl_Pout As Double, ByVal dbl_Tout As Double, _
                                ByVal dbl_Power As Double, ByVal dbl_SO2 As Double, ByVal dbl_O2 As Double, ByVal dbl_CO2 As Double, ByVal dbl_H2O As Double) As Double
                                
Call xpCoreCalculation(dbl_Pin, dbl_Tin, dbl_Pout, dbl_Tout, dbl_Power, dbl_SO2, dbl_O2, dbl_CO2, dbl_H2O)

xfBlowerFlowrate_PiTiPoToPSO2O2CO2H2O = dbl_Flowrate

End Function

