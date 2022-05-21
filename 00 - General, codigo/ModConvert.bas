Attribute VB_Name = "ModConvert"
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
' Module      : ModConvert
' DateTime    : 16/02/2016
' Author      : José García Herruzo
' Purpose     : This module contents functions to work with different units
' References  : N/A
' Requirements: N/A
' Functions   :
'               01-uT_CtoK
'               02-uT_KtoC
'               03-uT_CtoF
'               04-uT_FtoC
'               05-uT_FtoK
'               06-uT_KtoF
'               05-uT_RtoC
'               06-uT_CtoR
'               07-uT_RtoK
'               08-uT_KtoR
'               09-uT_FtoR
'               10-uT_RtoF
'               11-uE_CaltoJ
'               12-uE_JtoCal
'               13-uE_CaltoBTU
'               14-uE_BTUtoCal
'               15-uE_BTUtoJ
'               16-uE_JtoBTU
'               17-uP_PsitoBar
'               18-uP_BartoPsi
'               19-uM_kgtolb
'               20-uM_lbtokg
'               21-uP_inH2OtoPsi
'               22-uP_PsitoinH2O
'               23-uP_PatoBar
'               24-uP_BartoPa
'               25-uP_inH2OtoBar
'               26-uP_BartoinH2O
'               27-uP_inH2OtoPa
'               28-uP_PatoinH2O
'               29-uQ_gpmtocum
'               30-uQ_cumtogpm
'               31-uV_GalUSAtoCum
'               32-uV_CumtoGalUSA
'               33-uV_GalUKtoCum
'               34-uV_CumtoGalUK
'               35-uV_CuFttoCum
'               36-uV_CuCumtoFt
'               37-uL_Fttom
'               38-uL_mtoFt
'               39-uL_Fttoinch
'               40-uL_inchtoFt
'               41-uL_mmtoinch
'               42-uL_inchtomm
'               46-uV_GalUSAtoCuft
'               47-uV_CufttoGalUSA
'               48-uPower_kWtoHP
'               49-uPower_HPtokW
'               50-uP_inHgtoBar
'               51-uP_BartoinHg
'               52-uP_mH2OtoPa
'               53-uP_PatomH2O
'               54-uP_mH2OtoBar
'               55-uP_BartomH2O
' Procedures  : N/A
' Updates     :
'       DATE        USER    DESCRIPTION
'       31/03/2016  JGH     Functions from 21 to 28 are added
'       12/04/2016  JGH     Functions from 29 to 30 are added
'       20/04/2016  JGH     Functions from 31 to 34 are added
'       22/04/2016  JGH     Functions from 35 to 42 are added
'       20/06/2016  JGH     Functions from 46 to 47 are added
'       27/06/2016  JGH     Functions from 48 to 49 are added
'       05/06/2017  JGH     Functions from 51 to 52 are added
'       27/09/2017  JGH     Functions from 53 to 55 are added
'       12/01/2018  JGH     Function 41/42 are modified
'-----------------------------------------------------------------------------------------
'=========================================================================================
'====================== Temperature conversion funcions ==================================
'=========================================================================================
Public Function uT_CtoK(ByVal dbl_T As Double) As Double
    uT_CtoK = dbl_T + 273
End Function
Public Function uT_KtoC(ByVal dbl_T As Double) As Double
    uT_KtoC = dbl_T - 273
End Function
Public Function uT_CtoF(ByVal dbl_T As Double) As Double
    uT_CtoF = 32 + (dbl_T * 9 / 5)
End Function
Public Function uT_FtoC(ByVal dbl_T As Double) As Double
    uT_FtoC = (dbl_T - 32) * 5 / 9
End Function
Public Function uT_KtoF(ByVal dbl_T As Double) As Double
    uT_KtoF = 32 + ((dbl_T - 273) * 9 / 5)
End Function
Public Function uT_FtoK(ByVal dbl_T As Double) As Double
    uT_FtoK = (((dbl_T) - 32) * 5 / 9) + 273
End Function
Public Function uT_RtoC(ByVal dbl_T As Double) As Double
    uT_RtoC = (dbl_T - 491.67) / (9 / 5)
End Function
Public Function uT_CtoR(ByVal dbl_T As Double) As Double
    uT_CtoR = 491.67 + dbl_T * (9 / 5)
End Function
Public Function uT_RtoK(ByVal dbl_T As Double) As Double
    uT_RtoK = ((dbl_T - 491.67) / (9 / 5)) + 273
End Function
Public Function uT_KtoR(ByVal dbl_T As Double) As Double
    uT_KtoR = 491.67 + ((dbl_T - 273) * (9 / 5))
End Function
Public Function uT_RtoF(ByVal dbl_T As Double) As Double
    uT_RtoF = ((dbl_T - 32) + 491.67)
End Function
Public Function uT_FtoR(ByVal dbl_T As Double) As Double
    uT_FtoR = ((dbl_T - 491.67) + 32)
End Function
'====================================================================================
'====================== energy conversion funcions ==================================
'====================================================================================
Public Function uE_CaltoJ(ByVal dbl_E As Double) As Double
    uE_CaltoJ = dbl_E * 4.1868
End Function
Public Function uE_JtoCal(ByVal dbl_E As Double) As Double
    uE_JtoCal = dbl_E / 4.1868
End Function
Public Function uE_BTUtoJ(ByVal dbl_E As Double) As Double
    uE_BTUtoJ = dbl_E / 0.0009478171
End Function
Public Function uE_JtoBTU(ByVal dbl_E As Double) As Double
    uE_JtoBTU = dbl_E * 0.0009478171
End Function
Public Function uE_BTUtoCal(ByVal dbl_E As Double) As Double
    uE_BTUtoCal = dbl_E / 0.003968321
End Function
Public Function uE_CaltoBTU(ByVal dbl_E As Double) As Double
    uE_CaltoBTU = dbl_E * 0.003968321
End Function
'=========================================================================================
'========================= Pressure conversion funcions ==================================
'=========================================================================================
Public Function uP_PsitoBar(ByVal dbl_P As Double) As Double
    uP_PsitoBar = dbl_P / 14.5037738007
End Function
Public Function uP_BartoPsi(ByVal dbl_P As Double) As Double
    uP_BartoPsi = dbl_P * 14.5037738007
End Function
Public Function uP_inH2OtoPsi(ByVal dbl_P As Double) As Double
    uP_inH2OtoPsi = dbl_P * 0.03612728691
End Function
Public Function uP_PsitoinH2O(ByVal dbl_P As Double) As Double
    uP_PsitoinH2O = dbl_P / 0.03612728691
End Function
Public Function uP_PatoBar(ByVal dbl_P As Double) As Double
    uP_PatoBar = dbl_P / 100000
End Function
Public Function uP_BartoPa(ByVal dbl_P As Double) As Double
    uP_BartoPa = dbl_P * 100000
End Function
Public Function uP_inH2OtoBar(ByVal dbl_P As Double) As Double
    uP_inH2OtoBar = dbl_P * 0.03612728691 / 14.5037738007
End Function
Public Function uP_BartoinH2O(ByVal dbl_P As Double) As Double
    uP_BartoinH2O = dbl_P * 14.5037738007 / 0.03612728691
End Function
Public Function uP_inH2OtoPa(ByVal dbl_P As Double) As Double
    uP_inH2OtoPa = dbl_P * 0.03612728691 / 14.5037738007 * 100000
End Function
Public Function uP_PatoinH2O(ByVal dbl_P As Double) As Double
    uP_PatoinH2O = dbl_P * 14.5037738007 / 0.03612728691 / 100000
End Function
Public Function uP_mmHgtoBar(ByVal dbl_P As Double) As Double
    uP_mmHgtoBar = dbl_P / 750.062
End Function
Public Function uP_BartommHg(ByVal dbl_P As Double) As Double
    uP_BartommHg = dbl_P * 750.062
End Function
Public Function uP_mH2OtoPa(ByVal dbl_P As Double) As Double
    uP_mH2OtoPa = dbl_P * 9806.38
End Function
Public Function uP_PatomH2O(ByVal dbl_P As Double) As Double
    uP_PatomH2O = dbl_P / 9806.38
End Function
Public Function uP_mH2OtoBar(ByVal dbl_P As Double) As Double
    uP_mH2OtoBar = dbl_P * 9806.38 / 100000
End Function
Public Function uP_BartomH2O(ByVal dbl_P As Double) As Double
    uP_BartomH2O = dbl_P / 9806.38 * 100000
End Function
'=========================================================================================
'============================ Mass conversion funcions ===================================
'=========================================================================================
Public Function uM_kgtolb(ByVal dbl_M As Double) As Double
    uM_kgtolb = dbl_M * 2.20462
End Function
Public Function uM_lbtokg(ByVal dbl_M As Double) As Double
    uM_lbtokg = dbl_M / 2.20462
End Function
'=========================================================================================
'============================ Flow conversion funcions ===================================
'=========================================================================================
Public Function uQ_gpmtocum(ByVal dbl_Q As Double) As Double
    uQ_gpmtocum = dbl_Q * 60 / 264.17
End Function
Public Function uQ_cumtogpm(ByVal dbl_Q As Double) As Double
    uQ_cumtogpm = dbl_Q * 264.17 / 60
End Function
'=========================================================================================
'=========================== Volume conversion funcions ==================================
'=========================================================================================
Public Function uV_GalUSAtoCum(ByVal dbl_V As Double) As Double
    uV_GalUSAtoCum = dbl_V / 0.26417 / 1000
End Function
Public Function uV_CumtoGalUSA(ByVal dbl_V As Double) As Double
    uV_CumtoGalUSA = dbl_V * 0.26417 * 1000
End Function
Public Function uV_GalUKtoCum(ByVal dbl_V As Double) As Double
    uV_GalUKtoCum = dbl_V / 1000 / 0.21997
End Function
Public Function uV_CumtoGalUK(ByVal dbl_V As Double) As Double
    uV_CumtoGalUK = dbl_V * 0.21997 * 1000
End Function
Public Function uV_CuFttoCum(ByVal dbl_V As Double) As Double
    uV_CuFttoCum = dbl_V / 35.315
End Function
Public Function uV_CumtoCuFt(ByVal dbl_V As Double) As Double
    uV_CumtoCuFt = dbl_V * 35.315
End Function
Public Function uV_CuFttoGalUSA(ByVal dbl_V As Double) As Double
    uV_CuFttoCum = dbl_V / 0.13368
End Function
Public Function uV_GalUSAtoCuFt(ByVal dbl_V As Double) As Double
    uV_CumtoCuFt = dbl_V * 0.13368
End Function
'=========================================================================================
'=========================== Lenth conversion funcions ==================================
'=========================================================================================
Public Function uL_Fttom(ByVal dbl_L As Double) As Double
    uL_Fttom = dbl_L / 3.2808
End Function
Public Function uL_mtoFt(ByVal dbl_L As Double) As Double
    uL_mtoFt = dbl_L * 3.2808
End Function
Public Function uL_Fttoinch(ByVal dbl_L As Double) As Double
    uL_Fttoinch = dbl_L * 12
End Function
Public Function uL_inchtoFt(ByVal dbl_L As Double) As Double
    uL_inchtoFt = dbl_L / 12
End Function
Public Function uL_mmtoinch(ByVal dbl_L As Double) As Double
    uL_mmtoinch = dbl_L / 25.4
End Function
Public Function uL_inchtomm(ByVal dbl_L As Double) As Double
    uL_inchtomm = dbl_L * 25.4
End Function
'=========================================================================================
'=========================== Power conversion funcions ==================================
'=========================================================================================
Public Function uPower_kWtoHP(ByVal dbl_Power As Double) As Double
    uPower_kWtoHP = dbl_Power * 1.34102
End Function
Public Function uPower_HPtokW(ByVal dbl_Power As Double) As Double
    uPower_HPtokW = dbl_Power / 1.34102
End Function
