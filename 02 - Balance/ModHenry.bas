Attribute VB_Name = "ModHenry"
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
' Module      : ModHenry
' DateTime    : 18/04/2016
' Author      : José García Herruzo
' Purpose     : This module contents function to give the henry constant
' References  : N/A
' Requirements: N/A
' Functions   :
'               01-kH_px_O2W             atm (mol total) / mol gas
'               02-kH_px_C2H2W           atm (mol total) / mol gas
'               03-kH_px_N2W             atm (mol total) / mol gas
'               04-kH_px_CH4W            atm (mol total) / mol gas
'               05-kH_px_C2H6W           atm (mol total) / mol gas
'               06-kH_px_C3H8W           atm (mol total) / mol gas
' Procedures  : N/A
' Updates     :
'       DATE        USER    DESCRIPTION
'       N/A
'-----------------------------------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NOTES >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   1º- Temperature must be specified in Kelvin
'   2º- Negative value shows an error
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'========================== Public variables declaration =================================
Dim Aij As Double
Dim Bij As Double
Dim Cij As Double
Dim Dij As Double
Dim Eij As Double
Dim lnH As Double
'=========================================================================================
'---------------------------------------------------------------------------------------
' Function  : kH_px_O2W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for o2 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_O2W(ByVal dbl_T As Double) As Double

Aij = 144.3949115
Bij = -7775.06
Cij = -18.3974
Dij = -0.00944354
Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_O2W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_O2W = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : kH_px_N2W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for N2 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_N2W(ByVal dbl_T As Double) As Double

Aij = 164.9809115
Bij = -8432.77
Cij = -21.558
Dij = -0.00843624

Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_N2W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_N2W = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : kH_px_C2H2W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for C2H2 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_C2H2W(ByVal dbl_T As Double) As Double

Aij = 156.5089115
Bij = -8160.13
Cij = -21.4022
Dij = 0

Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_C2H2W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_C2H2W = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : kH_px_C2H6W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for C2H6 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_C2H6W(ByVal dbl_T As Double) As Double

Aij = 268.4139115
Bij = -13368.1
Cij = -37.5523
Dij = 0.00230129
Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_C2H6W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_C2H6W = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : kH_px_C3H8W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for C3H8 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_C3H8W(ByVal dbl_T As Double) As Double

Aij = 316.4579115
Bij = -15921.1
Cij = -44.3241
Dij = 0
Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_C3H8W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_C3H8W = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : kH_px_CH4W
' DateTime  : 18/04/2016
' Author    : José García Herruzo
' Purpose   : Returns henry constant for CH4 water system at specified temperature
' Arguments :
'               dbl_T       --> Dry bulbe temperature in kelvin
'---------------------------------------------------------------------------------------
Public Function kH_px_CH4W(ByVal dbl_T As Double) As Double

Aij = 183.7679115
Bij = -9111.67
Cij = -25.0379
Dij = 0.000143434
Eij = 0

lnH = Aij + (Bij / dbl_T) + (Cij * Log(dbl_T)) + (Dij * dbl_T) + (Eij / (dbl_T ^ 2))

kH_px_CH4W = Exp(lnH)

Exit Function
myhandler:
    '-- -1 is not possible --
    kH_px_CH4W = -1
    
End Function
