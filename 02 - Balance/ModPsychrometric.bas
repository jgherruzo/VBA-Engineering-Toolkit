Attribute VB_Name = "ModPsychrometric"
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
' Module      : ModPsychrometric
' DateTime    : 15/02/2016
' Author      : José García Herruzo
' Purpose     : This module contents function and procedures to solve psychrometric problems
' References  : N/A
' Requirements:
'               1-ModWaterSteamTables
' Functions   :
'               1-Hsat_PT           kgw/kg dry air
'               2-Habs_HrHsat       kgw/kg dry air
'               3-Habs_PTHr         kgw/kg dry air
'               4-airH_THabs        kcal/kg
'               5-airH_TPHr         kcal/kg
'               6-airH_THrHsat      kcal/kg
'               7-Twb_PairH         ºC
'               8-Twb_PTHabs        ºC
'               9-Twb_PTHr          ºC
'               10-Twb_PTHrHsat     ºC
'               11-AirMoi_Habs      kgw/kg total
'               12-AirMoi_PTHr      kgw/kg total
'               13-AirMoi_HrHsat    kgw/kg total
'               14-Pw_THr           bar
'               15-Tdew_THr         ºC
'               16-Hr_HsatHabs      kgw/kg dry air
'               17-Hr_PTHabs        kgw/kg dry air
'               18-Hr_PTairH        kgw/kg dry air
'               19-Hr_PTAirMoi      kgw/kg dry air
'               20-Hr_TTdew         kgw/kg dry air
'               21-T_HabsairH       ºC
'               22-Hsat_PTMW          kgw/kg dry gas
' Procedures  : N/A
' Updates     :
'       DATE        USER    DESCRIPTION
'       21/09/2017  JGH     Function 22 is added in order to determinate HS for any gas
'-----------------------------------------------------------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<< NOTES >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'   1º- Negative values as result of any funtion means that there is any error
'   2º- A Twb of 10000 means that the loop did not finish. Add any initial twb as last
'       argument if you are using function 7. Author advice: Use dry bulbe temperature
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'========================== Public variables declaration =================================
Const PMw As Double = 18.01528            '-- Water molecular weight --
Const PMa As Double = 28.9420982524272    '-- Air molecular weight --
Const Cpa As Double = 0.24                '-- Air Average specific heat --
Const Cpw As Double = 0.46                '-- Water Average specific heat --
Const VHw As Double = 595                 '-- Water vaporization heat at 0ºC --
'=========================================================================================

'---------------------------------------------------------------------------------------
' Function  : Hsat_PT
' DateTime  : 15/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns saturation humidity at the specified absolute
'               pressure and dry bulbe temperature
' Arguments :
'               dbl_P       --> System absolute pressure
'               dbl_T       --> Dry bulbe temperature/System temperature
'---------------------------------------------------------------------------------------
Public Function Hsat_PT(ByVal dbl_P As Double, ByVal dbl_T As Double) As Double

Dim lon_satP As Double '-- Saturation pressure at T --
Dim my_Value As Double

On Error GoTo myhandler

lon_satP = psat_T(dbl_T)

my_Value = (PMw * lon_satP) / ((dbl_P - lon_satP) * PMa)

Hsat_PT = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    Hsat_PT = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Habs_HrHsat
' DateTime  : 15/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns saturation humidity at the specified saturation
'               humidity and relative humidity
' Arguments :
'               dbl_Hr       --> Relative humidity
'               dbl_Hsat     --> Saturation humidity
'---------------------------------------------------------------------------------------
Public Function Habs_HrHsat(ByVal dbl_Hr As Double, ByVal dbl_Hsat As Double) As Double

Dim my_Value As Double

On Error GoTo myhandler

my_Value = dbl_Hr * dbl_Hsat

Habs_HrHsat = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    Habs_HrHsat = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Habs_PTHr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns saturation humidity at the specified saturation
'               humidity and relative humidity
' Arguments :
'               dbl_Hr       --> Relative humidity
'               dbl_P       --> System absolute pressure
'               dbl_T       --> Dry bulbe temperature/System temperature
'---------------------------------------------------------------------------------------
Public Function Habs_PTHr(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Hr As Double) As Double

Dim my_Value As Double
Dim dbl_myHsat As Double

On Error GoTo myhandler

dbl_myHsat = Hsat_PT(dbl_P, dbl_T)
my_Value = dbl_Hr * dbl_myHsat

Habs_PTHr = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    Habs_PTHr = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : airH_THabs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns enthalpy of the air at the specified absolute
'               humidity and system temperature and pressure
' Arguments :
'               dbl_T       --> Dry bulbe temperature/System temperature
'               dbl_Habs    --> Absolute humidity
'---------------------------------------------------------------------------------------
Public Function airH_THabs(ByVal dbl_T As Double, ByVal dbl_Habs As Double) As Double

Dim my_Value As Double

On Error GoTo myhandler

my_Value = Cpa * dbl_T + dbl_Habs * (VHw + Cpw * dbl_T)

airH_THabs = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    airH_THabs = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : airH_TPHr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns enthalpy of the air at the specified relative
'               humidity and system temperature and pressure
' Arguments :
'               dbl_T       --> Dry bulbe temperature/System temperature
'               dbl_P       --> System absolute pressure
'               dbl_Hr       --> Relative humidity
'---------------------------------------------------------------------------------------
Public Function airH_TPHr(ByVal dbl_T As Double, ByVal dbl_P As Double, ByVal dbl_Hr As Double) As Double

Dim my_Value As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = Habs_PTHr(dbl_P, dbl_T, dbl_Hr)
my_Value = Cpa * dbl_T + dbl_myHabs * (VHw + Cpw * dbl_T)

airH_TPHr = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    airH_TPHr = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : airH_THrHsat
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns enthalpy of the air at the specified absolute
'               humidity and system temperature
' Arguments :
'               dbl_T        --> Dry bulbe temperature/System temperature
'               dbl_Hr       --> Relative humidity
'               dbl_Hsat     --> Saturation humidity
'---------------------------------------------------------------------------------------
Public Function airH_THrHsat(ByVal dbl_T As Double, ByVal dbl_Hr As Double, ByVal dbl_Hsat As Double) As Double

Dim my_Value As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = Habs_HrHsat(dbl_Hr, dbl_Hsat)
my_Value = Cpa * dbl_T + dbl_myHabs * (VHw + Cpw * dbl_T)

airH_THrHsat = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    airH_THrHsat = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Twb_PairH
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns wet bulbe temperature at the specified enthalpy and
'               system pressure
' Arguments :
'               dbl_airH    --> Air enthalpy
'               dbl_P       --> System absolute pressure
'---------------------------------------------------------------------------------------
Public Function Twb_PairH(ByVal dbl_P As Double, ByVal dbl_airH As Double, Optional dbl_initialTwb As Double) As Double

Dim dbl_myT As Double
Dim dbl_myHsat As Double
Dim dbl_myairH As Double
Dim dbl_myError As Double
Dim lon_Counter As Long

On Error GoTo myhandler

'-- Set first iteration temperature at 25ºC and any error <> 0 --
'-- To avoid problems with value 0, Twb is solved using K and then retunrs C --
If dbl_initialTwb <> 0 Then

    dbl_myT = dbl_initialTwb + 273

Else

    dbl_myT = 298

End If

dbl_myError = 1
'-- per each loop --
Do While dbl_myError <> 0
    
    dbl_myHsat = Hsat_PT(dbl_P, dbl_myT - 273)
    dbl_myairH = airH_THabs(dbl_myT - 273, dbl_myHsat)
    
    If dbl_airH < 0 And dbl_airH > dbl_myairH And dbl_myairH < 0 Then
    
        dbl_myError = ((Abs(dbl_airH) - Abs(dbl_myairH)) / (dbl_airH)) * 0.01
    
    Else
                    
        If dbl_myairH < 1 Then
            
            If dbl_initialTwb < 0 And dbl_initialTwb > -2 Then
            
                dbl_myError = ((Abs(dbl_airH) - Abs(dbl_myairH)) / Abs(dbl_airH))
    
            Else
            
                dbl_myError = ((Abs(dbl_airH) - Abs(dbl_myairH)) / Abs(dbl_airH)) * 0.01
                
            End If
            
        Else
        
            dbl_myError = (Abs(dbl_airH) - Abs(dbl_myairH)) / Abs(dbl_airH)
            
        End If
        
    End If
        
    If dbl_myError > 1 Then

        dbl_myError = 1

    ElseIf dbl_myError < -1 Then

        dbl_myError = -0.9999

    End If
    
    '-- Establish the new temperature --7
    If lon_Counter < 1000 Then
    
        dbl_myT = dbl_myT + dbl_myT * dbl_myError * 0.01
    
    Else
    
        dbl_myT = dbl_myT + dbl_myT * dbl_myError * 0.001
        
    End If
    'dbl_myT = T_HabsairH(dbl_myHsat, dbl_myairH)

    lon_Counter = lon_Counter + 1
    If lon_Counter > 2000 Then
                    
        '-- 10000 is not possible, means that you must to specified a better T --
        If dbl_myError < 0.0000000001 And dbl_myError > -0.0000000001 Then
        
            Exit Do
            
        Else
        
            Twb_PairH = 10000
            Exit Function
        
        End If
        
    End If
    
Loop

Twb_PairH = dbl_myT - 273
Exit Function
myhandler:
    '-- -1000 is not possible --
    Twb_PairH = -1000
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Twb_PTHabs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns wet bulbe temperature at the specified air absolute
'                humidity and system pressure and temperature
' Arguments :
'               dbl_Habs    --> Air absolute humidity
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Twb_PTHabs(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Habs As Double) As Double

Dim dbl_myairH As Double
Dim dbl_myvalue As Double

'-- Calculate air enthalpy --
dbl_myairH = airH_THabs(dbl_T, dbl_Habs)

'-- use original function --
dbl_myvalue = Twb_PairH(dbl_P, dbl_myairH, dbl_T)

Twb_PTHabs = dbl_myvalue

End Function
'---------------------------------------------------------------------------------------
' Function  : Twb_PTHr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns wet bulbe temperature at the specified air relative
'                humidity and system pressure and temperature
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Twb_PTHr(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Hr As Double) As Double

Dim dbl_myairH As Double
Dim dbl_myvalue As Double

'-- Calculate air enthalpy --
dbl_myairH = airH_TPHr(dbl_T, dbl_P, dbl_Hr)

'-- use original function --
dbl_myvalue = Twb_PairH(dbl_P, dbl_myairH, dbl_T)

Twb_PTHr = dbl_myvalue

End Function
'---------------------------------------------------------------------------------------
' Function  : Twb_PTHrHsat
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns wet bulbe temperature at the specified air relative
'                humidity, system pressure and temperature and saturation humidity
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_Hsat    --> Air saturation humidity
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Twb_PTHrHsat(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Hr As Double _
                            , ByVal dbl_Hsat As Double) As Double

Dim dbl_myairH As Double
Dim dbl_myvalue As Double

'-- Calculate air enthalpy --
dbl_myairH = airH_THrHsat(dbl_T, dbl_Hr, dbl_Hsat)

'-- use original function --
dbl_myvalue = Twb_PairH(dbl_P, dbl_myairH, dbl_T)

Twb_PTHrHsat = dbl_myvalue

End Function
'---------------------------------------------------------------------------------------
' Function  : AirMoi_Habs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns air moisture depending on the specified absolute
'               humidity
' Arguments :
'               dbl_Habs    --> Air absolute humidity
'---------------------------------------------------------------------------------------
Public Function AirMoi_Habs(ByVal dbl_Habs As Double) As Double

Dim dbl_myvalue As Double

On Error GoTo myhandler

dbl_myvalue = 1 - 1 / (1 + dbl_Habs)

AirMoi_Habs = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    AirMoi_Habs = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : AirMoi_PTHr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns air moisture at the specified air relative
'                humidity and system pressure and temperature
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function AirMoi_PTHr(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Hr As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = Habs_PTHr(dbl_P, dbl_T, dbl_Hr)
dbl_myvalue = 1 - 1 / (1 + dbl_myHabs)

AirMoi_PTHr = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    AirMoi_PTHr = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : AirMoi_HrHsat
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns air moisture at the specified air relative
'                humidity and system pressure and temperature
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_Hsat    --> Air saturation humidity
'---------------------------------------------------------------------------------------
Public Function AirMoi_HrHsat(ByVal dbl_Hr, ByVal dbl_Hsat As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = Habs_HrHsat(dbl_Hr, dbl_Hsat)
dbl_myvalue = 1 - 1 / (1 + dbl_myHabs)

AirMoi_HrHsat = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    AirMoi_HrHsat = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Pw_THr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns vapor pressure at the specified air relative
'                humidity and system temperature
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Pw_THr(ByVal dbl_T As Double, ByVal dbl_Hr As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myPsat As Double

On Error GoTo myhandler

dbl_myPsat = psat_T(dbl_T)
dbl_myvalue = dbl_myPsat * dbl_Hr

Pw_THr = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Pw_THr = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Tdew_THr
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns vdew point at the specified air relative
'                humidity and system temperature
' Arguments :
'               dbl_Hr      --> Air relative humidity
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Tdew_THr(ByVal dbl_T As Double, ByVal dbl_Hr As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myPsat As Double

On Error GoTo myhandler

dbl_myPsat = Pw_THr(dbl_T, dbl_Hr)
dbl_myvalue = Tsat_p(dbl_myPsat)

Tdew_THr = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Tdew_THr = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hr_HsatHabs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns relative humidity at specified absolute humidity and
'               saturation humidity
' Arguments :
'               dbl_Hsat    --> Air relative humidity
'               dbl_Habs    --> Air relative humidity
'---------------------------------------------------------------------------------------
Public Function Hr_HsatHabs(ByVal dbl_Hsat As Double, ByVal dbl_Habs As Double) As Double

Dim dbl_myvalue As Double

On Error GoTo myhandler

dbl_myvalue = dbl_Habs / dbl_Hsat

Hr_HsatHabs = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Hr_HsatHabs = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hr_PTHabs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns relative humidity at specified absolute humidity and
'               system pressure and temperature
' Arguments :
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'               dbl_Habs    --> Air relative humidity
'---------------------------------------------------------------------------------------
Public Function Hr_PTHabs(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_Habs As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myHsat As Double

On Error GoTo myhandler

dbl_myHsat = Hsat_PT(dbl_P, dbl_T)
dbl_myvalue = dbl_Habs / dbl_myHsat

Hr_PTHabs = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Hr_PTHabs = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hr_PTairH
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns relative humidity at specified air enthalpy and
'               system temperature and pressure
' Arguments :
'               dbl_T       --> System absolute temperature
'               dbl_P       --> System absolute pressure
'               dbl_airH    --> Air relative humidity
'---------------------------------------------------------------------------------------
Public Function Hr_PTairH(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_airH As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = (dbl_airH - Cpa * dbl_T) / (VHw + Cpw * dbl_T)
dbl_myvalue = Hr_PTHabs(dbl_P, dbl_T, dbl_myHabs)

Hr_PTairH = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Hr_PTairH = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hr_PTAirMoi
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns relative humidity at specified air moisture and
'               system temperature and pressure
' Arguments :
'               dbl_P       --> System absolute pressure
'               dbl_T       --> System absolute temperature
'               dbl_AirMoi  --> Air moisture
'---------------------------------------------------------------------------------------
Public Function Hr_PTAirMoi(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_AirMoi As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_myHabs As Double

On Error GoTo myhandler

dbl_myHabs = (1 / (1 - dbl_AirMoi)) - 1

dbl_myvalue = Hr_PTHabs(dbl_P, dbl_T, dbl_myHabs)

Hr_PTAirMoi = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Hr_PTAirMoi = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hr_TTdew
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns relative humidity at dew point and system
'               temperature
' Arguments :
'               dbl_Tdew    --> System absolute temperature
'               dbl_T       --> System absolute temperature
'---------------------------------------------------------------------------------------
Public Function Hr_TTdew(ByVal dbl_T As Double, ByVal dbl_Tdew As Double) As Double

Dim dbl_myvalue As Double
Dim dbl_Pw As Double

On Error GoTo myhandler

dbl_Pw = psat_T(dbl_Tdew)

dbl_myvalue = dbl_Pw / psat_T(dbl_T)

Hr_TTdew = dbl_myvalue

Exit Function
myhandler:
    '-- -1 is not possible --
    Hr_TTdew = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : airH_THabs
' DateTime  : 16/02/2016
' Author    : José García Herruzo
' Purpose   : This function returns the temperature which get the specified enthalpy
'               with the specified humidity
' Arguments :
'               dbl_airH    --> Air enthalpy
'               dbl_Habs    --> Absolute humidity
'---------------------------------------------------------------------------------------
Public Function T_HabsairH(ByVal dbl_Habs As Double, ByVal dbl_airH As Double) As Double

Dim my_Value As Double

On Error GoTo myhandler

my_Value = (dbl_airH - dbl_Habs * VHw) / (Cpa + dbl_Habs * Cpw)

T_HabsairH = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    T_HabsairH = -1
    
End Function
'---------------------------------------------------------------------------------------
' Function  : Hsat_PTMW
' DateTime  : 15/02/2017
' Author    : José García Herruzo
' Purpose   : This function returns saturation humidity at the specified absolute
'               pressure, dry bulbe temperature and gas composition
' Arguments :
'               dbl_P       --> System absolute pressure
'               dbl_T       --> Dry bulbe temperature/System temperature
'               dbl_MW      --> Molecular weight of the dry gas
'---------------------------------------------------------------------------------------
Public Function Hsat_PTMW(ByVal dbl_P As Double, ByVal dbl_T As Double, ByVal dbl_MW As Double) As Double

Dim lon_satP As Double '-- Saturation pressure at T --
Dim my_Value As Double

On Error GoTo myhandler

lon_satP = psat_T(dbl_T)

my_Value = (PMw * lon_satP) / ((dbl_P - lon_satP) * dbl_MW)

Hsat_PTMW = my_Value

Exit Function
myhandler:
    '-- -1 is not possible --
    Hsat_PTMW = -1
    
End Function
