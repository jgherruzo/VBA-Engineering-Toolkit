Attribute VB_Name = "ModAPE"
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
' Module    : ModAPE
' DateTime  : 06/04/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures special to work with Aspen
'               from for APE file
' References: Aspen Plus GUI XX.X Type Library
' Functions : N/A
' Procedures:
'               1-RunAPE
'               2-PutAPEInputValue
'               3-GetAPEOutputValue
'               4-RedimmyVarNumber
'               5-PropertiesExtractionStatus
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/22/2013  JGH     PropertiesExtractionStatus procedure is added
'       08/28/2013  JGH     Thermal conductivity parameter is added
'----------------------------------------------------------------------------------------

Public MixedRange() As Variant
Public SaltRange() As Variant
Public SolidRange() As Variant

Dim MixLimit As Variant
Dim SaltLimit As Variant
Dim SolidLimit As Variant

Dim lb_pH As Boolean

Dim int_SaltPointer As Integer
'---------------------------------------------------------------------------------------
' Procedure : RunAPE
' DateTime  : 06/04/2013
' Author    : José García Herruzo
' Purpose   : Launch simulation applied to APE app
' Arguments :
'             myArea                    --> Area to be simulated
'             mywb                      --> Workbook which will be used as stream
'                                           source and properties container
'---------------------------------------------------------------------------------------
Public Sub RunAPE(ByVal myArea As String, ByVal myWb As Workbook)

Dim Sheet As Worksheet
Dim myColumnNumber As Integer
Dim i As Integer
Dim myString As String
Dim lb_solid As Boolean
Dim int_TotalSheet As Integer
Dim int_CounterSheet As Integer

int_TotalSheet = 0
int_CounterSheet = 0
lb_IsVisible = True
Call RedimmyVarNumber(myArea)

If SolidLimit = 0 Then

    lb_solid = False

Else

    lb_solid = True
    
End If

For Each Sheet In Worksheets
    
    If InStr(1, Sheet.Name, "-NT-", vbBinaryCompare) <> 0 Then
        
        int_TotalSheet = int_TotalSheet + 1
        
    End If

Next

For Each Sheet In Worksheets
    
    If InStr(1, Sheet.Name, "-NT-", vbBinaryCompare) <> 0 Then
                
                int_CounterSheet = int_CounterSheet + 1
        Call UpdateP2("PFD " & int_CounterSheet & " to " & int_TotalSheet, int_CounterSheet, int_TotalSheet)
        
        myColumnNumber = ReturnColumn(myWb.Name, Sheet.Name, "D1")
        
        For i = 0 To myColumnNumber - 1
                
            Call UpdateP3("Stream " & i + 1 & " to " & myColumnNumber, i + 1, myColumnNumber)
            go_simulation.SuppressDialogs = True
                
            Call PutAPEInputValue(go_simulation, Sheet, i, lb_solid)

            ' -- run the simulation --
            Application.DisplayAlerts = False
            go_simulation.Engine.Reinit
            go_simulation.SuppressDialogs = False
            go_simulation.Engine.Run
                    
            Call GetAPEOutputValue(go_simulation, Sheet, i, myWb, lb_solid, lb_pH)

        Next i
    
    End If
    
Next Sheet

End Sub
'---------------------------------------------------------------------------------------
' Procedure : PutAPEInputValue
' DateTime  : 06/04/2013
' Author    : José García Herruzo
' Purpose   : Load simulation input values
' Arguments :
'             ao_Simulation             --> Variable which contents Aspen instance
'             myWS                      --> Worksheet where seacrh streams
'             myCounter                 --> Column offset to D column
'             lb_mysolid                --> True if the stream contents solid
'                                           components
'---------------------------------------------------------------------------------------
Public Sub PutAPEInputValue(ao_Simulation As IHapp, ByVal myws As Worksheet, ByVal myCounter As Integer, _
                            ByVal lb_mysolid As Boolean)

Dim lo_StreamNodeCol As IHNodeCol
Dim lo_DummyNode As IHNode
Dim counter As Integer
Dim Bolband As Boolean

counter = 0
Bolband = False

' -- Stream conditions --
ao_Simulation.Tree.Data.Blocks.Elements("B1").Input.Elements("TEMP").Value = myws.Range(RA_TEMP).Offset(0, myCounter).Value

If IsNumeric(myws.Range(RA_PRES).Offset(0, myCounter).Value) = True Then

    ao_Simulation.Tree.Data.Blocks.Elements("B1").Input.Elements("PRES").Value = myws.Range(RA_PRES).Offset(0, myCounter).Value
    
Else

    ao_Simulation.Tree.Data.Blocks.Elements("B1").Input.Elements("PRES").Value = "2"

End If

'-- Component values --
Set lo_StreamNodeCol = ao_Simulation.Tree.Data.Streams.Elements("1").Input.FLOW.MIXED.Elements

For Each lo_DummyNode In lo_StreamNodeCol
    
    If counter <= MixLimit Then
        
        lo_DummyNode.Value = myws.Range(MixedRange(counter)).Offset(0, myCounter).Value

    Else
        
       lo_DummyNode.Value = ""
    
    End If
    
    counter = counter + 1
    
Next

If lb_mysolid = True Then

    '-- Component values --
    Set lo_StreamNodeCol = ao_Simulation.Tree.Data.Streams.Elements("1").Input.FLOW.CISOLID.Elements
    counter = 0
    
    For Each lo_DummyNode In lo_StreamNodeCol
    
        If myws.Range(SolidRange(counter)).Offset(0, myCounter).Value > 0 Then
        
            lo_DummyNode.Value = myws.Range(SolidRange(counter)).Offset(0, myCounter).Value
        
        Else
        
            lo_DummyNode.Value = ""
            
        End If
        
        If myws.Range(SolidRange(counter)).Offset(0, myCounter).Value > 0 Then
        
            Bolband = True
        
        End If
        
        counter = counter + 1
        
        If counter > SolidLimit Or SolidLimit = 1 Then
            
            Exit For
        
        End If
        
    Next

    If Bolband = True Then
    
        ao_Simulation.Tree.Data.Streams.Elements("1").Input.Elements("TEMP").CISOLID.Value = "25"
        ao_Simulation.Tree.Data.Streams.Elements("1").Input.Elements("PRES").CISOLID.Value = "1"
    
    Else
    
        ao_Simulation.Tree.Data.Streams.Elements("1").Input.Elements("PRES").CISOLID.Value = ""
        ao_Simulation.Tree.Data.Streams.Elements("1").Input.Elements("TEMP").CISOLID.Value = ""
    
    End If
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetAPEOutputValue
' DateTime  : 07/10/2013
' Author    : José García Herruzo
' Purpose   : Load simulation input values
' Arguments :
'             ao_Simulation             --> Variable which contents Aspen instance
'             myWS                      --> Worksheet where seacrh streams
'             j                         --> Column offset to D column
'             myWb                      --> Simulated balance workbook
'             lb_mysolid                --> True if the stream contents solid
'                                           components
'             lb_mypH                   --> True if the stream requeried a pH analysis
'---------------------------------------------------------------------------------------
Public Sub GetAPEOutputValue(ao_Simulation As IHapp, ByVal myws As Worksheet, ByVal j As Integer, _
                            ByVal myWb As Workbook, ByVal lb_mysolid As Boolean, ByVal lb_mypH As Boolean)
                            
Dim myMixed As Variant
Dim mySolid As Variant
Dim myX As Variant
Dim myLiquid As Variant
Dim myVapor As Variant

Dim Vweight As Variant
Dim Lweight As Variant
Dim Sweight As Variant
Dim Aweight As Variant

Dim VDensity As Variant
Dim LDensity As Variant
Dim Sdensity As Variant
Dim Adensity As Variant

Dim VCp As Variant
Dim LCp As Variant
Dim SCp As Variant
Dim ACp As Variant

Dim VEnth As Variant
Dim LEnth As Variant
Dim SEnth As Variant
Dim AEnth As Variant

Dim VK As Variant
Dim LK As Variant
Dim SK As Variant
Dim AK As Variant

Dim Vvisco As Variant
Dim Lvisco As Variant

Dim VaporPress As Variant

Dim i As Integer

Dim lo_StreamNodeCol As IHNodeCol
Dim lo_DummyNode As IHNode
Dim lo_StreamNodeCol2 As IHNodeCol
Dim lo_DummyNode2 As IHNode
Dim counter As Integer
Dim saltCounter As Integer
Dim pH25 As Variant
Dim help100 As String

If lb_mypH = True Then

    pH25 = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").Elements("PH25").MIXED.LIQUID.Value
    
        ' -- Components --
    Set lo_StreamNodeCol = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements
    
    counter = 0
    
    If int_SaltPointer = 0 Then
        
        myws.Range(RA_WATER_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("WATER").Value
        myws.Range(RA_CAUSTIC_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("CAUSTIC").Value
        saltCounter = 1
        
    ElseIf int_SaltPointer = 1 Then
        
        myws.Range(RA_WATER_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("WATER").Value
        myws.Range(RA_CAUSTIC_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("CAUSTIC").Value
        myws.Range(RA_SULF_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("SULFURIC").Value
        myws.Range(RA_AMM_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("AMOH").Value
        saltCounter = 1
        
    ElseIf int_SaltPointer = 2 Then
        
        myws.Range(RA_WATER_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("WATER").Value
        myws.Range(RA_SULF_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("SULFURIC").Value
        myws.Range(RA_AMM_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("AMOH").Value
        myws.Range(RA_CACO3_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("CACO3(S)").Value
        myws.Range(RA_GYPSUM_SALT).Offset(0, j).Value = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.MIXED.Elements("CASO4(S)").Value
        saltCounter = 4
        
    End If
    
    For Each lo_DummyNode In lo_StreamNodeCol
    
        If counter >= MixLimit + SolidLimit + saltCounter And counter <= MixLimit + SaltLimit + SolidLimit + saltCounter Then
            
            myws.Range(SaltRange(counter - saltCounter - MixLimit - SolidLimit)).Offset(0, j).Value = lo_DummyNode.Value
        
        End If
        counter = counter + 1
        
    Next
        
Else

    pH25 = ""

End If
        
'<<<< VALUES EXTRACTING>>>>>

'-- Solid --
If lb_mysolid = True Then
    
    Set lo_StreamNodeCol2 = ao_Simulation.Tree.Data.Streams.Elements("2").Output.MASSFLOW.CISOLID.Elements
    
    For Each lo_DummyNode2 In lo_StreamNodeCol2
        
        help100 = lo_DummyNode2.Name
        mySolid = mySolid + lo_DummyNode2.Value
    
    Next
    
    SEnth = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").HMX.CISOLID.Elements("SOLID").Value
    SCp = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").CPMX.CISOLID.Elements("SOLID").Value
    Sdensity = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").RHOMX.CISOLID.Elements("SOLID").Value
    Sweight = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").MWMX.CISOLID.Elements("SOLID").Value
    SK = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").KMX.CISOLID.Elements("SOLID").Value
    
    If IsNumeric(myws.Range(RA_ACF_MASS).Offset(0, j).Value) And myws.Range(RA_ACF_MASS).Offset(0, j).Value <> 0 Then
    
        mySolid = mySolid / myws.Range(RA_ACF_MASS).Offset(0, j).Value

    Else
    
        mySolid = 0
    
    End If
    
Else

    SEnth = 0
    SCp = 0
    Sdensity = 0
    Sweight = 0
    SK = 0
    mySolid = 0

End If

'-- Mixed Vapor Fracc --
myX = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("VFRAC_OUT").MIXED.Value

'-- Liquid and vapor total frac--
myVapor = myX * (1 - mySolid)

If myVapor > 1 Then

    myVapor = 1
    
End If

myLiquid = (1 - myX) * (1 - mySolid)


'-- Molecular Weigth --
    '-- Vapor --
    Vweight = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").MWMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    Lweight = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").MWMX.MIXED.Elements("LIQUID").Value
        
    '-- Average --
    Aweight = myVapor * Vweight + myLiquid * Lweight + Sweight * mySolid
    
'-- Density --
    '-- Vapor --
    VDensity = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").RHOMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    LDensity = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").RHOMX.MIXED.Elements("LIQUID").Value
    
    '-- Average --
    Adensity = myVapor * VDensity + myLiquid * LDensity + Sdensity * mySolid
    
'-- Cp --
    '-- Vapor --
    VCp = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").CPMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    LCp = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").CPMX.MIXED.Elements("LIQUID").Value
        
    '-- Average --
    ACp = myVapor * VCp + myLiquid * LCp + SCp * mySolid
    
'-- Enthalpy --
    '-- Vapor --
    VEnth = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").HMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    LEnth = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").HMX.MIXED.Elements("LIQUID").Value
    
    '-- Average --
    AEnth = myVapor * VEnth + myLiquid * LEnth + SEnth * mySolid
    
'-- Viscosidad --
    '-- Vapor --
    Vvisco = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").MUMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    Lvisco = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").MUMX.MIXED.Elements("LIQUID").Value

'-- Vapor Pressure --
    VaporPress = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").PBUB.MIXED.LIQUID.Value

'-- Thermal conductivity --
    '-- Vapor --
    VK = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").KMX.MIXED.Elements("VAPOR").Value

    '-- Liquid --
    LK = go_simulation.Tree.Data.Streams.Elements("2").Output.Elements("STRM_UPP").KMX.MIXED.Elements("LIQUID").Value
    
    '-- Average --
    AK = myVapor * VK + myLiquid * LK + SK * mySolid
'<<<< VALUES WRITING >>>>>

myws.Range(RA_AVG_MW).Offset(0, j).Value = Aweight
myws.Range(RA_VAP_FRAC).Offset(0, j).Value = myVapor
myws.Range(RA_AVG_DEN).Offset(0, j).Value = Adensity
myws.Range(RA_AVG_CP).Offset(0, j).Value = ACp
myws.Range(RA_AVG_H).Offset(0, j).Value = AEnth
myws.Range(RA_VAP_DEN).Offset(0, j).Value = VDensity
myws.Range(RA_LIQ_DEN).Offset(0, j).Value = LDensity
myws.Range(RA_SOL_DEN).Offset(0, j).Value = Sdensity
myws.Range(RA_VAP_CP).Offset(0, j).Value = VCp
myws.Range(RA_LIQ_CP).Offset(0, j).Value = LCp
myws.Range(RA_SOL_CP).Offset(0, j).Value = SCp
myws.Range(RA_VAP_VIS).Offset(0, j).Value = Vvisco
myws.Range(RA_LIQ_VIS).Offset(0, j).Value = Lvisco
myws.Range(RA_BUB_POINT).Offset(0, j).Value = VaporPress
myws.Range(RA_VAPOR_PER).Offset(0, j).Value = myVapor
myws.Range(RA_LIQUID_PER).Offset(0, j).Value = myLiquid
myws.Range(RA_SOLID_PER).Offset(0, j).Value = mySolid
myws.Range(RA_pH).Offset(0, j).Value = pH25
myws.Range(RA_AVG_K).Offset(0, j).Value = AK
myws.Range(RA_VAP_K).Offset(0, j).Value = VK
myws.Range(RA_VAP_K).Offset(0, j).Value = LK
myws.Range(RA_SOL_K).Offset(0, j).Value = SK

Call PropertiesExtractionStatus(ao_Simulation, myws.Range("D1").Offset(0, j).Value, myWb, myws)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : RedimmyVarNumber
' DateTime  : 06/04/2013
' Author    : José García Herruzo
' Purpose   : Develop the component arrays based on the area
' Arguments :
'             myArea                    --> Area to be simulated
'---------------------------------------------------------------------------------------
Private Sub RedimmyVarNumber(ByVal myArea As String)

If myArea = "01900" Then

MixLimit = 5
SaltLimit = 10
SolidLimit = 1

ReDim MixedRange(MixLimit)
ReDim SolidRange(SolidLimit)
ReDim SaltRange(SaltLimit)

MixedRange(0) = RA_WATER
MixedRange(1) = RA_GLUCOSE
MixedRange(2) = RA_NUTRIENTS
MixedRange(3) = RA_ACID_SUL
MixedRange(4) = RA_CAUSTIC
MixedRange(5) = RA_AMM_HYDR

SolidRange(0) = RA_ENZYMES

SaltRange(0) = RA_SALT1
SaltRange(1) = RA_SALT2
SaltRange(2) = RA_SALT3
SaltRange(3) = RA_SALT4
SaltRange(4) = RA_OH
SaltRange(5) = RA_H3O
SaltRange(6) = RA_NH4
SaltRange(7) = RA_HSO4
SaltRange(8) = RA_SO4
SaltRange(9) = RA_NA
SaltRange(10) = RA_NH3

int_SaltPointer = 1

lb_pH = True

ElseIf myArea = "02100" Then

MixLimit = 3
SaltLimit = 0
SolidLimit = 0

ReDim MixedRange(MixLimit)
MixedRange(0) = RA_WATER
MixedRange(1) = RA_ETHANOL
MixedRange(2) = RA_DENATUR
MixedRange(3) = RA_CORR_INH

lb_pH = False

ElseIf myArea = "04000" Or myArea = "04500" Or myArea = "05000" Or myArea = "07000" Then

MixLimit = 1
SaltLimit = 0
SolidLimit = 0

ReDim MixedRange(MixLimit)
MixedRange(0) = RA_WATER

lb_pH = False

ElseIf myArea = "06000" Then

MixLimit = 2
SaltLimit = 0
SolidLimit = 0

ReDim MixedRange(MixLimit)
MixedRange(0) = RA_WATER
MixedRange(1) = RA_OXYGEN
MixedRange(2) = RA_NITROGEN

lb_pH = False

ElseIf myArea = "07850" Then

MixLimit = 1
SaltLimit = 4
SolidLimit = 0

ReDim MixedRange(MixLimit)
ReDim SaltRange(SaltLimit)

MixedRange(0) = RA_WATER
MixedRange(1) = RA_CAUSTIC

SaltRange(0) = RA_SALT1
SaltRange(1) = RA_SALT2
SaltRange(2) = RA_OH
SaltRange(3) = RA_H3O
SaltRange(4) = RA_NA

int_SaltPointer = 0

lb_pH = True

ElseIf myArea = "12000" Then

MixLimit = 10
SaltLimit = 17
SolidLimit = 3

ReDim MixedRange(MixLimit)
ReDim SolidRange(SolidLimit)
ReDim SaltRange(SaltLimit)

MixedRange(0) = RA_WATER
MixedRange(1) = RA_GLUCOSE
MixedRange(2) = RA_XYLOSE
MixedRange(3) = RA_ACID_SUL
MixedRange(4) = RA_AMM_HYDR
MixedRange(5) = RA_FURFURAL
MixedRange(6) = RA_CACO3
MixedRange(7) = RA_GYPSUM
MixedRange(8) = RA_OXYGEN
MixedRange(9) = RA_NITROGEN
MixedRange(10) = RA_CARB_DIO

SaltRange(0) = RA_SALT3
SaltRange(1) = RA_SALT4
SaltRange(2) = RA_SALT5
SaltRange(3) = RA_SALT6
SaltRange(4) = RA_SALT7
SaltRange(5) = RA_SALT8
SaltRange(6) = RA_SALT9
SaltRange(7) = RA_OH
SaltRange(8) = RA_H3O
SaltRange(9) = RA_NH4
SaltRange(10) = RA_HSO4
SaltRange(11) = RA_SO4
SaltRange(12) = RA_NH3
SaltRange(13) = RA_CA
SaltRange(14) = RA_HCO3
SaltRange(15) = RA_CO3
SaltRange(16) = RA_NH2COO
SaltRange(17) = RA_CaOH

SolidRange(0) = RA_CELLULOSE
SolidRange(1) = RA_XYLAN
SolidRange(2) = RA_ASH
SolidRange(3) = RA_OTH_INS_SOL

int_SaltPointer = 2

lb_pH = True

ElseIf myArea = "16000" Or myArea = "09000" Or myArea = "19000" Then

MixLimit = 12
SaltLimit = 17
SolidLimit = 5

ReDim MixedRange(MixLimit)
ReDim SolidRange(SolidLimit)
ReDim SaltRange(SaltLimit)

MixedRange(0) = RA_WATER
MixedRange(1) = RA_ETHANOL
MixedRange(2) = RA_GLUCOSE
MixedRange(3) = RA_XYLOSE
MixedRange(4) = RA_NUTRIENTS
MixedRange(5) = RA_ACID_SUL
MixedRange(6) = RA_AMM_HYDR
MixedRange(7) = RA_FURFURAL
MixedRange(8) = RA_CACO3
MixedRange(9) = RA_GYPSUM
MixedRange(10) = RA_OXYGEN
MixedRange(11) = RA_NITROGEN
MixedRange(12) = RA_CARB_DIO

SaltRange(0) = RA_SALT3
SaltRange(1) = RA_SALT4
SaltRange(2) = RA_SALT5
SaltRange(3) = RA_SALT6
SaltRange(4) = RA_SALT7
SaltRange(5) = RA_SALT8
SaltRange(6) = RA_SALT9
SaltRange(7) = RA_OH
SaltRange(8) = RA_H3O
SaltRange(9) = RA_NH4
SaltRange(10) = RA_HSO4
SaltRange(11) = RA_SO4
SaltRange(12) = RA_NH3
SaltRange(13) = RA_CA
SaltRange(14) = RA_HCO3
SaltRange(15) = RA_CO3
SaltRange(16) = RA_NH2COO
SaltRange(17) = RA_CaOH

SolidRange(0) = RA_CELLULOSE
SolidRange(1) = RA_XYLAN
SolidRange(2) = RA_ASH
SolidRange(3) = RA_ENZYMES
SolidRange(4) = RA_ORGANISMS
SolidRange(5) = RA_OTH_INS_SOL

int_SaltPointer = 2

lb_pH = True

ElseIf myArea = "18000" Then

MixLimit = 12
SaltLimit = 0
SolidLimit = 5

ReDim MixedRange(MixLimit)
ReDim SolidRange(SolidLimit)
ReDim SaltRange(SaltLimit)

MixedRange(0) = RA_WATER
MixedRange(1) = RA_ETHANOL
MixedRange(2) = RA_GLUCOSE
MixedRange(3) = RA_XYLOSE
MixedRange(4) = RA_NUTRIENTS
MixedRange(5) = RA_ACID_SUL
MixedRange(6) = RA_AMM_HYDR
MixedRange(7) = RA_FURFURAL
MixedRange(8) = RA_CACO3
MixedRange(9) = RA_GYPSUM
MixedRange(10) = RA_OXYGEN
MixedRange(11) = RA_NITROGEN
MixedRange(12) = RA_CARB_DIO

SolidRange(0) = RA_CELLULOSE
SolidRange(1) = RA_XYLAN
SolidRange(2) = RA_ASH
SolidRange(3) = RA_ENZYMES
SolidRange(4) = RA_ORGANISMS
SolidRange(5) = RA_OTH_INS_SOL

int_SaltPointer = 10

lb_pH = False

End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PropertiesExtractionStatus
' DateTime  : 07/22/2013
' Author    : José García Herruzo
' Purpose   : Extract stream simulation status in special ws
' Arguments :
'             ao_Simulation             --> Variable which contents Aspen instance
'             strStream                 --> Simulated stream
'             mywb                      --> Simulated workbook
'---------------------------------------------------------------------------------------
Public Sub PropertiesExtractionStatus(ByVal ao_Simulation As IHapp, ByVal strStream As String, ByVal myWb As Workbook, ByVal SourceWs As Worksheet)

Dim lastrow As Integer
Dim myws As Worksheet
Dim lo_Block As IHNode
Dim dateToday As Variant

dateToday = Day(Now) & "\" & Month(Now) & "\" & Year(Now)

If SheetExist(myWb.Name, "Aspen Log") = False Then

    Call AddNewSheet(myWb.Name, "Aspen Log")
    myWb.Worksheets("Aspen Log").Range("A1").Value = "Stream ID"
    myWb.Worksheets("Aspen Log").Range("B1").Value = "PFD"
    myWb.Worksheets("Aspen Log").Range("C1").Value = "Sim Status"
    myWb.Worksheets("Aspen Log").Range("D1").Value = "Block ID"
    myWb.Worksheets("Aspen Log").Range("E1").Value = "Block Type"
    myWb.Worksheets("Aspen Log").Range("F1").Value = "Block section"
    myWb.Worksheets("Aspen Log").Range("G1").Value = "Last update:"
    
End If
    
Set myws = myWb.Worksheets("Aspen Log")
          
If myWb.Worksheets("Aspen Log").Range("H1").Value <> dateToday Then

    myWb.Worksheets("Aspen Log").Range("H1").Value = dateToday
    myws.Range("A2:F16000").ClearContents

End If
         
lastrow = myws.Range("A16000").End(xlUp).Row
lastrow = lastrow + 1

 ' -- Stream ID --
myws.Range("A" & lastrow).Value = strStream
 ' -- retrieve block calculation status --
myws.Range("B" & lastrow).Value = SourceWs.Name
myws.Range("C" & lastrow).Value = aplGetCompStatus(ao_Simulation.Tree.Data.AttributeValue(HAP_COMPSTATUS))
    
    ' -- set intermediate collection object to simplify --
    Set lo_Block = go_simulation.Tree.Data.Blocks.Elements("B1")
        
        ' -- retrieve block name --
myws.Range("d" & lastrow).Value = lo_Block.Name
        
        ' -- retrieve block type --
myws.Range("E" & lastrow).Value = _
            lo_Block.AttributeValue(HAP_RECORDTYPE)
            
        ' -- retrieve block section (GLOBAL by default) --
myws.Range("F" & lastrow).Value = _
            lo_Block.AttributeValue(HAP_SECTION)
        
If myws.Range("C" & lastrow).Value = "Results Available" Then

    Call PaintGreen(myws.Range("C" & lastrow))

Else

    Call PaintRed(myws.Range("C" & lastrow))

End If

Set myws = Nothing

End Sub
