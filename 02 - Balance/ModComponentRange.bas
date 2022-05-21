Attribute VB_Name = "ModComponentRange"
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
' Module    : ModComponentRange
' DateTime  : 06/26/2013
' Author    : José García Herruzo
' Purpose   : This module contents a variable by balance parameters and its range
' References: N/A
' Functions : N/A
' Procedures:
'               1-Update_Components_Range
' Status    : OPEN
' Updates   :
'       DATE        USER    DESCRIPTION
'       06/26/2013  JGH     CaCO3 and H2SO4 with them salts are added
'       08/28/2013  JGH     New Properties
'       10/02/2013  JGH     Real viscosity is added
'       10/14/2013  JGH     Ca salts are added
'       11/04/2013  JGH     TIS & TSS ranges are added
'       11/29/2013  JGH     TS*, Non-oxidant, bisulphite, HCL, oxidant and Real density
'                           are added
'       19/12/2013  JGH     TS* is ruled out
'       20/12/2013  JGH     Oil,FE2,CU2,alcalinity,FE2mg,Cumg,Hardness, O2mg, CL, TDS,
'                           Oilmg and Mg are added
'       23/12/2013  JGH     Change alcalinity by SiO2
'       23/12/2013  JGH     Change TDS by Conductivity
'       23/12/2013  JGH     Fiber, fillers, plastics, inerts and kNa are added
'       24/01/2014  JGH     Straw components are added
'       04/02/2014  JGH     Media, Antifoam, H3PO4 and CO are added
'       03/03/2014  JGH     Cl, S, Urea, CaCl2 and line parameters are added
'       10/04/2014  JGH     Stream routing parameter are added
'       02/06/2014  JGH     Ethanol equilibrium rows are removed, experimental Cp and pH
'                           are added. Component range selector are created. NaCLO,
'                           cellobiose, lactoside, penicillin are added.
'       24/06/2014  JGH     xsB2GComponent is added
'       27/06/2014  JGH     Update_Components_Range is added again. Resto of function
'                           are deleted
'       06/05/2015  JGH     NG components are added
'----------------------------------------------------------------------------------------
'-- Design variables --
Public RA_ACF_MASS As String
Public RA_ACF_LIQ As String
Public RA_ACF_VAP As String
Public RA_DESIGN_FACTOR As String
Public RA_DF_MASS As String
Public RA_DF_LIQ As String
Public RA_DF_VAP As String
Public RA_TEMP As String
Public RA_PRES As String
Public RA_VAP_FRAC As String

'-- Liquid & Soluble Solid --
Public RA_WATER As String
Public RA_ETHANOL As String
Public RA_GLUCOSE As String
Public RA_GLU_OLIG As String
Public RA_XYLOSE As String
Public RA_XYL_OLIG As String
Public RA_ARA_OLIG As String
Public RA_GALACTOSE As String
Public RA_MANNOSE As String
Public RA_BIO_EXTR As String
Public RA_OTH_INORG As String
Public RA_NUTRIENTS As String
Public RA_OTHER_SS As String
Public RA_ACID_SUL As String
Public RA_CAUSTIC As String
Public RA_AMM_HYDR As String
Public RA_ACETIC As String
Public RA_FURFURAL As String
Public RA_HMF As String
Public RA_OTH_MET_PROD As String
Public RA_DENATUR As String
Public RA_CORR_INH As String
Public RA_FLOCULANT As String
Public RA_CACO3 As String
Public RA_GYPSUM As String
Public RA_COAGULANT As String
Public RA_NON_OXI As String
Public RA_BISULF As String
Public RA_HCL As String
Public RA_OIL As String
Public RA_ANTIFOAM As String
Public RA_H3PO4 As String
Public RA_MEDIA As String

'-- Gases --
Public RA_OXYGEN As String
Public RA_OZONE As String
Public RA_NITROGEN As String
Public RA_METHANE As String
Public RA_CARB_DIO As String
Public RA_HYDR_SULF As String
Public RA_NITRO_OXI As String
Public RA_SULFUR_DIO As String
Public RA_OTH_GAS As String
Public RA_OXI As String
Public RA_CO As String
Public RA_NACLO As String
Public RA_HNO3 As String
Public RA_ETHANE As String

'-- Salts --
Public RA_PROPANE As String
Public RA_BUTANE As String
Public RA_SULF_SALT As String
Public RA_AMM_SALT As String
Public RA_CACO3_SALT As String
Public RA_GYPSUM_SALT As String
Public RA_SALT1 As String
Public RA_SALT2 As String
Public RA_SALT3 As String
Public RA_SALT4 As String
Public RA_SALT5 As String
Public RA_SALT6 As String
Public RA_SALT7 As String
Public RA_SALT8 As String
Public RA_SALT9 As String
Public RA_OH As String
Public RA_H3O As String
Public RA_NH4 As String
Public RA_HSO4 As String
Public RA_SO4 As String
Public RA_NA As String
Public RA_NH3 As String
Public RA_CA As String
Public RA_HCO3 As String
Public RA_CO3 As String
Public RA_NH2COO As String
Public RA_CAOH As String
Public RA_Fe2 As String
Public RA_Cu2 As String
Public RA_k As String

'-- Insoluble Solid --
Public RA_CELLULOSE As String
Public RA_XYLAN As String
Public RA_GALACTAN As String
Public RA_MANNAN As String
Public RA_ARABINAN As String
Public RA_LIGNIN As String
Public RA_PROTEIN As String
Public RA_ACETATE As String
Public RA_ASH As String
Public RA_ENZYMES As String
Public RA_ORGANISMS As String
Public RA_OTH_INS_SOL As String
Public RA_ACETYL As String
Public RA_CELLOBIOSE As String
Public RA_LACTOSIDE As String
Public RA_PENICILIN As String

'-- Waste components --
Public RA_ORG_FRAC As String
Public RA_TEXT As String
Public RA_PAP_CARD As String
Public RA_PET As String
Public RA_HDPE As String
Public RA_PP As String
Public RA_FILM As String
Public RA_OTH_PLAS As String
Public RA_NON_FE As String
Public RA_FE As String
Public RA_INERT As String
Public RA_GLASS As String
Public RA_BRISCKS As String
Public RA_HAZAR_WAS As String
Public RA_ELECT_WAS As String
Public RA_WOOD As String
Public RA_GUMS As String
Public RA_BULK_WAS As String

'-- Straw --
Public RA_STRAW As String
Public RA_STRING As String
Public RA_FORGEIN As String
Public RA_DUST As String


'-- Properties --
Public RA_AVG_MW As String
Public RA_AVG_DEN As String
Public RA_AVG_CP As String
Public RA_AVG_H As String
Public RA_VAP_DEN As String
Public RA_LIQ_DEN As String
Public RA_SOL_DEN As String
Public RA_VAP_CP As String
Public RA_LIQ_CP As String
Public RA_SOL_CP As String
Public RA_VAP_VIS As String
Public RA_LIQ_VIS As String
Public RA_BUB_POINT As String

Public RA_exp_CP As String
Public RA_AVG_K As String
Public RA_VAP_K As String
Public RA_LIQ_K As String
Public RA_SOL_K As String
Public RA_real_VIS As String
Public RA_real_DES As String
Public RA_exp_pH As String

'-- Balance Parameters --
Public RA_VAPOR_PER As String
Public RA_LIQUID_PER As String
Public RA_SOLID_PER As String
Public RA_MCPTT As String
Public RA_TS_PER As String
Public RA_MOISTURE_PER As String
Public RA_ETHA_PER As String
Public RA_ASH_PER As String
Public RA_GLUCAN_PER As String
Public RA_XYLAN_PER As String
Public RA_GLUCOSE_PER As String
Public RA_XYLOSE_PER As String
Public RA_CO2_PER As String
Public RA_pH As String
Public RA_DBO As String
Public RA_DQO As String
Public RA_CONC_H2SO4 As String
Public RA_H2O_H2SO4 As String

Public RA_TIS As String
Public RA_TSS As String
Public RA_TOC As String
Public RA_SiO2 As String
Public RA_Fe2mg As String
Public RA_Cu2mg As String
Public RA_HARDNESS As String
Public RA_O2mg As String
Public RA_Clmg As String
Public RA_CONDUCTIVITY As String
Public RA_OILmg As String
Public RA_FIBERS As String
Public RA_FILLERS As String
Public RA_PLASTICS As String
Public RA_INERTS As String

Public RA_CL As String
Public RA_S As String
Public RA_UREA As String
Public RA_CaCl2 As String

'-- Stream list parameter --
Public RA_STREAM_NUMBER As String
Public RA_SERVICE As String
Public RA_PUMP_CRITERIA As String
Public RA_FROM As String
Public RA_TO As String

Public lon_TOTAL_PAR As Long
'---------------------------------------------------------------------------------------
' Procedure : Update_Components_Range
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Update variables Range
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub Update_Components_Range()

lon_TOTAL_PAR = 261

'-- Design variables --
 RA_ACF_MASS = "D2"
 RA_ACF_LIQ = "D3"
 RA_ACF_VAP = "D4"
 RA_DESIGN_FACTOR = "D5"
 RA_DF_MASS = "D6"
 RA_DF_LIQ = "D7"
 RA_DF_VAP = "D8"
 RA_TEMP = "D9"
 RA_PRES = "D10"
 RA_VAP_FRAC = "D11"

'-- Liquid & Soluble Solid --
 RA_WATER = "D12"
 RA_ETHANOL = "D13"
 RA_GLUCOSE = "D14"
 RA_GLU_OLIG = "D15"
 RA_XYLOSE = "D16"
 RA_XYL_OLIG = "D17"
 RA_ARA_OLIG = "D18"
 RA_GALACTOSE = "D19"
 RA_MANNOSE = "D20"
 RA_BIO_EXTR = "D21"
 RA_OTH_INORG = "D22"
 RA_NUTRIENTS = "D23"
 RA_OTHER_SS = "D24"
 RA_ACID_SUL = "D25"
 RA_CAUSTIC = "D26"
 RA_AMM_HYDR = "D27"
 RA_ACETIC = "D28"
 RA_FURFURAL = "D29"
 RA_HMF = "D30"
 RA_OTH_MET_PROD = "D31"
 RA_DENATUR = "D32"
 RA_CORR_INH = "D33"
 RA_FLOCULANT = "D34"
 RA_CACO3 = "D35"
 RA_GYPSUM = "D36"
 RA_COAGULANT = "D37"
 RA_NON_OXI = "D38"
 RA_BISULF = "D39"
 RA_HCL = "D40"
 RA_OIL = "D41"
 RA_ANTIFOAM = "D42"
 RA_H3PO4 = "D43"
 RA_MEDIA = "D44"
 
'-- Gases --
 RA_OXYGEN = "D45"
 RA_OZONE = "D46"
 RA_NITROGEN = "D47"
 RA_METHANE = "D48"
 RA_CARB_DIO = "D49"
 RA_HYDR_SULF = "D50"
 RA_NITRO_OXI = "D51"
 RA_SULFUR_DIO = "D52"
 RA_OTH_GAS = "D53"
 RA_OXI = "D54"
 RA_CO = "D55"
 RA_NACLO = "D56"
 RA_HNO3 = "D57"
 RA_ETHANE = "D58"
 
'-- Salts --
 RA_PROPANE = "D59"
 RA_BUTANE = "D60"
 RA_SULF_SALT = "D61"
 RA_AMM_SALT = "D62"
 RA_CACO3_SALT = "D63"
 RA_GYPSUM_SALT = "D64"
 RA_SALT1 = "D65"
 RA_SALT2 = "D66"
 RA_SALT3 = "D67"
 RA_SALT4 = "D68"
 RA_SALT5 = "D69"
 RA_SALT6 = "D70"
 RA_SALT7 = "D71"
 RA_SALT8 = "D72"
 RA_SALT9 = "D73"
 RA_OH = "D74"
 RA_H3O = "D75"
 RA_NH4 = "D76"
 RA_HSO4 = "D77"
 RA_SO4 = "D78"
 RA_NA = "D79"
 RA_NH3 = "D80"
 RA_CA = "D81"
 RA_HCO3 = "D82"
 RA_CO3 = "D83"
 RA_NH2COO = "D84"
 RA_CAOH = "D85"
 RA_Fe2 = "D86"
 RA_Cu2 = "D87"
 RA_k = "D88"
 
'-- Insoluble Solid --
 RA_CELLULOSE = "D89"
 RA_XYLAN = "D90"
 RA_GALACTAN = "D91"
 RA_MANNAN = "D92"
 RA_ARABINAN = "D93"
 RA_LIGNIN = "D94"
 RA_PROTEIN = "D95"
 RA_ACETATE = "D96"
 RA_ASH = "D97"
 RA_ENZYMES = "D98"
 RA_ORGANISMS = "D99"
 RA_OTH_INS_SOL = "D100"
 RA_ACETYL = "D101"
 RA_CELLOBIOSE = "D102"
 RA_LACTOSIDE = "D103"
 RA_PENICILIN = "D104"

'-- Waste components --
 RA_ORG_FRAC = "D111"
 RA_TEXT = "D112"
 RA_PAP_CARD = "D113"
 RA_PET = "D114"
 RA_HDPE = "D115"
 RA_PP = "D116"
 RA_FILM = "D117"
 RA_OTH_PLAS = "D118"
 RA_NON_FE = "D119"
 RA_FE = "D120"
 RA_INERT = "D121"
 RA_GLASS = "D122"
 RA_BRISCKS = "D123"
 RA_HAZAR_WAS = "D124"
 RA_ELECT_WAS = "D125"
 RA_WOOD = "D126"
 RA_GUMS = "D127"
 RA_BULK_WAS = "D128"
 '-- Straw components --
 RA_STRAW = "D129"
 RA_STRING = "D130"
 RA_FORGEIN = "D131"
 RA_DUST = "D132"

'-- Properties --
 RA_AVG_MW = "D133"
 RA_AVG_DEN = "D134"
 RA_AVG_CP = "D135"
 RA_AVG_H = "D136"
 RA_VAP_DEN = "D137"
 RA_LIQ_DEN = "D138"
 RA_SOL_DEN = "D139"
 RA_VAP_CP = "D140"
 RA_LIQ_CP = "D141"
 RA_SOL_CP = "D142"
 RA_VAP_VIS = "D143"
 RA_LIQ_VIS = "D144"
 RA_BUB_POINT = "D145"

 RA_exp_CP = "D148"
 RA_AVG_K = "D149"
 RA_VAP_K = "D150"
 RA_LIQ_K = "D151"
 RA_SOL_K = "D152"
 RA_real_VIS = "D153"
 RA_real_DES = "D154"
 RA_exp_pH = "D150"
 
'-- Parameters --
 RA_VAPOR_PER = "D156"
 RA_LIQUID_PER = "D157"
 RA_SOLID_PER = "D158"
 RA_MCPTT = "D159"
 RA_TS_PER = "D160"
 RA_MOISTURE_PER = "D161"
 RA_ETHA_PER = "D162"
 RA_ASH_PER = "D163"
 RA_GLUCAN_PER = "D164"
 RA_XYLAN_PER = "D165"
 RA_GLUCOSE_PER = "D166"
 RA_XYLOSE_PER = "D167"
 RA_CO2_PER = "D168"
 RA_pH = "D169"
 RA_DBO = "D170"
 RA_DQO = "D171"

 RA_CONC_H2SO4 = "D172"
 RA_H2O_H2SO4 = "D173"

 RA_TIS = "D178"
 RA_TSS = "D179"
 RA_TOC = "D180"
 RA_SiO2 = "D181"
 RA_Fe2mg = "D182"
 RA_Cu2mg = "D183"
 RA_HARDNESS = "D184"
 RA_O2mg = "D185"
 RA_Clmg = "D186"
 RA_CONDUCTIVITY = "D187"
 RA_OILmg = "D188"
 RA_FIBERS = "D189"
 RA_FILLERS = "D190"
 RA_PLASTICS = "D191"
 RA_INERTS = "D192"

RA_CL = "D193"
RA_S = "D194"
RA_UREA = "D195"
RA_CaCl2 = "D196"

'-- Stream list parameter --
RA_STREAM_NUMBER = "D197"
RA_SERVICE = "D198"
RA_PUMP_CRITERIA = "D199"
RA_FROM = "D200"
RA_TO = "D201"

End Sub
