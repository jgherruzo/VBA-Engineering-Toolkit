Attribute VB_Name = "ModRound"
'           .---.        .-----------
'          /     \  __  /    ------
'         / /     \(..)/    -----
'        //////   ' \/ `   ---
'       //// / // :    : ---
'      // /   /  /`    '--
'     // /        //..\\
'   o===|========UU====UU=====-  -==========================o
'                '//||\\`
'                       DEVELPOPED BY JGH
'
'   -=====================|===o  o===|======================-
Option Explicit
'----------------------------------------------------------------------------------------
' Module    : ModRound
' DateTime  : 03/15/2013
' Author    : José García Herruzo; Copied from http://support.microsoft.com/kb/196652/es
' Purpose   : This module contents function to round number
' References: N/A
' Functions :
'               1-AsymDown
'               2-SymDown
'               3-AsymUp
'               4-symUp
'               5-AsymArith
'               6-SymArith
'               7-BRound
'               8-RandRound
'               9-AltRound
'               10-ATruncDigits
'               11-AsymArithDec
'               12-Round2CB
' Procedures: N/A
' Status    : CLOSE
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A         N/A     N/A
'----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Function  : AsymDown
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Asymmetrically rounds numbers down - similar to Int().
'                 Negative numbers get more negative.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function AsymDown(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

AsymDown = Int(X * Factor) / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : SymDown
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Symmetrically rounds numbers down - similar to Fix().
'                 Truncates all numbers toward 0.
'                 Same as AsymDown for positive numbers.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function SymDown(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

SymDown = Fix(X * Factor) / Factor

'Alternately:
'SymDown = AsymDown(Abs(X), Factor) * Sgn(X)

End Function
'---------------------------------------------------------------------------------------
' Function  : AsymUp
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Asymmetrically rounds numbers fractions up.
'                 Same as SymDown for negative numbers.
'                 Similar to Ceiling
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function AsymUp(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

Dim Temp As Double
Temp = Int(X * Factor)
AsymUp = (Temp + IIf(X = Temp, 0, 1)) / Factor

End Function

'---------------------------------------------------------------------------------------
' Function  : SymUp
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Symmetrically rounds fractions up - that is, away from 0.
'                 Same as AsymUp for positive numbers.
'                 Same as AsymDown for negative numbers
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function SymUp(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

Dim Temp As Double

Temp = Fix(X * Factor)
SymUp = (Temp + IIf(X = Temp, 0, Sgn(X))) / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : AsymArith
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Asymmetric arithmetic rounding - rounds .5 up always.
'                 Similar to Java worksheet Round function.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function AsymArith(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

AsymArith = Int(X * Factor + 0.5) / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : SymArith
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Symmetric arithmetic rounding - rounds .5 away from 0.
'                 Same as AsymArith for positive numbers.
'                 Similar to Excel Worksheet Round function
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function SymArith(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

SymArith = Fix(X * Factor + 0.5 * Sgn(X)) / Factor

'Alternately:
'SymArith = Abs(AsymArith(X, Factor)) * Sgn(X)

End Function
'---------------------------------------------------------------------------------------
' Function  : BRound
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Banker's rounding.
'                 Rounds .5 up or down to achieve an even number.
'                 Symmetrical by definition.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function BRound(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

'For smaller numbers:
'BRound = CLng(X * Factor) / Factor

Dim Temp As Double
Dim FixTemp As Double

Temp = X * Factor
FixTemp = Fix(Temp + 0.5 * Sgn(X))

'Handle rounding of .5 in a special manner
If Temp - Int(Temp) = 0.5 Then

    If FixTemp / 2 <> Int(FixTemp / 2) Then ' Is Temp odd
        
        'Reduce Magnitude by 1 to make even
        FixTemp = FixTemp - Sgn(X)
        
    End If
    
End If

BRound = FixTemp / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : RandRound
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Random rounding.
'                 Rounds .5 up or down in a random fashion.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function RandRound(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

'-- Should Execute Randomize statement somewhere prior to calling --
Dim Temp As Double
Dim FixTemp As Double

Temp = X * Factor
FixTemp = Fix(Temp + 0.5 * Sgn(X))

'-- Handle rounding of .5 in a special manner --

If Temp - Int(Temp) = 0.5 Then

    '-- Reduce Magnitude by 1 in half the cases --
    FixTemp = FixTemp - Int(Rnd * 2) * Sgn(X)
    
End If

RandRound = FixTemp / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : AltRound
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : Alternating rounding.
'                 Alternates between rounding .5 up or down.
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function AltRound(ByVal X As Double, Optional ByVal Factor As Double = 1) As Double

Static fReduce As Boolean
Dim Temp As Double
Dim FixTemp As Double

Temp = X * Factor
FixTemp = Fix(Temp + 0.5 * Sgn(X))

'-- Handle rounding of .5 in a special manner --
If Temp - Int(Temp) = 0.5 Then
    '-- Alternate between rounding .5 down (negative) and up (positive) --
    If (fReduce And Sgn(X) = 1) Or (Not fReduce And Sgn(X) = -1) Then
    
        ' -- Or, replace the previous If statement with the following to
        '    alternate between rounding .5 to reduce magnitude and increase
        '   magnitude--
        
        If fReduce Then
        
            FixTemp = FixTemp - Sgn(X)
            
        End If
        
        fReduce = Not fReduce
        
    End If

End If

AltRound = FixTemp / Factor

End Function
'---------------------------------------------------------------------------------------
' Function  : ADownDigits
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   :  Same as AsyncTrunc but takes different arguments
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Public Function ADownDigits(ByVal X As Double, Optional ByVal Digits As Integer = 0) As Double

ADownDigits = AsymDown(X, 10 ^ Digits)

End Function
'---------------------------------------------------------------------------------------
' Function  : AsymArithDec
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : de asimétrico aritméticos redondeo utilizando el tipo de datos decimal
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Function AsymArithDec(ByVal X As Variant, Optional ByVal Factor As Variant = 1) As Variant

If Not IsNumeric(X) Then
    
    AsymArithDec = X
    
Else

    If Not IsNumeric(Factor) Then
    
        Factor = 1
        AsymArithDec = Int(CDec(X * Factor) + 0.5)
        
End If

End Function
'---------------------------------------------------------------------------------------
' Function  : ADownDigits
' DateTime  : 22/04/2014
' Author    : Microsoft
' Purpose   : variación rígida que se realiza redondeo bancario 2 dígitos decimales
' Arguments :
'             x                        --> Value to ROUND
'---------------------------------------------------------------------------------------
Function Round2CB(ByVal X As Currency) As Currency
Round2CB = CCur(X / 100) * 100
End Function

