Attribute VB_Name = "ModErrorHandler"
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
' Module    : ModErrorHandler
' DateTime  : 08/14/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for work with Error
' References: N/A
' Functions : N/A
' Procedures:
'               1-AddError
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Public ErrorContainer() As String
'-- If you are using this mod, please, remember that you must initializate this var
' into a global Mod whit the next sentence=> ErrorCounter=0 --
Public ErrorCounter As Variant
'---------------------------------------------------------------------------------------
' Procedure : AddError
' DateTime  : 08/14/2013
' Author    : José García Herruzo
' Purpose   : Add a new error to the error Matrix
' Arguments :
'             str_myMsg                 --> String which contents the error msg
'---------------------------------------------------------------------------------------
Public Sub AddError(ByVal str_myMsg As String)

'-- Redim the container --
If ErrorCounter = 0 Then

    ReDim ErrorContainer(1)
    ErrorContainer(0) = str_myMsg

Else

    ReDim ErrorContainer(ErrorCounter)
    ErrorContainer(ErrorCounter) = str_myMsg
    
End If

ErrorCounter = ErrorCounter + 1

End Sub
