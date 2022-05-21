Attribute VB_Name = "ModLog"
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
' Module    : ModLog
' DateTime  : 08/14/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for work with Log
' References: N/A
' Functions : N/A
' Procedures:
'               1-AddLog
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Public LogContainer() As String
'-- If you are using this mod, please, remember that you must initializate these var
' into a global Mod whit the next sentence=> LogCounter=0 & int_CurrentLevel=0,1,2 or 3
' depending on the tool status --
' NOTE: int_CurrentLevel
'                       0-> Each function Step
'                       1-> Control functions Steps
'                       2-> Calling main functions
'                       3-> Errors

Public LogCounter As Variant
Public int_CurrentLevel As Variant
'---------------------------------------------------------------------------------------
' Procedure : AddLog
' DateTime  : 08/14/2013
' Author    : José García Herruzo
' Purpose   : Add a new log msg to the log Matrix
' Arguments :
'             str_myMsg                 --> String which contents the error msg
'             int_myLevel               --> Level of the msg
'                                               *0-> Each function Step
'                                               *1-> Calling functions&Procedure
'                                               *2-> Calling from form
'                                               *3-> Errors
'---------------------------------------------------------------------------------------
Public Sub AddLog(ByVal str_myMsg As String, ByVal int_myLevel As Integer)

'-- Redim the container --
If int_myLevel >= int_CurrentLevel Then
    If LogCounter = 0 Then
    
        ReDim LogContainer(1)
        LogContainer(0) = Now & Chr(9) & str_myMsg
    
    Else
    
        ReDim Preserve LogContainer(LogCounter)
        LogContainer(LogCounter) = Now & Chr(9) & str_myMsg
        
    End If

End If

LogCounter = LogCounter + 1

End Sub

