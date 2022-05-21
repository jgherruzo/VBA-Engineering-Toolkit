Attribute VB_Name = "ModEnhacements"
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
' Module    : ModEnhacements
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : This module contents procedures and function in order to improve the sight
' References: N/A
' Functions :
'               1-xlGetStatus
' Procedures:
'               1-xlStartSettings
'               2-xlEndSettings
'               3-PaintGreen
'               4-PaintRed
'               5-PaintNoColor
'               6-PaintOrange
'               7-PaintGrey
'               8-PaintPastel
'               9-PaintBlue
'               10-PaintBlack
'               11-ShellAndWait
' Updates   :
'       DATE        USER    DESCRIPTION
'       06/06/2013  JGH     Paint procedures are added
'       07/24/2013  JGH     Pantones are updated
'       07/24/2013  JGH     Paint black procedure is added
'       20/06/2014  JGH     ShellAndWait is added
'       26/11/2014  JGH     xlGetStatus is added
'       24/10/2017  JGH     ShellAndWait is removed
'       06/03/2019  JGH     Procedure 1 and 2 are modified
'----------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Procedure : xlStartSettings
' DateTime  : 03/15/2013
' Author    : Aspen support/Antonio Rodriguez Donaire
' Purpose   : subroutine to change settings to put Excel "on hold" until a long process
'             is completed
' Arguments :
'             as_StatusString           --> Msg to show
'---------------------------------------------------------------------------------------
Public Sub xlStartSettings(Optional as_StatusString As String = "Working")

With Application
   .Calculation = xlCalculationManual
   .ScreenUpdating = False
   .EnableEvents = False
   .Cursor = xlWait
   .StatusBar = as_StatusString
End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlEndSettings
' DateTime  : 03/15/2013
' Author    : Aspen support/Antonio Rodriguez Donaire
' Purpose   : subroutine to change settings to default after a long process is completed
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub xlEndSettings()

With Application
   .Calculation = xlCalculationAutomatic
   .ScreenUpdating = True
   .EnableEvents = True
   .Cursor = xlDefault
   .StatusBar = False
   .WindowState = xlMaximized
End With

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintGreen
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 579
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintGreen(ByVal myRange As Range)

myRange.Interior.Color = RGB(199, 214, 163)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintRed
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 172
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintRed(ByVal myRange As Range)

myRange.Interior.Color = RGB(242, 79, 0)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintNoColor
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Unpaint a range
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintNoColor(ByVal myRange As Range)

myRange.Interior.Color = xlNone

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintOrange
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 156
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintOrange(ByVal myRange As Range)

myRange.Interior.Color = RGB(237, 194, 130)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintGrey
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 421
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintGrey(ByVal myRange As Range)

myRange.Interior.Color = RGB(191, 186, 181)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintPastel
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 608
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintPastel(ByVal myRange As Range)

myRange.Interior.Color = RGB(237, 232, 173)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintBlue
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 304
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintBlue(ByVal myRange As Range)

myRange.Interior.Color = RGB(173, 219, 227)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : PaintBlack
' DateTime  : 06/06/2013
' Author    : José García Herruzo
' Purpose   : Pantone 405
' Arguments :
'             myRange                  --> Range to paint
'---------------------------------------------------------------------------------------
Public Sub PaintBlack(ByVal myRange As Range)

myRange.Interior.Color = RGB(102, 89, 77)

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xlGetStatus
' DateTime  : 18/11/2014
' Author    : José García Herruzo
' Purpose   : This procedure built status bar
' Arguments :
'             dou_Value                 --> status
'---------------------------------------------------------------------------------------
Public Function xlGetStatus(ByVal int_Value As Integer) As String

Dim i As Integer
Dim str_Temp As String

For i = 0 To int_Value

    str_Temp = str_Temp & ">"

Next i

str_Temp = str_Temp & " -0- "

For i = int_Value To 100

    str_Temp = str_Temp & "<"

Next i

str_Temp = str_Temp & " " & int_Value & " %"

xlGetStatus = str_Temp

End Function
