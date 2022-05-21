Attribute VB_Name = "ModFunction_v1"
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
' Module    : ModFunction_v1
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : This module contents function not classify
' References: N/A
' Functions :
'               1-CheckingLimit
'               2-YesNoQuestion
'               3-GetSmallLetterLessFirst
'               4-AddDigits
'               5-xlReorderMatrix
'               6-xlIsEven
'               7-xlGetInitials
'               8-txtNoAcc
' Procedures:
'               1-AddDigits
'               2-PDFCreatorPrint
'               3-xlWaitAMoment
'               4-PrintWordPDFCreator
'               5-xsGetDecimalFormat
' Status    : OPEN
' Updates   :
'       DATE        USER    DESCRIPTION
'       08/22/2013  JGH     GetSmallLetterLessFirst function is developed
'       08/26/2013  JGH     AddDigits procedure is developed
'       09/25/2013  JGH     PDFCreatorPrint & xlWaitAMoment are added
'       10/02/2013  JGH     AsymUp & RoundDown are added
'       17/01/2014  JGH     PDFCreatorPrint is modified to not call xlWaitAMoment
'       17/01/2014  JGH     PrintWordPDFCreator are added
'       17/01/2014  JGH     EHS is added
'       10/02/2014  JGH     PDFCreatorPrint is modified to add exit sub before to error
'                           handler
'       15/04/2014  JGH     xlReorderMatrix is added
'       22/04/2014  JGH     AsymUp & RoundDown are removed. See ModRound for more
'                           details
'       20/05/2014  JGH     xsGetDecimalFormat is added from ModBECS_DB
'       24/09/2014  JGH     xlIsEven is added
'       29/12/2014  JGH     xlGetInitials and txtNoAcc are added
'       30/12/2014  JGH     GetSmallLetterLessFirst is modified to avoid error with empty
'                           strings
'       10/03/2015  JGH     xlReorderMatrix is modified to order from first row
'----------------------------------------------------------------------------------------
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    Public Error_Function_v1 As Integer
'--------------------------- ERROR INFORMATION BOX --------------------------------------
'   KEY         FUNTION or PROCEDURE                                                    '
'   1           PDFCreatorPrint                                                         '
'   2           PrintWordPDFCreator                                                     '
'   3           xlReorderMatrix                                                         '
'------------------------- END ERROR INFORMATION BOX ------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'---------------------------------------------------------------------------------------
' Function  : CheckingLimit
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Return true if the value is between the limit
' Arguments :
'             myValue                  --> Value to chekc
'             myLimit                  --> +- Batery limit
'---------------------------------------------------------------------------------------
Public Function CheckingLimit(ByVal myValue As Variant, ByVal myLimit As Variant) As Boolean

If myValue < myLimit And myValue > -myLimit Then

    CheckingLimit = True
    
Else

    CheckingLimit = False

End If

End Function

'---------------------------------------------------------------------------------------
' Function  : YesNoQuestion
' DateTime  : 07/10/2013
' Author    : José García Herruzo
' Purpose   : Return true if you click in yes
' Arguments :
'             str_Question             --> Question string
'             str_Title                --> Box tittle
'---------------------------------------------------------------------------------------
Public Function YesNoQuestion(ByVal str_Question As String, Optional str_Title As String) As Boolean

Dim Answer As String

    Answer = MsgBox(str_Question, vbQuestion + vbYesNo, str_Title)

    If Answer = vbNo Then
        
        YesNoQuestion = False
    
    Else
    
        YesNoQuestion = True
        
    End If

End Function

'---------------------------------------------------------------------------------------
' Function  : GetSmallLetterLessFirst
' DateTime  : 08/22/2013
' Author    : José García Herruzo
' Purpose   : Convert a sentence-> First letter Capital letter and the others small
'             letters
' Arguments :
'             ra_myRange               --> Source of the string
'             str_Title                --> Box tittle
'---------------------------------------------------------------------------------------
Public Function GetSmallLetterLessFirst(ByVal str_myString As String) As String

Dim str_CapitalLetter As String
Dim str_SmallLetter As String
Dim int_Lenth As String

If str_myString <> "" Then

    '-- Exract string lenth --
    int_Lenth = Len(str_myString)
    '-- Extract first letter --
    str_CapitalLetter = Left(str_myString, 1)
    '-- Extract the others --
    str_SmallLetter = Right(str_myString, int_Lenth - 1)
    '-- Convert the string --
    str_myString = UCase(str_CapitalLetter) & LCase(str_SmallLetter)
    '-- Return the converted string --
    GetSmallLetterLessFirst = str_myString
    
Else

    GetSmallLetterLessFirst = str_myString

End If

End Function
'---------------------------------------------------------------------------------------
' Function  : AddDigits
' DateTime  : 08/26/2013
' Author    : José García Herruzo
' Purpose   : Add 0 and convert to text format
' Arguments :
'             myColumn                 --> the letter of the column
'             myFirstRow               --> First row
'             str_worksheet            --> Worksheet name
'             str_Workbook             --> Worksheet name
'---------------------------------------------------------------------------------------
Public Sub AddDigits(ByVal myColumn As String, ByVal myFirstRow As String, ByVal str_worksheet As String, ByVal str_Workbook As String)

Dim myLastRow As Integer
Dim i As Integer
Dim var_help As Variant
Dim str_help As String
Dim ra As Range

myLastRow = Workbooks(str_Workbook).Worksheets(str_worksheet).Range(myColumn & "16000").End(xlUp).Row
myLastRow = myLastRow - myFirstRow

Set ra = Workbooks(str_Workbook).Worksheets(str_worksheet).Range(myColumn & myFirstRow)

For i = 0 To myLastRow
    var_help = ra.Offset(i, 0).Value
    
    If var_help < 10000 And var_help > 10 Then
    
        str_help = "0" & var_help
    
    ElseIf var_help < 10 Then
    
        str_help = "00" & var_help
    
    Else
    
        str_help = var_help
    
    End If
    
    With ra.Offset(i, 0)
        .NumberFormat = "@"
        .Value = str_help
    End With

Next i

End Sub
'---------------------------------------------------------------------------------------
' Function  : PDFCreatorPrint
' DateTime  : 09/25/2013
' Author    : José García Herruzo
' Purpose   : Print Excel to PDF
' Arguments :
'             myWS                     --> worksheet to be printed
'             myPath                   --> Path where pdf must be saved
'             str_FileName             --> pdf name
'---------------------------------------------------------------------------------------
Public Sub PDFCreatorPrint(ByVal myws As Worksheet, ByVal myPath As String, ByVal str_FileName As String)
     
Dim pdfjob As Object
Dim bol_IsStarted As Boolean

On Error GoTo myhandler

'-- set bands to false --
bol_IsStarted = False

Set pdfjob = CreateObject("PDFCreator.clsPDFCreator")
    
With pdfjob
    
    If .cStart("/NoProcessingAtStartup") = False Then
            
        MsgBox "Can't initialize PDFCreator.", vbCritical + vbOKOnly, "PrtPDFCreator"
        Exit Sub
    
    End If
    bol_IsStarted = True
    .cOption("UseAutosave") = 1
    .cOption("UseAutosaveDirectory") = 1
    .cOption("AutosaveDirectory") = myPath
    .cOption("AutosaveFilename") = str_FileName
    .cOption("AutosaveFormat") = 0 ' 0 = PDF
    .cClearCache
    
End With
     
'Print the document to PDF
myws.PrintOut ActivePrinter:="PDFCreator"
     
'Wait until the print job has entered the print queue
Do Until pdfjob.cCountOfPrintjobs = 1
        
    DoEvents

Loop
pdfjob.cPrinterStop = False
     
'Wait until PDF creator is finished then release the objects
Do Until pdfjob.cCountOfPrintjobs = 0
    
    DoEvents

Loop
pdfjob.cClose
Set pdfjob = Nothing

Exit Sub

myhandler:
    If bol_IsStarted = True Then
    
        pdfjob.cClose
        
    End If
    Set pdfjob = Nothing
    Error_Function_v1 = 1

End Sub
'---------------------------------------------------------------------------------------
' Function  : xlWaitAMoment
' DateTime  : 09/25/2013
' Author    : José García Herruzo
' Purpose   : Stop the code during especificated time
' Arguments :
'             lon_Time                 --> time to be stopped in second
'---------------------------------------------------------------------------------------
Public Sub xlWaitAMoment(ByVal lon_Time As Long)

Dim a As Date
Dim b As Date

b = Now
a = b
Do While DateDiff("s", b, a) = lon_Time

    a = Now

Loop

End Sub
'---------------------------------------------------------------------------------------
' Function  : PrintWordPDFCreator
' DateTime  : 17/01/2014
' Author    : José García Herruzo
' Purpose   : Print Excel to PDF
' Arguments :
'             str_SourceDoc            --> Source document name
'             str_SourcePath           --> Source path
'             str_DestinyDoc           --> Final document name
'             str_destinyPath          --> Destiny directory
'---------------------------------------------------------------------------------------
Public Sub PrintWordPDFCreator(ByVal str_SourceDoc As String, ByVal str_SourcePath As String, _
                                ByVal str_DestinyDoc As String, ByVal str_DestinyPath As String, _
                                ByVal str_DocExtension As String)
     
Dim pdfjob As Object
Dim app_Word As Object
Dim fil_Word As Variant
Dim bol_IsStarted As Boolean
Dim bol_IsOpenend As Boolean
Dim bol_IsActivated As Boolean

On Error GoTo myhandler

'-- set bands to false --
bol_IsStarted = False
bol_IsOpenend = False
bol_IsActivated = False

'-- Word object is set --
Set app_Word = CreateObject("Word.Application")
bol_IsActivated = True

'-- PDF creator is set as active printer --
app_Word.ActivePrinter = "PDFCreator"

'-- PDF creator object is set and setup --
Set pdfjob = CreateObject("PDFCreator.clsPDFCreator")
    
With pdfjob

    If .cStart("/NoProcessingAtStartup") = False Then
        
        MsgBox "Can't initialize PDFCreator.", vbCritical + _
        vbOKOnly, "PrtPDFCreator"
        Exit Sub
        
    End If
    bol_IsStarted = True
    .cOption("UseAutosave") = 1
    .cOption("UseAutosaveDirectory") = 1
    .cOption("AutosaveDirectory") = str_DestinyPath
    .cOption("AutosaveFilename") = str_DestinyDoc
    .cOption("AutosaveFormat") = 0 ' 0 = PDF
    .cClearCache
    
End With
    
'-- Word document is opened --
Set fil_Word = app_Word.Documents.Open(str_SourcePath & "\" & str_SourceDoc & "." & str_DocExtension)
bol_IsOpenend = True

'-- Print --
fil_Word.PrintOut

'Wait until the print job has entered the print queue
Do Until pdfjob.cCountOfPrintjobs = 1
    
    DoEvents

Loop

pdfjob.cPrinterStop = False
     
'Wait until PDF creator is finished then release the objects
Do Until pdfjob.cCountOfPrintjobs = 0
    
    DoEvents
    
Loop

pdfjob.cClose
    
'-- close the word --
fil_Word.Close

'-- close word application --
app_Word.Quit

Set app_Word = Nothing
Set fil_Word = Nothing
Set pdfjob = Nothing
    
Exit Sub

myhandler:
    If bol_IsStarted = True Then

        pdfjob.cClose
        
    End If
    If bol_IsOpenend = True Then
    
        fil_Word.Close
    
    End If
    If bol_IsActivated = True Then
    
        app_Word.Quit
    
    End If
    
    Set app_Word = Nothing
    Set fil_Word = Nothing
    Set pdfjob = Nothing
    Error_Function_v1 = 2
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : xlReorderMatrix
' DateTime  : 15/04/2014
' Author    : José García Herruzo
' Purpose   : Reorder Matrix depending on the key
' Arguments :
'             source                   --> Array to be ordered
'             int_Key                  --> Order
'                                           * 0 --> Lowest to highest
'                                           * 1 --> highest to Lowest
'             int_KeyColumn            --> Column which contents ordering parameter
'             int_ColumnBound          --> Column bound
'---------------------------------------------------------------------------------------
Public Function xlReorderMatrix(ByRef source() As Variant, ByVal int_Key As Integer, ByVal int_KeyColumn As Integer, _
                            ByVal int_ColumnBound As Integer) As Variant()

Dim Arr() As Variant
Dim help() As Variant
Dim TempArr() As Variant
Dim order() As Variant
Dim pointer() As Long

Dim i As Long
Dim j As Integer

Dim k As Long
Dim l As Integer

Dim int_Counter As Long
Dim lon_Counter As Long

On Error GoTo Error_Handler:

'-- creamos un vector con los parámetros a ordenar, el vector de localización y el de ayuda --
ReDim help(UBound(source), 1)
ReDim order(0)
ReDim pointer(0)
lon_Counter = 0
int_Counter = 0

For i = 0 To UBound(source)

    help(int_Counter, 0) = source(i, int_KeyColumn)   '--> valor a ordenar
    help(int_Counter, 1) = i                          '--> situación del valor en matriz origen
    int_Counter = int_Counter + 1
    
Next i

If int_Key = 0 Then '--> seleccionamos ordenar de menor a mayor
    
    order(0) = 1E+32
    
    Do Until lon_Counter = UBound(source)
    
        '-- primero buscamos el menor valor --
        For i = 0 To UBound(help)
        
            If help(i, 0) < order(lon_Counter) Then
                '-- anotamos el menor valor y su localizacion original --
                order(lon_Counter) = help(i, 0)
                pointer(lon_Counter) = help(i, 1)
                
            End If
        
        Next i
        
        '-- Trasvasamos la información a la matriz temporal --
        '-- Todo menos el ya extraido --
        ReDim TempArr(UBound(help) - 1, 1)
        
        int_Counter = 0
        For i = 0 To UBound(help)
        
            If help(i, 1) <> pointer(UBound(pointer)) Then
            
                TempArr(int_Counter, 0) = help(i, 0)
                TempArr(int_Counter, 1) = help(i, 1)
                int_Counter = int_Counter + 1
                
            End If
        
        Next i
        
        '-- la devolvemos a la original --
        ReDim help(UBound(TempArr), 1)
        
        For i = 0 To UBound(help)
            
            help(i, 0) = TempArr(i, 0)
            help(i, 1) = TempArr(i, 1)
                
        Next i
        
        '-- aumentamos el tamaño de las soluciones --
        lon_Counter = lon_Counter + 1
        ReDim Preserve order(lon_Counter)
        ReDim Preserve pointer(lon_Counter)
        order(lon_Counter) = 1E+32
        '-- volvemos a empezar
        
    Loop
    
    '-- escribimos el último valor
    order(lon_Counter) = help(0, 0)
    pointer(lon_Counter) = help(0, 1)
    
ElseIf int_Key = 1 Then '-- ordenamos de mayor a menor

    order(0) = -1E+32
    
    Do Until lon_Counter = UBound(source) - 3
    
        '-- primero buscamos el menor valor --
        For i = 0 To UBound(help)
        
            If help(i, 0) > order(lon_Counter) Then
                '-- anotamos el menor valor y su localizacion original --
                order(lon_Counter) = help(i, 0)
                pointer(lon_Counter) = help(i, 1)
                
            End If
        
        Next i
        
        '-- Trasvasamos la información a la matriz temporal--
        ReDim TempArr(UBound(help) - 1, 1)
        
        int_Counter = 0
        For i = 0 To UBound(help)
        
            If help(i, 1) <> pointer(UBound(pointer)) Then
            
                TempArr(int_Counter, 0) = help(i, 0)
                TempArr(int_Counter, 1) = help(i, 1)
                int_Counter = int_Counter + 1
                
            End If
        
        Next i
        
        '-- la devolvemos a la original --
        ReDim help(UBound(TempArr), 1)
        
        For i = 0 To UBound(help)
            
            help(i, 0) = TempArr(i, 0)
            help(i, 1) = TempArr(i, 1)
                
        Next i
        
        '-- aumentamos el tamaño de las soluciones --
        lon_Counter = lon_Counter + 1
        ReDim Preserve order(lon_Counter)
        ReDim Preserve pointer(lon_Counter)
        order(lon_Counter) = -1E+32
        '-- volvemos a empezar
        
    Loop
    
    '-- escribimos el último valor
    order(lon_Counter) = help(0, 0)
    pointer(lon_Counter) = help(0, 1)

End If

'-- creamos la matriz a devolver --
ReDim Arr(UBound(source), int_ColumnBound)

'-- introducimos a mano los parámetros no ordenados ?¿--

'For j = 0 To int_ColumnBound
'
'    Arr(0, j) = source(0, j)
'    Arr(1, j) = source(1, j)
'    Arr(UBound(source), j) = source(UBound(source), j)
'
'Next j

'-- y ahora se descargan los ordenados --
For i = 0 To UBound(pointer)

    For j = 0 To int_ColumnBound
    
        Arr(i, j) = source(pointer(i), j)
        
    Next j

Next i

xlReorderMatrix = Arr()

Exit Function

Error_Handler:
    'xlReorderMatrix(0) = 0
    Error_Function_v1 = 3

End Function

'---------------------------------------------------------------------------------------
' Procedure : xsGetDecimalFormat
' DateTime  : 10/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure to select a specified format depending on the value
' Arguments :
'             dou_value                 --> Contained value
'             ra_myRange                --> Container
'---------------------------------------------------------------------------------------
Public Sub xsGetDecimalFormat(ByVal dou_value As Double, ByRef ra_myRange As Range)

If dou_value >= 100 Or dou_value < -100 Then
    
    ra_myRange.NumberFormat = "##,##0"
    
ElseIf dou_value >= 10 And dou_value < 100 Then

    ra_myRange.NumberFormat = "##,##0.0"
    
ElseIf dou_value <= -10 And dou_value > -100 Then

    ra_myRange.NumberFormat = "##,##0.0"
    
ElseIf dou_value >= 1 And dou_value < 10 Then

    ra_myRange.NumberFormat = "##,##0.00"
    
ElseIf dou_value <= -1 And dou_value > -10 Then

    ra_myRange.NumberFormat = "##,##0.00"
    
ElseIf dou_value >= 0.1 And dou_value < 1 Then

    ra_myRange.NumberFormat = "##,##0.000"
    
ElseIf dou_value <= -0.1 And dou_value > -1 Then

    ra_myRange.NumberFormat = "##,##0.000"
    
Else

    ra_myRange.NumberFormat = "0.00"
    
End If

End Sub
'---------------------------------------------------------------------------------------
' Function  : xlIsEven
' DateTime  : 24/09/2014
' Author    : José García Herruzo
' Purpose   : Return true if given value is even
' Arguments :
'             dou_value                 --> value
'---------------------------------------------------------------------------------------
Public Function xlIsEven(ByVal dou_value As Double) As Boolean

Dim X As Variant

X = dou_value Mod 2

If X = 0 Then

    xlIsEven = True

Else

    xlIsEven = False

End If

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetInitials
' DateTime  : 29/12/2014
' Author    : José García Herruzo
' Purpose   : Return the initials of the given name
' Arguments :
'             str_FullName              --> Full name
'---------------------------------------------------------------------------------------
Public Function xlGetInitials(ByVal str_FullName As String) As String

Dim arr_Temp() As String
Dim i As Integer
Dim str_Initials As String

'-- check if full name is empty --

If str_FullName <> "" Then

    arr_Temp = Split(str_FullName, " ")
    
    For i = 0 To UBound(arr_Temp)
    
        str_Initials = str_Initials & Mid(arr_Temp(i), 1, 1)
    
    Next i
    
    xlGetInitials = txtNoAcc(str_Initials)

Else

    xlGetInitials = str_Initials
    
End If

End Function
'---------------------------------------------------------------------------------------
' Function  : txtNoAcc
' DateTime  : 29/12/2014
' Author    : From http://www.excel-avanzado.com/8889/eliminar-tildes-con-macros.html
'                modified by José García Herruzo
' Purpose   : Remove accent
' Arguments :
'             texto                     --> Name with accent
'---------------------------------------------------------------------------------------
Public Function txtNoAcc(ByVal texto As String) As String

Dim largoTexto As Long, iX As Long
Dim Lett As Long

txtNoAcc = ""
largoTexto = Len(texto)

For iX = 1 To largoTexto

    Lett = Asc(Mid(texto, iX, 1))
    Select Case Lett
    
        Case Is = 225
        txtNoAcc = txtNoAcc & Chr(97)
        Case Is = 233
        txtNoAcc = txtNoAcc & Chr(101)
        Case Is = 237
        txtNoAcc = txtNoAcc & Chr(105)
        Case Is = 243
        txtNoAcc = txtNoAcc & Chr(111)
        Case Is = 250
        txtNoAcc = txtNoAcc & Chr(117)
        
        Case Is = 193
        txtNoAcc = txtNoAcc & Chr(65)
        Case Is = 201
        txtNoAcc = txtNoAcc & Chr(69)
        Case Is = 205
        txtNoAcc = txtNoAcc & Chr(73)
        Case Is = 211
        txtNoAcc = txtNoAcc & Chr(79)
        Case Is = 218
        txtNoAcc = txtNoAcc & Chr(85)
        
        Case Else
        txtNoAcc = txtNoAcc & Mid(texto, iX, 1)
    
    End Select

Next iX

End Function

