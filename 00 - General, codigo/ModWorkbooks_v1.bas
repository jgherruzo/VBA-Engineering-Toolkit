Attribute VB_Name = "ModWorkbooks_v1"
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
' Module    : ModWorkbooks_v1
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures to work with workbooks
' References: N/A
' Functions :
'               1-Close_WB_NoSave
'               2-FileExist
'               3-IsFileOpen
' Procedures:
'               1-CreateWorkbook
'               2-xp_Export_CSV
'               3-SeparadorPunto
'               4-SeparadorComa
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/04/2013  JGH     IsFileOpen is modified in order to launch out all errors
'       08/16/2013  JGH     Log is added
'       08/20/2013  JGH     IsFileOpened is modified to use with users
'       09/10/2013  JGH     CreateWorkbook is added
'       09/17/2013  JGH     FileExist is added
'       10/24/2013  JGH     Log is eliminated
'       29/01/2014  JGH     FileExist is modified
'       24/06/2020  JGH     Procedures 2 to 4 are added
'----------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Procedure : Close_WB_NoSave
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Close workbooks without saving it
' Arguments :
'             str_Complete_Path         --> Path+name
'---------------------------------------------------------------------------------------
Public Sub Close_WB_NoSave(ByVal str_Complete_Path As String)

On Error GoTo myhandler

Application.DisplayAlerts = False
Windows(str_Complete_Path).Activate
Application.CutCopyMode = False
ActiveWindow.Close (False)
Application.DisplayAlerts = True

Exit Sub
myhandler:
            
End Sub

'---------------------------------------------------------------------------------------
' Function  : IsFileOpen
' DateTime  : 04/24/2013
' Author    : Obtained from: http://support.microsoft.com/kb/291295/es; Modified by
'               José García Herruzo
' Purpose   : This function checks to see if a file is open or not. If the file is
'             already open, it returns True. If the file is not open, it returns
'             False. Otherwise, a run-time error occurs because there is
'             some other problem accessing the file.
' Arguments :
'             filename                   --> Path+name
'---------------------------------------------------------------------------------------
Public Function IsFileOpen(filename As String) As Boolean
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         IsFileOpen = False
        Case 70
         IsFileOpen = True
        Case Else
            Error errnum
    End Select

End Function

'---------------------------------------------------------------------------------------
' Function  : CreateWorkbook
' DateTime  : 07/10/2013
' Author    : José García Herruzo
' Purpose   : Generate a new workbook
' Arguments :
'             str_Path                   --> Path
'             str_WorkbookName           --> name
'---------------------------------------------------------------------------------------
Public Function CreateWorkbook(ByVal str_Path As String, str_WorkbookName) As Workbook
    
On Error GoTo myhandler

    Dim wb As Workbook
    
    Set wb = Workbooks.Add
    wb.SaveAs (str_Path & "\" & str_WorkbookName & ".xls")
    CreateWorkbook = wb
    
    Set wb = Nothing
    
Exit Function
myhandler:
Set wb = Nothing
            
End Function

'---------------------------------------------------------------------------------------
' Function  : FileExist
' DateTime  : 09/17/2013
' Author    : José García Herruzo
' Purpose   : Return true if file exist
' Arguments :
'             filename                   --> Path+name
'---------------------------------------------------------------------------------------
Public Function FileExist(filename As String) As Boolean
    Dim filenum As Integer, errnum As Integer
    Dim bol As Boolean
    
    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 53
         FileExist = False
        Case 76
         FileExist = False
        Case Else
        FileExist = True
    End Select

End Function

'---------------------------------------------------------------------------------------
' Function  : xp_Export_CSV
' DateTime  : 24/06/2013
' Author    : José García Herruzo
' Purpose   : Export CSV
' Arguments :
'             myWs                   --> Worksheet to be exported
'             str_File               --> Path+File where it must be exported
'---------------------------------------------------------------------------------------
Public Sub xp_Export_CSV(ByVal myWs As Worksheet, ByVal str_File As String)
    
    Dim CurrentWB As Workbook, TempWB As Workbook
    
    Set CurrentWB = ActiveWorkbook
    myWs.Activate
    ActiveWorkbook.ActiveSheet.UsedRange.Copy

    Set TempWB = Application.Workbooks.Add(1)
    With TempWB.Sheets(1).Range("A1")
      .PasteSpecial xlPasteValues
      .PasteSpecial xlPasteFormats
    End With

    Application.DisplayAlerts = False
    TempWB.SaveAs filename:=str_File, FileFormat:=xlCSV, CreateBackup:=False, Local:=True
    TempWB.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : SeparadorPunto
' DateTime  : 24/06/2013
' Author    : José García Herruzo
' Purpose   : Use . for decimal separator
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub SeparadorPunto()
Dim Tex As Variant, Car As Variant, Lar As Integer
Application.ScreenUpdating = False
On Error Resume Next
With Application
    .DecimalSeparator = "."
    .ThousandsSeparator = ","
    .UseSystemSeparators = False
End With
Application.ScreenUpdating = True
End Sub
'---------------------------------------------------------------------------------------
' Function  : SeparadorComa
' DateTime  : 24/06/2013
' Author    : José García Herruzo
' Purpose   : Use , for decimal separator
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub SeparadorComa()
Dim Tex As Variant, Car As Variant, Lar As Integer
Application.ScreenUpdating = False
On Error Resume Next
With Application
    .DecimalSeparator = ","
    .ThousandsSeparator = "."
    .UseSystemSeparators = False
End With
Application.ScreenUpdating = True
End Sub

