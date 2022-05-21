Attribute VB_Name = "ModFileAndFolder"
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
' Module    : ModFileAndFolder
' DateTime  : 05/29/2013
' Author    : José García Herruzo
' Purpose   : This module contents functions and procedures applied to files and folder
' References: Microsoft scripting runtime
' Functions :
'               1-GetSubfolder
'               2-GetFiles
'               3-GetSelectedFileInfo
'               4-xlFolderExist
'               5-xlGetFolderSize
'               6-xlGetExcelFilePathName
'               7-xf_FileExist
' Procedures:
'               1-CreateFolder
'               2-CopyFolder
'               3-MoveFolder
'               4-CopiaArchivo
'               5-DeleteFolder
'               6-xsDeleteFile
'               7-xf_FileExist
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/23/2013  JGH     GetSelectedFileInfo function is added
'       08/14/2013  JGH     Log is added
'       15/01/2014  JGH     Log is removed
'       16/01/2014  JGH     Error handler system (EHS) is added
'       16/01/2014  JGH     CopyFolder, MoveFolder, MoveFile and CopyFile are added
'       04/06/2014  JGH     xlFolderExist is added
'       18/06/2014  JGH     DeleteFolder is added
'       18/06/2014  JGH     GetFiles is modified to accept argument with or without "\"
'       23/06/2014  JGH     DeleteFile is added
'       25/06/2014  JGH     SHFILEOPSTRUCT and SHFileOperation are added. They are
'                           required for the new xsCopyFile
'       25/06/2014  JGH     xsCopyFile is added
'       14/07/2014  JGH     xlFolderExist is modified
'       15/09/2014  JGH     xlGetFolderSize is added
'       06/10/2014  JGH     xlGetExcelFilePathName is added
'       25/06/2020  JGH     function 7 is added and procedure 4 is updated
'----------------------------------------------------------------------------------------
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As String
End Type
 
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const FO_COPY = &H2
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                    Public Error_FileAndFolder As Integer
'--------------------------- ERROR INFORMATION BOX --------------------------------------
'   KEY         FUNTION or PROCEDURE                                                    '
'   1           GetFiles                                                                '
'   2           CreateFolder                                                            '
'   3           GetSelectedFileInfo                                                     '
'   4           CopyFolder                                                              '
'   5           MoveFolder                                                              '
'   6           xsCopyFile                                                              '
'   7           GetSubfolderList                                                        '
'   8           DeleteFolder                                                            '
'   9           xsDeleteFile                                                            '
'   10          xlGetFolderSize                                                            '
'------------------------- END ERROR INFORMATION BOX ------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'---------------------------------------------------------------------------------------
' Function  : GetSubfolder
' DateTime  : 05/29/2013
' Author    : José García Herruzo
' Purpose   : Returns a string with each subfolder separated by \
' Arguments :
'             str_myPath               --> Main path
'---------------------------------------------------------------------------------------
Public Function GetSubfolder(ByVal str_mypath As String) As String

Dim fso As Variant
Dim Directorio As Variant
Dim Subdirectorio As Variant
Dim str_myString As Variant

On Error GoTo myhandler

If Right(str_mypath, 1) <> "\" Then
    str_mypath = str_mypath & "\"
End If

Set fso = CreateObject("Scripting.FileSystemObject")
str_myString = ""
   
Set Directorio = fso.GetFolder(str_mypath)

For Each Subdirectorio In Directorio.SubFolders

    str_myString = str_myString & "\" & Subdirectorio.Name

Next

GetSubfolder = str_myString

Exit Function

myhandler:
    Set fso = Nothing

    GetSubfolder = ""
    Error_FileAndFolder = 7

End Function
    
'---------------------------------------------------------------------------------------
' Function  : GetFiles
' DateTime  : 06/03/2013
' Author    : José García Herruzo
' Purpose   : Returns a string with each file in a folder separated by \
' Arguments :
'             str_myPath               --> Main path
'---------------------------------------------------------------------------------------
Public Function GetFiles(ByVal str_mypath As String) As String

Dim str_FileName As String
Dim str_myString As Variant

If Right(str_mypath, 1) <> "\" Then
    str_mypath = str_mypath & "\"
End If


str_FileName = Dir(str_mypath)

Do While str_FileName <> ""

    str_myString = str_myString & "\" & str_FileName
    str_FileName = Dir

Loop

GetFiles = str_myString
            
Exit Function

myhandler:
    GetFiles = ""
    Error_FileAndFolder = 1

End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateFolder
' DateTime  : 07/10/2013
' Author    : José García Herruzo
' Purpose   : Create a new folder
' Arguments :
'             str_Path                 --> Main path
'             str_FolderName           --> Folder Name
'---------------------------------------------------------------------------------------
Public Sub CreateFolder(ByVal str_Path As String, ByVal str_FolderName As String)

On Error GoTo myhandler

If Right(str_Path, 1) <> "\" Then
    
    str_Path = str_Path & "\"

End If

'-- check if path exist --
If Dir(str_Path, vbDirectory) <> "" Then

    '-- check if folder exist --
    If Dir(str_Path & str_FolderName, vbDirectory) = "" Then
    
        MkDir (str_Path & str_FolderName)
    
    End If

End If
            
Exit Sub

myhandler:
    
    Error_FileAndFolder = 2
    
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetSelectedFileInfo
' DateTime  : 07/23/2013
' Author    : José García Herruzo
' Purpose   : Returns selected file path+name
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function GetSelectedFileInfo() As String

Dim GetJustTheFileName As String
Dim fso As Variant
Dim strPath As String
Dim strName As String

On Error GoTo myhandler

Set fso = CreateObject("Scripting.FileSystemObject")

GetJustTheFileName = Application.GetOpenFilename

'-- Extract file name --
strName = fso.GetFileName(GetJustTheFileName)
'-- Extract file path --
strPath = CurDir

GetSelectedFileInfo = strPath & "\" & strName
            
Exit Function

myhandler:
    Set fso = Nothing
    
    GetSelectedFileInfo = ""
    Error_FileAndFolder = 3
    
End Function
'---------------------------------------------------------------------------------------
' Function  : CopyFolder
' DateTime  : 16/01/2014
' Author    : José García Herruzo
' Purpose   : Copy the content of the selected folder to a destiny folder
' Arguments :
'             str_From                 --> Origin Path
'             str_To                   --> Destiny path
'---------------------------------------------------------------------------------------
Public Sub CopyFolder(ByVal str_From, ByVal str_To)

Dim fso As Variant
Dim myFolder As Variant

On Error GoTo myhandler

'-- Set the new file system object --
Set fso = CreateObject("Scripting.FileSystemObject")

'-- Set desire folder --
Set myFolder = fso.GetFolder(str_From)

'-- Copy it --
fso.CopyFolder str_From, str_To
            
Set fso = Nothing
Set myFolder = Nothing

Exit Sub

myhandler:
    Set fso = Nothing
    Set myFolder = Nothing
    
    Error_FileAndFolder = 4
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : MoveFolder
' DateTime  : 16/01/2014
' Author    : José García Herruzo
' Purpose   : Copy the content of the selected folder to a destiny folder and remove
'             original loaction
' Arguments :
'             str_From                 --> Origin Path
'             str_To                   --> Destiny path
'---------------------------------------------------------------------------------------
Public Sub MoveFolder(ByVal str_From, ByVal str_To)

Dim fso As Variant
Dim myFolder As Variant

On Error GoTo myhandler

'-- Set the new file system object --
Set fso = CreateObject("Scripting.FileSystemObject")

'-- Set desire folder --
Set myFolder = fso.GetFolder(str_From)

'-- Copy it --
fso.CopyFolder str_From, str_To

'-- Delete the original folder --
myFolder.Delete True
            
Set fso = Nothing
Set myFolder = Nothing

Exit Sub

myhandler:
    Set fso = Nothing
    Set myFolder = Nothing
    
    Error_FileAndFolder = 5

End Sub
'---------------------------------------------------------------------------------------
' Function  : CopiaArchivo
' DateTime  : 25/06/2020
' Author    : Emilio Sancha
' Purpose   : Copy selected file to selected destiny folder
' Arguments :
'             strCarpetaOrigen         --> Source Path
'             strNombreArchivo         --> File name
'             strCarpetaDestino        --> Destiny Path
'---------------------------------------------------------------------------------------
Public Function CopiaArchivo(strOrigen As String, strDestino As String, Optional blnMover As Boolean, Optional blnSobreescribir As Boolean) As Boolean
Dim fso As Object, _
    Archivo As Object

On Error GoTo TratamientoErrores

Set fso = CreateObject("Scripting.FileSystemObject")

' si el archivo a copiar no existe lanzó un error
If Not fso.FileExists(strOrigen) Then Err.Raise vbObjectError + 513, "CopiaArchivo", "El archivo a copiar """ & strOrigen & """ no existe"
' si el archivo destino existe procedo según corresponda
If fso.FileExists(strDestino) Then
   Select Case blnSobreescribir
      Case True
         fso.DeleteFile strDestino
      Case False
         Err.Raise vbObjectError + 513, "CopiaArchivo", "El archivo destino """ & strOrigen & """ ya existe"
   End Select
End If
Set Archivo = fso.GetFile(strOrigen)
Select Case blnMover
   Case True
      fso.MoveFile strOrigen, strDestino
   Case False
      fso.CopyFile strOrigen, strDestino
End Select

CopiaArchivo = True


Salir:
   On Error Resume Next
   If Not fso Is Nothing Then Set fso = Nothing
   On Error GoTo 0
   Exit Function


TratamientoErrores:
   MsgBox "Error " & Err & ": " & Err.Description & vbNewLine & Switch(Erl = 0, vbNullString, Not Erl = 0, "En linea: " & Erl) & "en el procedimiento: CopiaArchivo, del Módulo: " & Application.VBE.ActiveCodePane.CodeModule.Name, vbCritical + vbOKOnly ', DBEngine(0)(0).Properties("AppTitle")
   Resume Salir
Resume Next
End Function

'---------------------------------------------------------------------------------------
' Function  : xlFolderExist
' DateTime  : 04/06/2014
' Author    : José García Herruzo
' Purpose   : Check if specified foleder exist
' Arguments :
'             FolderPath               --> folder to be checked
'---------------------------------------------------------------------------------------
Public Function xlFolderExist(ByVal FolderPath As String) As Boolean

Dim fso As FileSystemObject

Set fso = CreateObject("Scripting.FileSystemObject")

If fso.FolderExists(FolderPath) = True Then

    xlFolderExist = True

Else

    xlFolderExist = False
    
End If

End Function

'---------------------------------------------------------------------------------------
' Function  : DeleteFolder
' DateTime  : 18/06/2014
' Author    : José García Herruzo
' Purpose   : Delete selected folder
' Arguments :
'             str_Path                 --> Path + file name
'---------------------------------------------------------------------------------------
Public Sub DeleteFolder(ByVal str_Path As String)

Dim fso As Variant
Dim myFolder As Variant

On Error GoTo myhandler

'-- Set the new file system object --
Set fso = CreateObject("Scripting.FileSystemObject")

'-- delete desire folder --
fso.DeleteFolder str_Path, True
            
Set fso = Nothing
Set myFolder = Nothing

Exit Sub

myhandler:
    Set fso = Nothing
    Set myFolder = Nothing
    
    Error_FileAndFolder = 8
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : xsDeleteFile
' DateTime  : 23/06/2014
' Author    : José García Herruzo
' Purpose   : This procedure remove selected file
' Arguments :
'             str_myPath               --> Path + file name
'---------------------------------------------------------------------------------------
Public Sub xsDeleteFile(ByVal str_mypath As String)

Dim fso As Variant

On Error GoTo myhandler

Set fso = CreateObject("Scripting.FileSystemObject")
   
If fso.FileExists(str_mypath) Then
    
    fso.DeleteFile str_mypath, True

Else

    Set fso = Nothing
    
End If

Exit Sub

myhandler:
    Set fso = Nothing
    MsgBox ("Problema borrando archivo"), vbCritical
    Error_FileAndFolder = 9

End Sub
'---------------------------------------------------------------------------------------
' Function  : xlGetFolderSize
' DateTime  : 15/09/2014
' Author    : José García Herruzo
' Purpose   : Returns folder size in bytes
' Arguments :
'             str_myPath               --> folder Path
'---------------------------------------------------------------------------------------
Public Function xlGetFolderSize(ByVal str_mypath As String) As Double

Dim fso As Object
Dim fld As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set fld = fso.GetFolder(str_mypath)

On Error GoTo myhandler

xlGetFolderSize = fld.Size

myhandler:
    Set fso = Nothing
    Set fld = Nothing
    Error_FileAndFolder = 10
    xlGetFolderSize = 0
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : xlGetExcelFilePathName
' DateTime  : 10/23/2013
' Author    : José García Herruzo
' Purpose   : This function returns excel MEB file path &name
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function xlGetExcelFilePathName(Optional ls_Header As String) As Variant

Dim ls_Filter As String
Dim lv_FilePathName As Variant
Dim fso As Variant

Dim help11 As String

Set fso = CreateObject("Scripting.FileSystemObject")
        
        help11 = ThisWorkbook.Path
        ChDrive help11
        ChDir help11
    
    Set fso = Nothing
    
    ' -- assign values to string for File/Open dialogue box --
    ls_Filter = "Excel, *.xls,"
    If ls_Header = "" Then
    
        ls_Header = "Select Excel File"
    
    End If
    
    ' -- use Excel GetOpenFilename function to get the path and name
    '    of the Aspen Plus file --
    lv_FilePathName = Application.GetOpenFilename(ls_Filter, 2, ls_Header)
    
    ' -- assign filepathname to the name of the function --
    xlGetExcelFilePathName = lv_FilePathName
    
End Function

'---------------------------------------------------------------------------------------
' Function  : xf_FileExist
' DateTime  : 25/06/2012
' Author    : José García Herruzo
' Purpose   : Return tru if exist
' Arguments :
'             str_myPath               --> Path + file name
'---------------------------------------------------------------------------------------
Public Function xf_FileExist(ByVal str_mypath As String) As Boolean

Dim fso As Variant

On Error GoTo myhandler

Set fso = CreateObject("Scripting.FileSystemObject")
   
If fso.FileExists(str_mypath) Then
    
    Set fso = Nothing
    xf_FileExist = True
    
Else
    
    Set fso = Nothing
    xf_FileExist = False
    
End If

Exit Function

myhandler:
    Set fso = Nothing
    MsgBox ("Problema comprobando el archivo"), vbCritical
    Error_FileAndFolder = 9

End Function
