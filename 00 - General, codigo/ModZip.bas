Attribute VB_Name = "ModZip"
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
' Module    : ModZip
' DateTime  : 18/06/2014
' Author    : José García Herruzo
' Purpose   : This module contents procedures and function to zip & zip files
' References: N/A
' Functions :
'               1-xlIs7ZipAvailable
' Procedures:
'               1-xsZipAllBrowse
'               2-xsZipAllIn
'               3-xsZipAllSelected
' Status    : OPENED
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Public Error_Zip As Integer
'--------------------------- ERROR INFORMATION BOX --------------------------------------
'   KEY         FUNTION or PROCEDURE                                                    '
'   1           xsZipAllBrowse                                                          '
'   2           xsZipAllIn                                                              '
'   3           xsZipAllSelected                                                        '
'------------------------- END ERROR INFORMATION BOX ------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Dim PathZipProgram As String
'--------------------------- Zip INFORMATION BOX --------------------------------------
    'Zip all the files in the folder and subfolders, -r is Include subfolders
    'ShellStr = PathZipProgram & "7z.exe a -r" _
    '         & " " & Chr(34) & NameZipFile & Chr(34) _
    '         & " " & Chr(34) & FolderName & "*.*" & Chr(34)

    'Note: you can replace the ShellStr with one of the example ShellStrings
    'below to test one of the examples


    'Zip the txt files in the folder and subfolders, use "*.xl*" for all excel files
    '        ShellStr = PathZipProgram & "7z.exe a -r" _
             '                 & " " & Chr(34) & NameZipFile & Chr(34) _
             '                 & " " & Chr(34) & FolderName & "*.txt" & Chr(34)

    'Zip all files in the folder and subfolders with a name that start with Week
    '        ShellStr = PathZipProgram & "7z.exe a -r" _
             '                 & " " & Chr(34) & NameZipFile & Chr(34) _
             '                 & " " & Chr(34) & FolderName & "Week*.*" & Chr(34)

    'Zip every file with the name ron.xlsx in the folder and subfolders
    '        ShellStr = PathZipProgram & "7z.exe a -r" _
             '                 & " " & Chr(34) & NameZipFile & Chr(34) _
             '                 & " " & Chr(34) & FolderName & "ron.xlsx" & Chr(34)

    'Add -ppassword -mhe of you want to add a password to the zip file(only .7z files)
    '                ShellStr = PathZipProgram & "7z.exe a -r -ppassword -mhe" _
                     '                                  & " " & Chr(34) & NameZipFile & Chr(34) _
                     '                                  & " " & Chr(34) & FolderName & "*.*" & Chr(34)

    'Add -seml if you want to open a mail with the zip attached
    '                ShellStr = PathZipProgram & "7z.exe a -r -seml" _
                     '                                  & " " & Chr(34) & NameZipFile & Chr(34) _
                     '                                  & " " & Chr(34) & FolderName & "*.*" & Chr(34)
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
'---------------------------------------------------------------------------------------
' Function  : xlIs7ZipAvailable
' DateTime  : 18/06/2014
' Author    : Ron de Bruin ; Modified by José García Herruzo
' Purpose   : Return true if 7-zip is installed
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Function xlIs7ZipAvailable() As Boolean

'Path of the Zip program
PathZipProgram = "C:\program files\7-Zip\"
If Right(PathZipProgram, 1) <> "\" Then
    PathZipProgram = PathZipProgram & "\"
End If

'Check if this is the path where 7z is installed.
If Dir(PathZipProgram & "7z.exe") = "" Then
    xlIs7ZipAvailable = False
    Exit Function
End If

xlIs7ZipAvailable = True

End Function

'---------------------------------------------------------------------------------------
' Function  : xsZipAllBrowse
' DateTime  : 18/06/2014
' Author    : Ron de Bruin ; Modified by José García Herruzo
' Purpose   : In this procedure you browse to the folder you want to zip, and zip it with
'               the same name in the same folder
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub xsZipAllBrowse()

Dim NameZipFile As String
Dim FolderName As String
Dim ShellStr As String
Dim DefPath As String
Dim Fld As Object
Dim Arr_str() As String
Dim i As Integer

On Error GoTo myhandler:

'Check if this is the path where 7z is installed.
If xlIs7ZipAvailable = False Then
    Error_Zip = 1
    Exit Sub
End If

'Browse to the folder with the files that you want to Zip
Set Fld = CreateObject("Shell.Application").BrowseForFolder(0, "Select folder to Zip", 512)
If Not Fld Is Nothing Then
    FolderName = Fld.Self.Path
    If Right(FolderName, 1) <> "\" Then
        FolderName = FolderName & "\"
    End If
    
    Arr_str = Split(FolderName, "\")
    DefPath = Arr_str(0)
    If Right(FolderName, 1) = "\" Then
        
        For i = 1 To UBound(Arr_str) - 2
        
            DefPath = DefPath & "\" & Arr_str(i)
        
        Next i
    
    Else
    
        For i = 1 To UBound(Arr_str) - 1
        
            DefPath = DefPath & "\" & Arr_str(i)
        
        Next i
        
    End If
    
If Right(DefPath, 1) <> "\" Then
    DefPath = DefPath & "\"
End If

'Set NameZipFile to the full path/name of the Zip file
NameZipFile = DefPath & Fld & ".zip"

    'Zip all the files in the folder and subfolders
    ShellStr = PathZipProgram & "7z.exe a -r" _
             & " " & Chr(34) & NameZipFile & Chr(34) _
             & " " & Chr(34) & FolderName & "*.*" & Chr(34)

    ShellAndWait ShellStr, vbHide

End If

Exit Sub
myhandler:
    Error_Zip = 1
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : xsZipAllIn
' DateTime  : 18/06/2014
' Author    : Ron de Bruin ; Modified by José García Herruzo
' Purpose   : In this procedure you browse to the folder you want to zip, and zip it with
'               the same name in the same folder
' Arguments :
'             str_DestinyPath            --> Path where you want file will be saved
'             str_SourcePath             --> Path including folder to be zipped
'             str_DestinyName            --> Zip file name
'---------------------------------------------------------------------------------------
Public Sub xsZipAllIn(ByVal str_SourcePath As String, ByVal str_DestinyPath As String, ByVal str_DestinyName As String)

Dim NameZipFile As String
Dim FolderName As String
Dim ShellStr As String

On Error GoTo myhandler:

'Check if this is the path where 7z is installed.
If xlIs7ZipAvailable = False Then
    Error_Zip = 2
    Exit Sub
End If

If Right(str_SourcePath, 1) <> "\" Then
    str_SourcePath = str_SourcePath & "\"
End If

If Right(str_DestinyPath, 1) <> "\" Then
    str_DestinyPath = str_DestinyPath & "\"
End If

NameZipFile = str_DestinyPath & str_DestinyName & ".zip"

'Zip all the files in the folder and subfolders, -r is Include subfolders
ShellStr = PathZipProgram & "7z.exe a -r" _
         & " " & Chr(34) & NameZipFile & Chr(34) _
         & " " & Chr(34) & str_SourcePath & "*.*" & Chr(34)

ShellAndWait ShellStr, vbHide

Exit Sub
myhandler:
    Error_Zip = 2
    
End Sub
'---------------------------------------------------------------------------------------
' Function  : xsZipAllSelected
' DateTime  : 18/06/2014
' Author    : Ron de Bruin ; Modified by José García Herruzo
' Purpose   : With this example you browse to the folder you want and select the files
'               that you want to zip. Use the Ctrl key to select more then one file or
'               select blocks of files with the shift key pressed. With Ctrl a you
'               select all files in the dialog.
' Arguments :
'             str_DestinyPath            --> Path where you want file will be saved
'             str_SourcePath             --> Path including folder to be zipped
'             str_DestinyName            --> Zip file name
'---------------------------------------------------------------------------------------
Public Sub xsZipAllSelected()
    Dim PathZipProgram As String, NameZipFile As String, FolderName As String
    Dim ShellStr As String, strDate As String, DefPath As String
    Dim NameList As String, sFileNameXls As String
    Dim vArr As Variant, FileNameXls As Variant, iCtr As Long

On Error GoTo myhandler:

'Check if this is the path where 7z is installed.
If xlIs7ZipAvailable = False Then
    Error_Zip = 3
    Exit Sub
End If

    DefPath = Application.DefaultFilePath
    If Right(DefPath, 1) <> "\" Then
        DefPath = DefPath & "\"
    End If

    'Create date/Time string, also the name of the Zip in this example
    strDate = Format(Now, "yyyy-mm-dd h-mm-ss")

    'Set NameZipFile to the full path/name of the Zip file
    'If you want to add the word "MyZip" before the date/time use
    'NameZipFile = DefPath & "MyZip " & strDate & ".zip"
    NameZipFile = DefPath & strDate & ".zip"

    FileNameXls = Application.GetOpenFilename(filefilter:="Excel Files, *.xl*", _
                                              MultiSelect:=True, Title:="Select the files that you want to add to the new zip file")

    If IsArray(FileNameXls) = False Then
        'do nothing
    Else
        NameList = ""
        For iCtr = LBound(FileNameXls) To UBound(FileNameXls)
            NameList = NameList & " " & Chr(34) & FileNameXls(iCtr) & Chr(34)
            vArr = Split(FileNameXls(iCtr), "\")
            sFileNameXls = vArr(UBound(vArr))

            If bIsBookOpen(sFileNameXls) Then
                MsgBox "You can't zip a file that is open!" & vbLf & _
                       "Please close: " & FileNameXls(iCtr)
                Exit Sub
            End If
        Next iCtr

        'Zip every file you have selected with GetOpenFilename
        ShellStr = PathZipProgram & "7z.exe a" _
                 & " " & Chr(34) & NameZipFile & Chr(34) _
                 & " " & NameList

        ShellAndWait ShellStr, vbHide

    End If

Exit Sub
myhandler:
    Error_Zip = 3
    
End Sub
