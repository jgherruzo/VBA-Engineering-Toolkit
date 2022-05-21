Attribute VB_Name = "ModTxT"
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
' Module    : ModTxT
' DateTime  : 05/29/2013
' Author    : José García Herruzo
' Purpose   : This module contents functions and procedures applied to txt files
' References: N/A
' Functions :
'               1-WriteTxTFile
' Procedures:
'               1-xlGetTxTContent
' Updates   :
'       DATE        USER    DESCRIPTION
'       18/02/2014  JGH     xlGetTxTContent is added
'----------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function  : WriteTxTFile
' DateTime  : 05/29/2013
' Author    : http://www.mrexcel.com/forum/excel-questions/535773-list-softwares-
'                installed-excel-using-visual-basic-applications.html
' Purpose   : write into a txt file
' Arguments :
'             sData                    --> String to be written
'             sFileName                --> file name
'---------------------------------------------------------------------------------------
Public Function WriteTxTFile(ByVal sData As String, ByVal sFileName As String) As Boolean

  Dim fso, OutFile, bWrite
  
  bWrite = True
  Set fso = CreateObject("Scripting.FileSystemObject")
  
  On Error Resume Next
  
  Set OutFile = fso.OpenTextFile(sFileName, 2, True)
  
  'Possibly need a prompt to close the file and one recursion attempt.
  If Err = 70 Then
  
    MsgBox "Could not write to file " & sFileName & ", results " & _
                 "not saved." & vbCrLf & vbCrLf & "This is probably " & _
                 "because the file is already open."
    bWrite = False
    
  ElseIf Err Then
  
    MsgBox Err & vbCrLf & Err.Description
    bWrite = False
    
  End If
  
  On Error GoTo 0
  
  If bWrite Then
  
    OutFile.WriteLine (sData)
    OutFile.Close
    
  End If
  
  Set fso = Nothing
  Set OutFile = Nothing
  
  WriteTxTFile = bWrite
  
End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetTxTContent
' DateTime  : 18/02/2014
' Author    : José García Herruzo
' Purpose   : Read from txt and return a string with the contents
' Arguments :
'             str_File                 --> file path $ name
'---------------------------------------------------------------------------------------
Public Function xlGetTxTContent(ByVal str_File As String) As String

Dim str_Code As String
Dim str_NewLine As String

Open str_File For Input As #1

Do Until EOF(1)

    Line Input #1, str_NewLine
    str_Code = str_Code & str_NewLine & vbCrLf
    
Loop

Close #1

xlGetTxTContent = str_Code

End Function

