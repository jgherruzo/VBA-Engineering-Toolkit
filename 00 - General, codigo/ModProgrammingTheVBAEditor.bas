Attribute VB_Name = "ModProgrammingTheVBAEditor"
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
' Module    : ModProgrammingTheVBAEditor
' DateTime  : 03/01/2014
' Author    : José García Herruzo; based on http://www.cpearson.com/excel/vbe.aspx
' Purpose   : This module contents functions and procedures to work with VBA editor
' References: Microsoft Visual Basic For Applications Extensibility 5.3
'              Remind Macros configuration.....
' Functions :
'               1-xlGetModules
'               2-xlGetDocModules
'               3-xlCodeLocation
' Procedures:
'               1-xsResetCode
'               2-xsRemoveModule
'               3-xsDeleteCode
'               4-xsDeleteProcedure
'               5-xsAddCode
'               6-xsReplaceCodeLine
'               7-ExportVBComponent
'               8-GetFileExtension
' Updates   :
'       DATE        USER    DESCRIPTION
'       18/02/2014  JGH     Procedures 2 and 3, function 1 and 2, are added
'       19/02/2014  JGH     Procedures 4 and 6 are, function 3, are added
'       18/05/2015  JGH     Functions 7 and 8 are added
'----------------------------------------------------------------------------------------
Dim VBAEditor As VBIDE.VBE
Dim VBProj As VBIDE.VBProject
Dim VBComps As VBIDE.VBComponents
Dim VBComp As VBIDE.VBComponent
Dim CodeMod As VBIDE.CodeModule
'---------------------------------------------------------------------------------------
' Procedure : xsRemoveAllModules
' DateTime  : 02/01/2014
' Author    : José García Herruzo
' Purpose   : This procedure remove each module on the selected workbook
' Arguments :
'             myWb                       --> Selected workbook
'---------------------------------------------------------------------------------------
Public Sub xsResetCode(ByRef mywb As Workbook)

Set VBComps = mywb.VBProject.VBComponents

For Each VBComp In VBComps

    If VBComp.Type <> vbext_ct_Document Then
        
            VBComps.Remove VBComp
            
    Else
        
        With VBComp.CodeModule
        
            .DeleteLines 1, .CountOfLines
            
        End With
        
    End If
    
Next VBComp
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : xlGetModules
' DateTime  : 18/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure returns all modules in the selected workbook
' Arguments :
'             myWb                       --> Selected workbook
'---------------------------------------------------------------------------------------
Public Function xlGetModules(ByRef mywb As Workbook) As String

Dim i As Integer
Dim int_ModCounter As Integer
Dim int_ModNumber As Integer
Dim int_Mark As Integer

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- Count the number of modules to be removed --
int_ModNumber = VBProj.VBComponents.Count - 1
int_ModCounter = 0

For i = 1 To int_ModNumber + 1
        
    If VBProj.VBComponents.Item(i).Type <> vbext_ct_Document Then
    
        int_ModCounter = 1 + int_ModCounter
    
    End If

Next i

If int_ModCounter < 0 Then

    xlGetModules = ""

ElseIf int_ModCounter = 0 Then

    For i = 1 To int_ModNumber + 1
        
        If VBProj.VBComponents.Item(i).Type <> vbext_ct_Document Then
        
            xlGetModules = VBProj.VBComponents.Item(i).Name
            Exit For
        
        End If
    
    Next i

Else
    
    For i = 1 To int_ModNumber + 1
        
        If VBProj.VBComponents.Item(i).Type <> vbext_ct_Document Then
        
            xlGetModules = VBProj.VBComponents.Item(i).Name
            int_Mark = i + 1
            Exit For
        
        End If
    
    Next i
    
    For i = int_Mark To int_ModNumber + 1
    
        If VBProj.VBComponents.Item(i).Type <> vbext_ct_Document Then
        
            xlGetModules = xlGetModules & "\" & VBProj.VBComponents.Item(i).Name
        
        End If
        
    Next i

End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : xlGetDocModules
' DateTime  : 18/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure returns all modules in the selected workbook
' Arguments :
'             myWb                       --> Selected workbook
'---------------------------------------------------------------------------------------
Public Function xlGetDocModules(ByRef mywb As Workbook) As String

Dim i As Integer
Dim int_ModCounter As Integer
Dim int_ModNumber As Integer
Dim int_Mark As Integer

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- Count the number of modules to be removed --
int_ModNumber = VBProj.VBComponents.Count - 1
int_ModCounter = 0

For i = 1 To int_ModNumber + 1
        
    If VBProj.VBComponents.Item(i).Type = vbext_ct_Document Then
    
        int_ModCounter = 1 + int_ModCounter
    
    End If

Next i

If int_ModCounter < 0 Then

    xlGetDocModules = ""

ElseIf int_ModCounter = 0 Then

    For i = 1 To int_ModNumber + 1
        
        If VBProj.VBComponents.Item(i).Type = vbext_ct_Document Then
        
            xlGetDocModules = VBProj.VBComponents.Item(i).Name
            Exit For
        
        End If
    
    Next i

Else
    
    For i = 1 To int_ModNumber + 1
        
        If VBProj.VBComponents.Item(i).Type = vbext_ct_Document Then
        
            xlGetDocModules = VBProj.VBComponents.Item(i).Name
            int_Mark = i + 1
            Exit For
        
        End If
    
    Next i
    
    For i = int_Mark To int_ModNumber + 1
    
        If VBProj.VBComponents.Item(i).Type = vbext_ct_Document Then
        
            xlGetDocModules = xlGetDocModules & "\" & VBProj.VBComponents.Item(i).Name
        
        End If
        
    Next i

End If
    
End Function
'---------------------------------------------------------------------------------------
' Procedure : xsRemoveModule
' DateTime  : 18/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure remove a specific module
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module to be removed
'---------------------------------------------------------------------------------------
Public Sub xsRemoveModule(ByRef mywb As Workbook, ByVal str_myModule As String)

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)

If VBComp.Type = vbext_ct_ClassModule Or VBComp.Type = vbext_ct_StdModule Then
    Application.DisplayAlerts = True
    VBProj.VBComponents.Remove VBComp
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : xsDeleteCode
' DateTime  : 18/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure remove the code
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module to be removed
'---------------------------------------------------------------------------------------
Public Sub xsDeleteCode(ByRef mywb As Workbook, ByVal str_myModule As String)

Dim StartLine As Long
Dim NumLines As Long
Dim ProcName As String

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

If VBComp.Type = vbext_ct_Document Then

    With CodeMod
    
        .DeleteLines 1, .CountOfLines
        
    End With

End If
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : xsAddCode
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure add code
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module to be removed
'             str_Code                   --> Code to be included
'---------------------------------------------------------------------------------------
Public Sub xsAddCode(ByRef mywb As Workbook, ByVal str_myModule As String, ByVal str_Code As String)

Dim LineNum As Long

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

LineNum = CodeMod.CountOfLines + 1
CodeMod.InsertLines LineNum, str_Code
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : xsDeleteProcedure
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure remove a specific procedure
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module where search the procedure
'             str_Procedure              --> Procedure to be removed
'---------------------------------------------------------------------------------------
Public Sub xsDeleteProcedure(ByRef mywb As Workbook, ByVal str_myModule As String, ByVal str_Procedure As String)
 
Dim StartLine As Long
Dim NumLines As Long

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

With CodeMod
    StartLine = .ProcStartLine(str_Procedure, vbext_pk_Proc)
    NumLines = .ProcCountLines(str_Procedure, vbext_pk_Proc)
    .DeleteLines StartLine:=StartLine, Count:=NumLines
End With
        
End Sub
'---------------------------------------------------------------------------------------
' Procedure : xsReplaceCodeLine
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure replace a code line
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module where search the code
'             str_KeyString              --> Key string to locate the line
'             str_NewLine                --> New line to be included
'---------------------------------------------------------------------------------------
Public Sub xsReplaceCodeLine(ByRef mywb As Workbook, ByVal str_myModule As String, _
                                ByVal str_KeyString As String, ByVal str_NewLine As String)
 
Dim lon_StartLine As Long
Dim lon_EndLine As Long
Dim lon_StartColumn As Long
Dim lon_EndColumn As Long
Dim bol_IsFound As Boolean

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

With CodeMod
    lon_StartLine = 1
    lon_EndLine = .CountOfLines
    lon_StartColumn = 1
    lon_EndColumn = 255
    bol_IsFound = .Find(target:=str_KeyString, StartLine:=lon_StartLine, StartColumn:=lon_StartColumn, _
        EndLine:=lon_EndLine, EndColumn:=lon_EndColumn, _
        wholeword:=True, MatchCase:=False, patternsearch:=False)
        
    .DeleteLines StartLine:=str_KeyString, Count:=1
    .InsertLines str_KeyString, str_NewLine
    
End With

End Sub
'---------------------------------------------------------------------------------------
' Function  : xlCodeLocation
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : This procedure returns word location
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module where search the code
'             str_KeyString              --> Key string to locate the line
'---------------------------------------------------------------------------------------
Public Function xlCodeLocation(ByRef mywb As Workbook, ByVal str_myModule As String, ByVal str_KeyString As String) As String
 
Dim lon_StartLine As Long
Dim lon_EndLine As Long
Dim lon_StartColumn As Long
Dim lon_EndColumn As Long
Dim bol_IsFound As Boolean

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

With CodeMod
    lon_StartLine = 1
    lon_EndLine = .CountOfLines
    lon_StartColumn = 1
    lon_EndColumn = 255
    bol_IsFound = .Find(target:=str_KeyString, StartLine:=lon_StartLine, StartColumn:=lon_StartColumn, _
        EndLine:=lon_EndLine, EndColumn:=lon_EndColumn, _
        wholeword:=True, MatchCase:=False, patternsearch:=False)
End With

If bol_IsFound = True Then

    xlCodeLocation = lon_StartLine
    
Else

    xlCodeLocation = ""
    
End If

End Function
'---------------------------------------------------------------------------------------
' Function  : xlGetLineCode1
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : Returns the first line which contents key string
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module where search the code
'             str_KeyString              --> Key string to locate the line
'---------------------------------------------------------------------------------------
Public Function xlGetLineCode1(ByRef mywb As Workbook, ByVal str_myModule As String, ByVal str_KeyString As String) As String
 
Dim lon_StartLine As Long
Dim lon_EndLine As Long
Dim lon_StartColumn As Long
Dim lon_EndColumn As Long
Dim bol_IsFound As Boolean
Dim str_Code As String

'-- set the visual basic project --
Set VBProj = mywb.VBProject
'-- set module --
Set VBComp = VBProj.VBComponents(str_myModule)
'-- set code --
Set CodeMod = VBComp.CodeModule

With CodeMod
    lon_StartLine = 1
    lon_EndLine = .CountOfLines
    lon_StartColumn = 1
    lon_EndColumn = 255
    bol_IsFound = .Find(target:=str_KeyString, StartLine:=lon_StartLine, StartColumn:=lon_StartColumn, _
        EndLine:=lon_EndLine, EndColumn:=lon_EndColumn, _
        wholeword:=True, MatchCase:=False, patternsearch:=False)
End With

If bol_IsFound = True Then

    xlGetLineCode1 = CodeMod.Lines(lon_StartLine, 1)

Else

    xlGetLineCode1 = ""

End If
    
End Function
'---------------------------------------------------------------------------------------
' Function  : xlModuleExists
' DateTime  : 19/02/2014
' Author    : José García Herruzo
' Purpose   : Returns true if module exist
' Arguments :
'             myWb                       --> Selected workbook
'             str_myModule               --> Module where search the code
'---------------------------------------------------------------------------------------
Public Function xlModuleExists(ByRef mywb As Workbook, ByVal str_myModule As String) As Boolean

'-- set the visual basic project --
Set VBProj = mywb.VBProject

On Error Resume Next
xlModuleExists = CBool(Len(VBProj.VBComponents(str_myModule).Name))

End Function
'---------------------------------------------------------------------------------------
' Function  : ExportVBComponent
' DateTime  : 18/05/2015
' Author    : cpearson
' Purpose   : Export module to a txt
' Arguments :
'             VBComp                     --> Selected workbook
'             FolderName                 --> Path
'             FileName                   --> File name
'             OverwriteExisting          --> Optional argument
'---------------------------------------------------------------------------------------
Public Function ExportVBComponent(VBComp As VBIDE.VBComponent, FolderName As String, Optional FileName As String, _
            Optional OverwriteExisting As Boolean = True) As Boolean
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This function exports the code module of a VBComponent to a text
' file. If FileName is missing, the code will be exported to
' a file with the same name as the VBComponent followed by the
' appropriate extension.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim Extension As String
Dim FName As String
Extension = GetFileExtension(VBComp:=VBComp)
If Trim(FileName) = vbNullString Then
    FName = VBComp.Name & Extension
Else
    FName = FileName
    If InStr(1, FName, ".", vbBinaryCompare) = 0 Then
        FName = FName & Extension
    End If
End If

If StrComp(Right(FolderName, 1), "\", vbBinaryCompare) = 0 Then
    FName = FolderName & FName
Else
    FName = FolderName & "\" & FName
End If

If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
    If OverwriteExisting = True Then
        Kill FName
    Else
        ExportVBComponent = False
        Exit Function
    End If
End If

VBComp.Export FileName:=FName
ExportVBComponent = True

End Function
'---------------------------------------------------------------------------------------
' Function  : GetFileExtension
' DateTime  : 18/05/2015
' Author    : cpearson
' Purpose   : Export module to a txt
' Arguments :
'             VBComp                     --> module
'---------------------------------------------------------------------------------------
Private Function GetFileExtension(VBComp As VBIDE.VBComponent) As String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' This returns the appropriate file extension based on the Type of
' the VBComponent.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Select Case VBComp.Type
        Case vbext_ct_ClassModule
            GetFileExtension = ".cls"
        Case vbext_ct_Document
            GetFileExtension = ".cls"
        Case vbext_ct_MSForm
            GetFileExtension = ".frm"
        Case vbext_ct_StdModule
            GetFileExtension = ".bas"
        Case Else
            GetFileExtension = ".bas"
    End Select
    
End Function

