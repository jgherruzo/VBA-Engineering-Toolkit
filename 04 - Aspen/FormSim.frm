VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FormSim 
   Caption         =   "APE Menu"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   OleObjectBlob   =   "FormSim.frx":0000
   StartUpPosition =   2  'Centrar en pantalla
End
Attribute VB_Name = "FormSim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

'---------------------------------------------------------------------------------------
' Event     : cmdClose_Click
' Author    : José García Herruzo
' Purpose   : Unload form and close this file
'---------------------------------------------------------------------------------------
Private Sub cmdClose_Click()

Application.DisplayAlerts = False
ThisWorkbook.Save
Unload Me
ThisWorkbook.Close
Application.DisplayAlerts = True

End Sub

'---------------------------------------------------------------------------------------
' Event     : cmdRun_Click
' Author    : José García Herruzo
' Purpose   : Simulate stream conditions and extract selected properties
'---------------------------------------------------------------------------------------
Private Sub cmdRun_Click()

Dim LastRow As Integer
Dim wb_Bal As Workbook
Dim str_help As String
Dim str_help2 As String

'- Extract path --
str_help2 = tv_Files.SelectedItem.Key
str_help = tv_Files.Nodes(str_help2).FullPath
'-- Make visible or not sim file --
lb_IsVisible = chk_IsVisible.Value

'-- Hide form and show process form--
Me.Hide
FormProgress.Show vbModeless
'-- Update progress-
Call UpdateP1("Opening excel file", 1, 5)

'-- Excel file will not be updated and deactivate msgbox --
Call Update_Data_Into_ExcelCell(str_BalPath & "\" & str_help, "Setup", "1", "Update Flag")
Call Update_Data_Into_ExcelCell(str_BalPath & "\" & str_help, "Setup", "1", "Msg Flag")
'-- Open selected balance file --
Set wb_Bal = Workbooks.Open(str_BalPath & "\" & str_help)

'-- Update progress --
Call UpdateP1("Opening simulation file", 2, 5)
'-- Open selected Sim file --
Call OpenSimFile(str_SimPath & "\" & tv_Files.Nodes(str_help2).Parent & "_vF\" & L_SimName.Caption)

'-- Update progress --
Call UpdateP1("Extracting properties", 3, 5)
'-- Run selected Sim file --
Call RunAPE(tv_Files.Nodes(str_help2).Parent, wb_Bal)

'-- Update progress --
Call UpdateP1("Closing simulation file ", 4, 5)
'-- Close selected Sim file --
Call CloseSimFile

'-- Update progress --
Call UpdateP1("Closing excel file ", 5, 5)
Application.DisplayAlerts = False
'-- Restar Excel status, save it and close it --
Call Update_Data_Into_ExcelCell(str_BalPath & "\" & str_help, "Setup", "0", "Update Flag")
Call Update_Data_Into_ExcelCell(str_BalPath & "\" & str_help, "Setup", "0", "Msg Flag")
wb_Bal.Save
wb_Bal.Close
Application.DisplayAlerts = True

'-- Add new log note --
LastRow = ws_Log.Range("A16000").End(xlUp).Row
LastRow = LastRow - 1
ws_Log.Range("A2").Offset(LastRow, 0).Value = tv_Files.Nodes(str_help2).Tag
ws_Log.Range("A2").Offset(LastRow, 1).Value = Now

Set wb_Bal = Nothing

'-- Unload progress form and show main form --
Unload FormProgress
Me.Show

End Sub

'---------------------------------------------------------------------------------------
' Event     : CommandButton1_Click
' Author    : José García Herruzo
' Purpose   : Unload form in order to work with excel file
'---------------------------------------------------------------------------------------
Private Sub CommandButton1_Click()

Unload Me

End Sub

'---------------------------------------------------------------------------------------
' Event     : tv_files_Click
' Author    : José García Herruzo
' Purpose   : Search if simulation file is available for selected excel file
'---------------------------------------------------------------------------------------
Private Sub tv_files_Click()

Dim LastRow As Integer
Dim str_NodeRootPath As String
Dim str_NodeRootKey As String
Dim i As Integer

'-- Reset log list box --
lbo_SimLog.Clear
    
str_NodeRootKey = tv_Files.SelectedItem.Key

str_NodeRootPath = tv_Files.Nodes(str_NodeRootKey).FullPath

'-- Look for excel file and simulation file --
Call GetFileInfo(str_NodeRootPath)

Call AvailableSimFile

'-- Seach log information about selected file
LastRow = ws_Log.Range("A16000").End(xlUp).Row
LastRow = LastRow - 2

If L_BalanceName.Caption <> "" Then

    For i = 0 To LastRow
    
        If ws_Log.Range("A2").Offset(i, 0).Value = tv_Files.Nodes(str_NodeRootKey).Tag Then
        
            lbo_SimLog.AddItem ws_Log.Range("A2").Offset(i, 1).Value
        
        End If
    
    Next i

End If

End Sub

'---------------------------------------------------------------------------------------
' Event     : UserForm_Initialize
' Author    : José García Herruzo
' Purpose   : Shows availables Excel balance files into a treeview
'---------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()

Dim str_myFolders() As String
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim str_myHelpPath As String
Dim str_myFiles() As String
Dim str_myFileName() As String
Dim myNode As Node
Dim myNode2 As Node

Frame_Sim.Enabled = False
L_LockRuncmd.Visible = True
    
'-- Extract availables folders --
str_myFolders() = Split(GetSubfolder(str_BalPath), "\")

    For i = 1 To UBound(str_myFolders) - 2
        
        '-- Add each foler into a node --
        Set myNode2 = AddParentNode(tv_Files, str_myFolders(i), "ID " & i)
        str_myHelpPath = str_BalPath & "\" & str_myFolders(i) & "\"
        
        '-- Extract available files into a selected folder --
        str_myFiles() = Split(GetFiles(str_myHelpPath), "\")
        
        For j = 1 To UBound(str_myFiles)
            
            str_myFileName() = Split(str_myFiles(j), ".")
            '-- Add children node for each file into a parent node folder --
            Call AddChildrenNode(tv_Files, myNode2, str_myFiles(j), "ID " & i & " " & j)
            
        Next j
        
    Next i

End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetFileInfo
' Author    : José García Herruzo
' Purpose   : Search if simulation file is available for selected excel file
'---------------------------------------------------------------------------------------
Private Sub GetFileInfo(ByVal mystr_NodeRootPath As String)

Dim myInfo() As String
Dim myInfo2() As String
Dim str_myHelpPath As String
Dim str_myFiles() As String
Dim i As Integer

'-- Reset label--
L_SimName.Caption = ""
L_BalanceName.Caption = ""

'-- Extract Excel file name --
myInfo() = Split(mystr_NodeRootPath, "\")

'-- It is not a Excel file --
If UBound(myInfo) < 1 Then

    Exit Sub

End If

'-- Extract file information --
If UBound(myInfo) >= 1 Then
    myInfo2() = Split(myInfo(1), ".")
    
    If myInfo2(2) <> "" Then
        '-- Excel complete name --
        L_BalanceName.Caption = myInfo(1)
        str_myHelpPath = str_SimPath & "\" & myInfo2(0) & "_vF\"
        
        '-- Availables files into a sim folder --
        str_myFiles() = Split(GetFiles(str_myHelpPath), "\")
        
        '-- Search and show bkp file --
        For i = 1 To UBound(str_myFiles)
        
            If InStr(str_myFiles(i), ".bkp") <> 0 Then
            
                L_SimName.Caption = str_myFiles(i)
            
            End If
        
        Next i
    
    End If
    
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : AvailableSimFile
' Author    : José García Herruzo
' Purpose   : Enable or not run simulation option if sim file is available or not
'---------------------------------------------------------------------------------------
Private Sub AvailableSimFile()

If L_SimName.Caption <> "" Then

    Frame_Sim.Enabled = True
    L_LockRuncmd.Visible = False

Else
    
    Frame_Sim.Enabled = False
    L_LockRuncmd.Visible = True

End If

End Sub

'---------------------------------------------------------------------------------------
' Event     : UserForm_QueryClose
' Author    : José García Herruzo
' Purpose   : You cannot close this form from X button
'---------------------------------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
     
    If CloseMode = 0 Then
        Cancel = True
        MsgBox "The X is disabled, please use a button on the form.", vbCritical
    End If
     
End Sub
