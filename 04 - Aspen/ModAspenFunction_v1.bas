Attribute VB_Name = "ModAspenFunction_v1"
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
' Module    : ModAspenFunction_v1
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures for work with Aspen
' References: Aspen Plus GUI XX.X Type Library
' Functions :
'               1-IsSimLoaded
'               2-TellMeTheAspenVersion
'               3-aplGetCompStatus
'               4-xlIsStream
' Procedures:
'               1-OpenSimFile
'               2-CloseSimFile
'               3-OpenIpnFile
'               4-ExportIpnFile
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/22/2013  JGH     aplGetCompStatus function is added
'       17/02/2014  JGH     OpenIpnFile and ExportIpnFile procedures are developed
'       17/02/2014  JGH     OpenSimFile is modified to select aspen version
'       04/07/2014  JGH     str_AspenVersion is added
'       18/05/2016  JGH     xlIsStream is added
'----------------------------------------------------------------------------------------
Public lb_IsLoaded As Boolean
Public lb_IsVisible As Boolean
Public go_simulation As IHapp
Public obj_Aspen As HappLS
Public str_AspenVersion As String
'---------------------------------------------------------------------------------------
' Function  : IsSimLoaded
' DateTime  : 03/15/2013
' Author    : Aspen plus support. Modify by José García Herruzo
' Purpose   : function to check if a simulation has been loaded
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function IsSimLoaded() As Boolean

Dim lb_IsLoaded As Boolean

    ' -- special syntax for object variable to test if an object has been
    '    placed into the variable --
    If go_simulation Is Nothing Then
        lb_IsLoaded = False
    Else
        lb_IsLoaded = True
    End If ' go_Simulation Is Nothing
    
    IsSimLoaded = lb_IsLoaded
    
End Function
'---------------------------------------------------------------------------------------
' Function  : TellMeTheAspenVersion
' DateTime  : 08/04/2013
' Author    : José García Herruzo
' Purpose   : Return Aspen version in a string
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function TellMeTheAspenVersion() As String

Dim fso As Variant
Dim SourcePAth As Variant
Dim myFolder As Variant
Dim Part() As String

On Error GoTo TellMeTheAspenVersion_Error

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set SourcePAth = fso.GetFolder("C:\Archivos de programa\AspenTech")
    
        For Each myFolder In SourcePAth.SubFolders
        
            If InStr(1, myFolder.Name, "Aspen Plus", vbTextCompare) <> 0 Then
            
                Part = Split(myFolder.Name, "V")
                
                TellMeTheAspenVersion = Part(1)
                
                Exit Function
                
            End If
            
        Next

Exit Function
TellMeTheAspenVersion_Error:
MsgBox ("ModWorkbook: TellMeTheAspenVersion_Error " & Err.Number & ": " & Err.Description), vbCritical
MsgBox ("Aspen features will be deactivated"), vbInformation

FormMenu.cmdAspenLaunch.Enabled = False

End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenSimFile
' DateTime  : 06/03/2013
' Author    : Aspen plus support. Modify by José García Herruzo
' Purpose   : Open a simulation file
' Arguments :
'             lv_FilePathName          --> File path+file name
'             str_aspV                 --> Aspen version: "Apwn.Document.XX.X"
'---------------------------------------------------------------------------------------
Public Sub OpenSimFile(ByVal lv_FilePathName As String, ByVal str_aspV As String)
    
    ' -- change the cursor to a hourglass --
    Application.Cursor = xlWait
    
    ' -- check if the Cancel button was used in the
    '    File/Open dialogue box --
    If lv_FilePathName <> "" Then
        ' -- check to see if simulation is loaded --
        If IsSimLoaded = True Then
            ' -- set the simulation to Not Visible so it can be closed --
            go_simulation.Visible = False
            
            ' -- close simulation --
            Set go_simulation = Nothing
        End If 'IsSimLoaded = True
        
        ' -- open the selected Aspen Plus File --
        Set go_simulation = GetObject(lv_FilePathName, str_aspV)
    
        ' -- make the Aspen Plus GUI Visible/Not Visible based on check box value --
        go_simulation.Visible = True

    End If 'lv_FilePathName <> False
    
    ' -- change the cursor to default --
    Application.Cursor = xlDefault
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CloseSimFile
' DateTime  : 06/03/2013
' Author    : Aspen plus support. Modify by José García Herruzo
' Purpose   : Open a simulation file
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Sub CloseSimFile()

    ' -- check to see if simulation is loaded --
    If IsSimLoaded = True Then
    
        ' -- change the cursor to a hourglass --
        Application.Cursor = xlWait
        
        ' -- set the simulation to Not Visible so it can be closed --
        go_simulation.Visible = False
        
        ' -- close simulation --
        Set go_simulation = Nothing

        ' -- change cursor to default --
        Application.Cursor = xlDefault
    
    End If 'IsSimLoaded = True
    
End Sub

'---------------------------------------------------------------------------------------
' Procedure : aplGetCompStatus
' DateTime  : 07/22/2013
' Author    : Aspen plus support.
' Purpose   : Return sim status
' Arguments :
'             al_CompStatus               --> Node reference
'---------------------------------------------------------------------------------------
Public Function aplGetCompStatus(al_CompStatus As Long) As String
' -- function that read an input HAP_COMPSTATUS variable and
'    returns a descriptive string --

Dim ls_Return As String

    If ((al_CompStatus And HAP_RESULTS_SUCCESS) = HAP_RESULTS_SUCCESS) Then
        ls_Return = "Results Available"
    ElseIf ((al_CompStatus And HAP_NORESULTS) = HAP_NORESULTS) Then
        ls_Return = "No Results Available"
    ElseIf ((al_CompStatus And HAP_RESULTS_WARNINGS) = HAP_RESULTS_WARNINGS) Then
        ls_Return = "Results Available with Warnings"
    ElseIf ((al_CompStatus And HAP_RESULTS_INACCESS) = HAP_RESULTS_INACCESS) Then
        ls_Return = "Results In Access"
    ElseIf ((al_CompStatus And HAP_RESULTS_INCOMPAT) = HAP_RESULTS_INCOMPAT) Then
        ls_Return = "Results Incompatable"
    ElseIf ((al_CompStatus And HAP_RESULTS_ERRORS) = HAP_RESULTS_ERRORS) Then
        ls_Return = "Results Available with Errors"
    End If
    aplGetCompStatus = ls_Return
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : OpenIpnFile
' DateTime  : 17/02/2014
' Author    : Emmanuel Lejeune. Modify by José García Herruzo
' Purpose   : Open an input file like simulation file
' Arguments :
'             lv_FilePathName          --> File path+file name
'---------------------------------------------------------------------------------------
Public Sub OpenIpnFile(ByVal lv_FilePathName As String)
    
    ' -- change the cursor to a hourglass --
    Application.Cursor = xlWait
    
    ' -- check if the Cancel button was used in the
    '    File/Open dialogue box --
    If lv_FilePathName <> "" Then
        ' -- check to see if simulation is loaded --
        If IsSimLoaded = True Then
            ' -- set the simulation to Not Visible so it can be closed --
            obj_Aspen.Visible = False
            
            ' -- close simulation --
            Set obj_Aspen = Nothing
            
        End If 'IsSimLoaded = True
        
        ' -- create aspen object --
        Set obj_Aspen = GetObject("Apwn.Documents")
        
        ' -- open from ipn File --
        obj_Aspen.InitFromFile2 (lv_FilePathName)
        
        ' -- make the Aspen Plus GUI Visible/Not Visible based on check box value --
        obj_Aspen.Visible = lb_IsVisible

    End If 'lv_FilePathName <> False
    
    ' -- change the cursor to default --
    Application.Cursor = xlDefault
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : ExportIpnFile
' DateTime  : 17/02/2014
' Author    : Emmanuel Lejeune. Modify by José García Herruzo
' Purpose   : Generate an input file from simulation file
' Arguments :
'             lv_FilePathName          --> File path+file name
'---------------------------------------------------------------------------------------
Public Sub ExportIpnFile(ByVal lv_FilePathName As String)
    
    obj_Aspen.Export HAPEXP_INPUT, lv_FilePathName
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : xlIsStream
' DateTime  : 18/05/2016
' Author    : David Cassaseca. Modify by José García Herruzo
' Purpose   : Check if a stream is in the sim file
' Arguments :
'             ao_Simulation            --> Sim object
'             str_myTag                --> Stream tag
'---------------------------------------------------------------------------------------
Public Function xlIsStream(ByVal ao_Simulation As IHapp, ByVal str_myTag As String) As Boolean
    
Dim lb_exist As Boolean
Dim HN_node As IHNode
Dim Temp As Variant

Set HN_node = ao_Simulation.Tree.FindNode("\Data\Streams")
lb_exist = False

For Each Temp In HN_node.Elements

    If Temp.Name = str_myTag Then
    
        lb_exist = True
    
    End If
    
Next Temp

xlIsStream = lb_exist

End Function


