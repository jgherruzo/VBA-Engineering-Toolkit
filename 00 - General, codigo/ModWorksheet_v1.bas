Attribute VB_Name = "ModWorksheet_v1"
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
' Module    : ModWorksheet_v1
' DateTime  : 05/20/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures to work with worksheets
' References: N/A
' Functions :
'               1-SheetExist
'               2-ReturnColumn
'               3-Get_Column_Letter
' Procedures:
'               1-AddNewSheet
'               2-WriteFromTxT
' Updates   :
'       DATE        USER    DESCRIPTION
'       08/16/2013  JGH     Log is added
'       10/24/2013  JGH     Log is eliminated
'----------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function  : SheetExist
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Close workbooks without saving it
' Arguments :
'             str_WB_Name                --> Workbooks name
'             str_WS_Name                --> Worksheets name
'---------------------------------------------------------------------------------------
Public Function SheetExist(ByVal str_WB_Name As String, ByVal str_WS_Name As String) As Boolean

Dim h As Integer

On Error GoTo myhandler

Windows(str_WB_Name).Activate

For h = 1 To Sheets.Count

    If Sheets(h).Name = str_WS_Name Then
    
        SheetExist = True
        Exit Function
        
    Else
    
        SheetExist = False
        
    End If

Next h

Exit Function
myhandler:
            
End Function

'---------------------------------------------------------------------------------------
' Procedure : AddNewSheet
' DateTime  : 03/15/2013
' Author    : José García Herruzo
' Purpose   : Close workbooks without saving it
' Arguments :
'             myWB                       --> Workbooks name
'             mySheetName                --> Worksheets name
'---------------------------------------------------------------------------------------
Public Sub AddNewSheet(ByVal myWB As String, ByVal mySheetName As String)

On Error GoTo myhandler

    '-- add sheet and name it --
    Workbooks(myWB).Worksheets.Add
    ActiveSheet.Name = mySheetName

Exit Sub
myhandler:
            
End Sub

'---------------------------------------------------------------------------------------
' Function  : ReturnColumn
' DateTime  : 05/20/2013
' Author    : José García Herruzo
' Purpose   : Close workbooks without saving it
' Arguments :
'             myWB                       --> name
'             m_sheet                    --> worksheets name
'             str_Range                  --> Range to start to count
'---------------------------------------------------------------------------------------
Public Function ReturnColumn(ByVal myWB As String, ByVal m_sheet As String, ByVal str_Range As String) As Integer

Dim a As Variant
Dim b As Integer

On Error GoTo myhandler

If Workbooks(myWB).Worksheets(m_sheet).Range(str_Range).Value = "" Then

    ReturnColumn = 0

Else

    a = 1
    b = 0
    
    Do Until a = ""
        
        a = Workbooks(myWB).Worksheets(m_sheet).Range(str_Range).Offset(0, b).Value
        b = b + 1
        
    Loop
    
    ReturnColumn = b - 1

End If

Exit Function
myhandler:

            
End Function

'---------------------------------------------------------------------------------------
' Function  : Get_Column_Letter
' DateTime  : 05/10/2013
' Author    : José García Herruzo; Based on http://support.microsoft.com/kb/153318/es
' Purpose   : Return the column letter
' Arguments :
'             ra_myRange                 --> Objetive range
'---------------------------------------------------------------------------------------
Public Function Get_Column_Letter(ByVal ra_myRange As Range) As String
       
Dim MyColumn As String
Dim Here As String

On Error GoTo myhandler

' Get the address of the selected range in the current selection
Here = ra_myRange.Address

' Because .Address is $<columnletter>$<rownumber>, drop the first
' character and the characters after the column letter(s).
MyColumn = Mid(Here, InStr(Here, "$") + 1, InStr(2, Here, "$") - 2)

' return the column letter
Get_Column_Letter = MyColumn

Exit Function
myhandler:
            
End Function

'---------------------------------------------------------------------------------------
' Procedure : WriteFromTxT
' DateTime  : 05/29/2013
' Author    : José García Herruzo
' Purpose   : Extract info from txt to excel
' Arguments : N/A
'             wb_myWB                    --> Workbook where data must be written
'             str_WS_Name                --> Worksheet name where data must be written
'             str_SourceTxT              --> Source file name
'---------------------------------------------------------------------------------------
Public Sub WriteFromTxT(ByVal wb_myWb As Workbook, ByVal str_WS_Name As String, ByVal str_SourceTxT As String, ByVal str_EndRange As String)

    Dim oWS As Worksheet
        
    On Error GoTo myhandler
        
    If SheetExist(wb_myWb.Name, str_WS_Name) = True Then
    
        Set oWS = wb_myWb.Worksheets(str_WS_Name)
    
    Else
    
        Set oWS = wb_myWb.Worksheets.Add(After:=Worksheets(Worksheets.Count))
        oWS.Name = str_WS_Name
    
    End If
    
    '-- mola, buscar info de query tables --
     oWS.QueryTables.Add(Connection:="TEXT;" _
    & str_SourceTxT, Destination:=oWS.Range("" & str_EndRange & "")).Refresh
   
Exit Sub
myhandler:
            
End Sub

