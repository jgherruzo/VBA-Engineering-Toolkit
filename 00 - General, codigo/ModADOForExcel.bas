Attribute VB_Name = "ModADOForExcel"
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
' Module    : ModADOForExcel
' DateTime  : 05/16/2013
' Author    : José García Herruzo; Based on Emilio Sancha website example
' Purpose   : This module contents function and procedures for work with Excel as Database
' References: Microsoft Data Objects 2.X Library
'             Microsoft ADO Ext. x.x for DDL and Security
' Functions :
'               1-Connect_To_Excel
'               2-Get_Searched_Data
'               3-GetExcelWorksheet
' Procedures:
'               1-Extract_Data_From_Excel
'               2-Update_Data_Into_ExcelCell
'               3-Extract_Data_From_Excel_WithTittle
' Updates   :
'       DATE        USER    DESCRIPTION
'       07/23/2013  JGH     GetExcelWorksheet function is added
'       07/23/2013  JGH     Extract_Data_From_Excel_WithTittle function is added
'       07/23/2013  JGH     Get public each function in this module
'       08/16/2013  JGH     Log is added
'       10/08/2013  JGH     EndWb is add like argument to Extract_Data_From_Excel
'       10/09/2013  JGH     EndWb is add like argument to
'                           Extract_Data_From_Excel_WithTittle
'       19/12/2013  JGH     Ado log is changed by Error_ADOForExcel
'       17/01/2014  JGH     EHS is added
'       19/12/2013  JGH     Extract_Data_From_Excel is modified to improve error handler
'       06/11/2014  JGH     Each procedure is modified to improve error handler
'       09/03/2015  JGH     Get_Searched_Data is modified to call bol_IsFound
'----------------------------------------------------------------------------------------
'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
                        Public Error_ADOForExcel As Integer
'--------------------------- ERROR INFORMATION BOX --------------------------------------
'   KEY         FUNTION or PROCEDURE                                                    '
'   1           GetExcelWorksheet                                                       '
'   2           Get_Searched_Data                                                       '
'   3           GetExcelWorksheet                                                       '
'   4           Update_Data_Into_ExcelCell                                              '
'------------------------- END ERROR INFORMATION BOX ------------------------------------
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<

'---------------------------------------------------------------------------------------
' Procedure : Extract_Data_From_Excel
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Extract data from Excel like a Database
' Arguments :
'             mySourceWBCompletePath    --> Path+Name of the workbook which is the data
'                                           source
'             mySourceWS                --> The worksheets where the data are content
'             mySourceRange             --> The source data range
'             myEndWB                   --> Where result must be written
'             myEndSheet                --> Where result must be written
'             myEndRange                --> Where result must be written
'---------------------------------------------------------------------------------------
Public Sub Extract_Data_From_Excel(ByVal mySourceWBCompletePath As String, ByVal mySourceWS As String, _
                                    ByVal mySourceRange As String, ByVal myEndWb As Workbook, ByVal myEndSheet As String, _
                                    ByVal myEndRange As String)

Dim Connection As ADODB.Connection
Dim rs_Query As ADODB.Recordset
Dim bytColumna As Byte
Dim bol_IsContencted As Boolean
Dim bol_IsFound As Boolean
On Error GoTo myhandler

bol_IsContencted = False
bol_IsFound = False

Set Connection = Connect_To_Excel(mySourceWBCompletePath)

bol_IsContencted = True

Set rs_Query = Get_Searched_Data(Connection, mySourceWS, mySourceRange)

If Error_ADOForExcel = 3 Then

    bol_IsFound = False
    If bol_IsContencted = True Then
    
        Connection.Close
    
    End If
    Set Connection = Nothing
    Set rs_Query = Nothing
    Exit Sub
    
Else

    bol_IsFound = True
    
End If

' -- Data are written --
If Not (rs_Query.EOF And rs_Query.BOF) Then

    myEndWb.Worksheets(myEndSheet).Range(myEndRange).CopyFromRecordset rs_Query
    
End If

' -- Close the every object --
If Not rs_Query Is Nothing Then

    rs_Query.Close
    Set rs_Query = Nothing
    
End If

Connection.Close
Set Connection = Nothing
    
Exit Sub
myhandler:
    If bol_IsFound = True And Error_ADOForExcel <> 3 Then
    
        rs_Query.Close
        
    End If
    
    If bol_IsContencted = True Then
    
        Connection.Close
    
    End If
    Set Connection = Nothing
    Set rs_Query = Nothing
    
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Extract_Data_From_Excel_WithTittle
' DateTime  : 07/23/2013
' Author    : José García Herruzo
' Purpose   : Extract data from Excel like a Database
' Arguments :
'             mySourceWBCompletePath    --> Path+Name of the workbook which is the data
'                                           source
'             mySourceWS                --> The worksheets where the data are content
'             mySourceRange             --> The source data range
'             myEndWB                   --> Where result must be written
'             myEndSheet                --> Where result must be written
'             myEndRange                --> Where result must be written
'---------------------------------------------------------------------------------------
Public Sub Extract_Data_From_Excel_WithTittle(ByVal mySourceWBCompletePath As String, ByVal mySourceWS As String, _
                                    ByVal mySourceRange As String, ByVal myEndWb As Workbook, ByVal myEndSheet As String, _
                                    ByVal myEndRange As String)

Dim Connection As ADODB.Connection
Dim rs_Query As ADODB.Recordset
Dim bytColumna As Byte
Dim i As Integer
Dim x As Integer
Dim bol_IsContencted As Boolean
Dim bol_IsFound As Boolean

On Error GoTo myhandler

bol_IsContencted = False
bol_IsFound = False

Set Connection = Connect_To_Excel(mySourceWBCompletePath)
bol_IsContencted = True

Set rs_Query = Get_Searched_Data(Connection, mySourceWS, mySourceRange)
bol_IsFound = True

' -- Data are written --
If Not (rs_Query.EOF And rs_Query.BOF) Then

    myEndWb.Worksheets(myEndSheet).Range(myEndRange).CopyFromRecordset rs_Query
    x = rs_Query.Fields.Count
    
    For i = 0 To x - 1
    
        myEndWb.Worksheets(myEndSheet).Range(myEndRange).Offset(-1, i).Value = rs_Query.Fields(i).Name
    
    Next

End If

' -- Close the every object --
If Not rs_Query Is Nothing Then

    rs_Query.Close
    Set rs_Query = Nothing
    
End If

Connection.Close
Set Connection = Nothing
    
Exit Sub
myhandler:
    If bol_IsFound = True And Error_ADOForExcel <> 3 Then
    
        rs_Query.Close
        
    End If
    
    If bol_IsContencted = True Then
    
        Connection.Close
    
    End If
    Set Connection = Nothing
    Set rs_Query = Nothing
                
End Sub
'---------------------------------------------------------------------------------------
' Procedure : Update_Data_Into_ExcelCell
' DateTime  : 05/16/2013
' Author    : José García Herruzo
' Purpose   : Update in one value; For more cell, or add new values use other function
' Arguments :
'             mySourceWBCompletePath    --> Path+Name of the workbook which is the data
'                                           destiny
'             mySourceWS                --> The worksheets where the data will be written
'             myData                    --> Data to be written
'             str_TableField            --> The tittle of the column were data must be
'                                           written
'---------------------------------------------------------------------------------------
Public Function Update_Data_Into_ExcelCell(ByVal mySourceWBCompletePath As String, ByVal mySourceWS As String, _
                                    ByVal myData As String, ByVal str_TableField As String)

Dim Connection As ADODB.Connection
Dim rs_Query As ADODB.Recordset
Dim bol_IsFound As Boolean
Dim bol_IsConnected As Boolean

On Error GoTo Write_Data_Into_Excel_Error

bol_IsConnected = False
bol_IsFound = False

Set Connection = Connect_To_Excel(mySourceWBCompletePath)
bol_IsConnected = True


Set rs_Query = New ADODB.Recordset
    
    Set rs_Query = New ADODB.Recordset
    
    With rs_Query
        .CursorLocation = adUseClient
        .Open "Select * from [" & mySourceWS & "$]", Connection, adOpenStatic, adLockOptimistic
        bol_IsFound = True
        .Update str_TableField, myData
        .Close
    End With

Connection.Close
Set Connection = Nothing
Set rs_Query = Nothing
    
Exit Function

Write_Data_Into_Excel_Error:
    If bol_IsFound = True And Error_ADOForExcel <> 3 Then
    
        rs_Query.Close
        
    End If
    
    If bol_IsConnected = True Then
    
        Connection.Close
    
    End If
    Set Connection = Nothing
    Set rs_Query = Nothing
    Error_ADOForExcel = 4

End Function
'---------------------------------------------------------------------------------------
' Function  : Connect_To_Excel
' DateTime  : 05/07/2013
' Author    : José García Herruzo; Ayuda http://www.excelpatas.com/2010/05/escribir-datos-en-un-libro-cerrado-por.html
' Purpose   : Connect to excel as database; Only .xls files
' Arguments :
'             str_Workbook              --> Path+Name of the workbook which is the data
'                                           source
'---------------------------------------------------------------------------------------
Public Function Connect_To_Excel(ByVal str_Workbook As String) As ADODB.Connection
    
Dim str_Connection As String
Dim ado_Connection As ADODB.Connection
Dim str_Provider As String
Dim str_Version As String
Dim bol_IsConnected As Boolean

On Error GoTo myhandler

    bol_IsConnected = False
    
    'Tipo de conexión según el archivo en el que se use la función
    If VBA.Val(VBA.Mid(Application.Version, 1, VBA.InStr(1, Application.Version, ".") - 1)) >= 12 Then
    
        str_Provider = "Microsoft.ACE.OLEDB.12.0" 'Excel 2007
        
    Else
    
        str_Provider = "Microsoft.Jet.OLEDB.4.0" 'Excel 2003
        
    End If
    
    str_Version = VBA.Right(str_Workbook, VBA.Len(str_Workbook) - VBA.InStr(str_Workbook, "."))
    
    'Versión del archivo destino
    If str_Version Like "xls?" Then
    
        str_Version = "Excel 12.0" 'Excel 2003
        
    ElseIf str_Version Like "xls" Then 'Excel 2007
    
        str_Version = "Excel 8.0"
        
    Else
    
        'No se está trabajando con una archivo de Excel válido
        str_Version = "Excel 8.0"
        
    End If
    
    str_Connection = "Provider=" & str_Provider & ";" & _
              "Data Source=" & str_Workbook & ";" & _
              "Extended Properties=""" & str_Version & ";HDR=Yes"";"
              
Set ado_Connection = CreateObject("ADODB.Connection")

ado_Connection.Open str_Connection
bol_IsConnected = True

Set Connect_To_Excel = ado_Connection

'ado_Connection.Close
Set ado_Connection = Nothing
    
Exit Function
myhandler:
    
    If bol_IsConnected = True Then
    
        ado_Connection.Close
    
    End If

    Set ado_Connection = Nothing
    Error_ADOForExcel = 2
            
End Function

'---------------------------------------------------------------------------------------
' Function  : Get_Searched_Data
' DateTime  : 05/07/2013
' Author    : José García Herruzo
' Purpose   : Made a query
' Arguments :
'             ado_Connection            --> Connection to Excel database
'             mySourceWS                --> The worksheets where the data are content
'             mySourceRange             --> The source data range
'---------------------------------------------------------------------------------------
Public Function Get_Searched_Data(ByVal ado_Connection As ADODB.Connection, ByVal mySourceWS As String, _
                                ByVal mySourceRange As String) As ADODB.Recordset

Dim str_Query As String
Dim str_Source As String
Dim rs_Query As ADODB.Recordset
Dim bol_IsOpened As Boolean
Dim bol_IsFound As Boolean

On Error GoTo myhandler

bol_IsOpened = False

str_Source = "[" & mySourceWS & "$" & mySourceRange & "]"

str_Query = _
        "SELECT " & vbCrLf & _
        "     * " & vbCrLf & _
        "FROM " & vbCrLf & _
        "     " & str_Source

Set rs_Query = CreateObject("ADODB.Recordset")

rs_Query.Open str_Query, ado_Connection, adOpenDynamic, adLockOptimistic, adCmdText
bol_IsOpened = True

Set Get_Searched_Data = rs_Query

Set rs_Query = Nothing

Exit Function
myhandler:
    If bol_IsFound = True And Error_ADOForExcel <> 3 Then
    
        rs_Query.Close
        
    End If
    Set rs_Query = Nothing
    Error_ADOForExcel = 3
            
End Function
'---------------------------------------------------------------------------------------
' Function  : GetExcelWorksheet
' DateTime  : 07/23/2013
' Author    : José García Herruzo
' Purpose   : Return a string with each worksheet in the selected workbook
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function GetExcelWorksheet(ByVal strmyFile As String) As String

Dim myworkbook As ADOX.Catalog
Dim myWorksheet As ADOX.Table
Dim Tmp As Variant

On Error GoTo myhandler

Set myworkbook = New ADOX.Catalog

myworkbook.ActiveConnection = _
"Provider=MSDASQL.1;Data Source=Excel Files;Initial Catalog=" & strmyFile

For Each myWorksheet In myworkbook.Tables

    Tmp = Application.Substitute(myWorksheet.Name, "'", "")
    If Right(Tmp, 1) = "$" Then
    
        GetExcelWorksheet = GetExcelWorksheet & "\" & Left(Tmp, Len(Tmp) - 1)
    
    End If

Next myWorksheet

Set myworkbook = Nothing

Exit Function
myhandler:
    Set myworkbook = Nothing
    Error_ADOForExcel = 1
            
End Function

