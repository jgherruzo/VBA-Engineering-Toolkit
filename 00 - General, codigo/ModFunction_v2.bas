Attribute VB_Name = "ModFunction_v2"
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
' Module    : ModFunction_v2
' DateTime  : 05/29/2013
' Author    : http://www.mrexcel.com/forum/excel-questions/535773-list-softwares- _
                installed-excel-using-visual-basic-applications.html
'               <by Bill James>; Modified by José García Herruzo
' Purpose   : This module contents function not classify. It is applied to obtained
'               the installed software
' References: N/A
' Functions :
'               1-GetProbedID
'               2-GetAddRemove
'               3-GetDTFileName
'               4-BubbleSort
'               5-IsSoftwareInstalled
' Procedures:
'               1-GetInstalledSoftWare
' Updates   :
'       DATE        USER    DESCRIPTION
'       N/A
'----------------------------------------------------------------------------------------
Private sFileName

'---------------------------------------------------------------------------------------
' Procedure : CheckingLimit
' DateTime  : 05/29/2013
' Author    : See mod info
' Purpose   : Controls the procedure calling to the rest of the function
' Arguments :
'             StrComputer               --> Value to check
'---------------------------------------------------------------------------------------
Public Sub GetInstalledSoftWare(Optional StrComputer As String = "")

    Dim sCompName As String
    Dim sTitle As String
    Dim s As String

    If StrComputer = "" Then
    
        StrComputer = "."
        
    End If
    
    '-- Get PC ID --
    sCompName = GetProbedID(StrComputer)
    
    '-- Check scompname is not empty--
    If Len(sCompName) > 0 Then
    
        '-- Make txt file name --
        sFileName = "C:\" & sCompName & "_" & GetDTFileName() & "_Software.txt"
        
        '-- Search on register each installed software --
        s = GetAddRemove(StrComputer)
        
        '-- Write this string into a txt file--
        Call WriteTxTFile(s, sFileName)
        
        Do
        
        DoEvents
        Loop Until Len(Dir(sFileName)) <> 0
        
        '--delete the worksheet--
        If SheetExist(ThisWorkbook.Name, "Installed_Software") = True Then
            
            Application.DisplayAlerts = False
            ThisWorkbook.Worksheets("Installed_Software").Delete
            Application.DisplayAlerts = True
            
        End If
        
        '-- Write the info into an excel file --
        Call WriteFromTxT(ThisWorkbook, "Installed_Software", sFileName, "A1")
        
        '--delete the developed txt file--
        Kill sFileName
        
    End If
    
End Sub

'---------------------------------------------------------------------------------------
' Function  : GetAddRemove
' DateTime  : 05/29/2013
' Author    : See mod info;Function credit to Torgeir Bakken
' Purpose   : Search on register each installed software
' Arguments :
'             sComp                    --> Computer ID
'---------------------------------------------------------------------------------------
Private Function GetAddRemove(ByVal sComp As String) As String

  Dim cnt, oReg, sBaseKey, iRC, aSubKeys
  Dim sCompName As String
  Const HKLM = &H80000002  'HKEY_LOCAL_MACHINE
  Dim sKey, sValue, sTmp, sVersion, sDateValue, sYr, sMth, sDay
  
  Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
              sComp & "/root/default:StdRegProv")
  sBaseKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
  iRC = oReg.EnumKey(HKLM, sBaseKey, aSubKeys)
  
  For Each sKey In aSubKeys
  
    iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, "DisplayName", sValue)
    
    If iRC <> 0 Then
    
      oReg.GetStringValue HKLM, sBaseKey & sKey, "QuietDisplayName", sValue
      
    End If
    
    If sValue <> "" Then
    
      iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, _
                                "DisplayVersion", sVersion)
      If sVersion <> "" Then
      
        sValue = sValue & vbTab & "Ver: " & sVersion
        
      Else
      
        sValue = sValue & vbTab
        
      End If
      
      iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, _
                                "InstallDate", sDateValue)
      If sDateValue <> "" Then
      
        sYr = Left(sDateValue, 4)
        sMth = Mid(sDateValue, 5, 2)
        sDay = Right(sDateValue, 2)
        
        'some Registry entries have improper date format
        On Error Resume Next
        sDateValue = DateSerial(sYr, sMth, sDay)
        On Error GoTo 0
        
        If sDateValue <> "" Then
        
          sValue = sValue & vbTab & "Installed: " & sDateValue
          
        End If
        
      End If
      
      sTmp = sTmp & sValue & vbCrLf
      
    cnt = cnt + 1
    End If
    
  Next
  
  sTmp = BubbleSort(sTmp)
  GetAddRemove = "INSTALLED SOFTWARE (" & cnt & ") - " & sCompName & _
                 " - " & Now() & vbCrLf & vbCrLf & sTmp
End Function

'---------------------------------------------------------------------------------------
' Function  : BubbleSort
' DateTime  : 05/29/2013
' Author    : See mod info
' Purpose   : Returns date time value
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Function BubbleSort(sTmp) As String

  'cheapo bubble sort
  Dim aTmp, i, j, temp
  aTmp = Split(sTmp, vbCrLf)
  For i = UBound(aTmp) - 1 To 0 Step -1
    For j = 0 To i - 1
      If LCase(aTmp(j)) > LCase(aTmp(j + 1)) Then
        temp = aTmp(j + 1)
        aTmp(j + 1) = aTmp(j)
        aTmp(j) = temp
      End If
    Next
  Next
  BubbleSort = Join(aTmp, vbCrLf)
  
End Function

'---------------------------------------------------------------------------------------
' Function  : GetProbedID
' DateTime  : 05/29/2013
' Author    : See mod info
' Purpose   : Returns Pc ID
' Arguments :
'             sComp                    --> String with computer name
'---------------------------------------------------------------------------------------
Private Function GetProbedID(sComp) As String

  Dim objWMIService, colItems, objItem
  On Error Resume Next
  Set objWMIService = GetObject("winmgmts:\\" & sComp & "\root\cimv2")
  
  If Err.Number = 462 Then MsgBox Err.Description
  Set colItems = objWMIService.ExecQuery("Select SystemName from " & _
                                         "Win32_NetworkAdapter", , 48)
  For Each objItem In colItems
    GetProbedID = objItem.SystemName
  Next
    
End Function

'---------------------------------------------------------------------------------------
' Function  : GetDTFileName
' DateTime  : 05/29/2013
' Author    : See mod info
' Purpose   : Returns date time value
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Function GetDTFileName() As String

  Dim sNow, sMth, sDay, sYr, sHr, sMin, sSec
  sNow = Now
  sMth = Right("0" & Month(sNow), 2)
  sDay = Right("0" & Day(sNow), 2)
  sYr = Right("00" & Year(sNow), 4)
  sHr = Right("0" & Hour(sNow), 2)
  sMin = Right("0" & Minute(sNow), 2)
  sSec = Right("0" & Second(sNow), 2)
  GetDTFileName = sMth & sDay & sYr & "_" & sHr & sMin & sSec
  
End Function

'---------------------------------------------------------------------------------------
' Function  : IsSoftwareInstalled
' DateTime  : 05/29/2013
' Author    : José García Herruzo
' Purpose   : Search if Aspen Plus is installed
' Arguments : N/A
'             str_wsName                --> Worksheet where you want to search
'             str_Range                 --> column where software names are written
'---------------------------------------------------------------------------------------
Public Function IsSoftwareInstalled(ByVal ws_myWs As Worksheet, ByVal str_Range As String, ByVal str_mySoftware As String) As Boolean

Dim LastRow As Integer
Dim i As Integer

LastRow = ws_myWs.Range(str_Range).Offset(16000, 0).End(xlUp).Row

For i = 0 To LastRow

    If InStr(ws_myWs.Range(str_Range).Offset(i, 0).Value, str_mySoftware) <> 0 Then
    
        IsSoftwareInstalled = True
        Exit Function
    
    End If
    
Next i

IsSoftwareInstalled = False

End Function
