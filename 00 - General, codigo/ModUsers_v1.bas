Attribute VB_Name = "ModUsers_v1"
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
' Module    : ModUsers_v1
' DateTime  : 05/08/2013
' Author    : José García Herruzo
' Purpose   : This module contents function and procedures to work with users
' References: N/A
' Functions :
'               1-GetUser
'               2-LastUser
'               2-GetFullUserName
'               2-fGetDCName
'               2-fStrFromPtrW
' Procedures: N/A
' Updates   :
'       DATE        USER    DESCRIPTION
'       29/12/2014  JGH     GetFullUserName and required functions are added
'----------------------------------------------------------------------------------------

'<<<< Api GetUserName declaration >>>>>
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" ( _
    ByVal lpBuffer As String, nSize As Long) As Long
    
Private Declare Sub sapiCopyMem Lib "kernel32" Alias "RtlMoveMemory" _
    (Destination As Any, Source As Any, ByVal Length As Long)
    
Private Declare Function apiNetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" _
    (ByVal buffer As Long) As Long

Private Declare Function apiNetUserGetInfo Lib "netapi32.dll" Alias "NetUserGetInfo" _
    (servername As Any, username As Any, ByVal level As Long, bufptr As Long) As Long

Private Declare Function apiNetGetDCName Lib "netapi32.dll" Alias "NetGetDCName" _
    (ByVal servername As Long, ByVal DomainName As Long, bufptr As Long) As Long
    
Private Declare Function apilstrlenW Lib "kernel32" Alias "lstrlenW" _
    (ByVal lpString As Long) As Long
    
'<<<< Tipo de datos para obtener información >>>>>
Private Type USER_INFO_2
    usri2_name As Long
    usri2_password  As Long  ' Null
    usri2_password_age  As Long
    usri2_priv  As Long
    usri2_home_dir  As Long
    usri2_comment  As Long
    usri2_flags  As Long
    usri2_script_path  As Long
    usri2_auth_flags  As Long
    usri2_full_name As Long
    usri2_usr_comment  As Long
    usri2_parms  As Long
    usri2_workstations  As Long
    usri2_last_logon  As Long
    usri2_last_logoff  As Long
    usri2_acct_expires  As Long
    usri2_max_storage  As Long
    usri2_units_per_week  As Long
    usri2_logon_hours  As Long
    usri2_bad_pw_count  As Long
    usri2_num_logons  As Long
    usri2_logon_server  As Long
    usri2_country_code  As Long
    usri2_code_page  As Long
End Type
 
Private Const ERROR_SUCCESS = 0&
'---------------------------------------------------------------------------------------
' Function  : GetUser
' DateTime  : 08/05/2013
' Author    : José García Herruzo
' Purpose   : Return a string with the username
' Arguments : N/A
'---------------------------------------------------------------------------------------
Public Function GetUser() As String
       
    Dim Nombre As String, ret As Long
       
    ' Buffer
    Nombre = Space$(250)
       
    ' Tamaño
    ret = Len(Nombre)
       
    If GetUserName(Nombre, ret) = 0 Then
        GetUser = vbNullString
    Else
        ' Extrae solo los caracteres
        GetUser = Left$(Nombre, ret - 1)
    End If
       
End Function

'---------------------------------------------------------------------------------------
' Function  : LastUser
' DateTime  : 08/05/2013
' Author    : Code by Helen from http://www.visualbasicforum.com/index.php?s=
'             Modify by Emilio Sancha
'             Modify by José García Herruzo
' Purpose   : This routine gets the Username of the File In Use
' Arguments :
'             strRuta                   --> Path+name
'---------------------------------------------------------------------------------------
Public Function LastUser(strRuta As String) As String
'// Credit goes to Helen for code & Mark for the idea
'// Insomniac for xl97 inStrRev
'// Amendment 25th June 2004 by IFM
'// : Name changes will show old setting
'// : you need to get the Len of the Name stored just before
'// : the double Padded Nullstrings

Dim strArchivo As String, _
    strFlag1 As String, _
    strFlag2 As String, _
    i As Integer, _
    j As Integer, _
    lngArchivo As Long, _
    bytTamañoNombre As Byte

strFlag1 = Chr(0) & Chr(0)
strFlag2 = Chr(32) & Chr(32)

lngArchivo = FreeFile
Open strRuta For Binary As #lngArchivo
' creo una cadena vacía del mismo tamaño del libro
strArchivo = Space(LOF(lngArchivo))
' inserto en ella el contenido del libro
Get 1, , strArchivo
Close #lngArchivo

' busco en ella la primera aparición de dos espacios seguidos
j = InStr(1, strArchivo, strFlag2)

' según la versión de VBA, busco de un modo u otro la aparición de dos caracteres nulos seguidos
#If Not VBA6 Then
    '// Xl97
    For i = j - 1 To 1 Step -1
        If Mid(strArchivo, i, 1) = Chr(0) Then Exit For
    Next
    i = i + 1
#Else
    '// Xl2000+
    i = InStrRev(strArchivo, strFlag1, j) + Len(strFlag1)
#End If

' extraigo el nombre del usuario
bytTamañoNombre = Asc(Mid(strArchivo, i - 3, 1))
LastUser = Mid(strArchivo, i, bytTamañoNombre)

End Function
'---------------------------------------------------------------------------------------
' Function  : GetFullUserName
' DateTime  : 29/12/2014
' Author    : Code by Zwarrior from http://www.chw.net/foro/lenguajes-programacion/
'                253925-vba-cargar-nombre-usuario-completo-resuelto.html
'             Modify by José García Herruzo
' Purpose   : This routine gets the full User name
' Arguments :
'             strUserName               --> Short user name
'---------------------------------------------------------------------------------------
Public Function GetFullUserName(Optional strUserName As String) As String
' Obtener el Nombre completo de usuario usando un UserID
' Para Windows NT/2000/XP/Vista
' Omitir el parámetro strUserName
' Hará que se obtenga el nombre del usuario con sesión iniciada
On Error GoTo ErrHandler

Dim pBuf As Long
Dim dwRec As Long
Dim pTmp As USER_INFO_2
Dim abytPDCName() As Byte
Dim abytUserName() As Byte
Dim lngRet As Long
Dim i As Long
 
    ' Unicode
    abytPDCName = fGetDCName() & vbNullChar
    
    If (Len(strUserName) = 0) Then strUserName = GetUser()
    abytUserName = strUserName & vbNullChar
 
    lngRet = apiNetUserGetInfo( _
                            abytPDCName(0), _
                            abytUserName(0), _
                            2, _
                            pBuf)
    If (lngRet = ERROR_SUCCESS) Then
        Call sapiCopyMem(pTmp, ByVal pBuf, Len(pTmp))
        GetFullUserName = fStrFromPtrW(pTmp.usri2_full_name)
    End If
 
    Call apiNetAPIBufferFree(pBuf)
ExitHere:
    Exit Function
ErrHandler:
    GetFullUserName = vbNullString
    Resume ExitHere
End Function
'---------------------------------------------------------------------------------------
' Function  : fGetDCName
' DateTime  : 29/12/2014
' Author    : Code by Zwarrior from http://www.chw.net/foro/lenguajes-programacion/
'                253925-vba-cargar-nombre-usuario-completo-resuelto.html
'             Modify by José García Herruzo
' Purpose   : This routine gets the full User name
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Function fGetDCName() As String

Dim pTmp As Long
Dim lngRet As Long
Dim abytBuf() As Byte
 
lngRet = apiNetGetDCName(0, 0, pTmp)

If lngRet = ERROR_SUCCESS Then

    fGetDCName = fStrFromPtrW(pTmp)
    
End If

Call apiNetAPIBufferFree(pTmp)

End Function
'---------------------------------------------------------------------------------------
' Function  : fStrFromPtrW
' DateTime  : 29/12/2014
' Author    : Code by Zwarrior from http://www.chw.net/foro/lenguajes-programacion/
'                253925-vba-cargar-nombre-usuario-completo-resuelto.html
'             Modify by José García Herruzo
' Purpose   : Funciones para llamado a las librerías y obtención de datos
' Arguments : N/A
'---------------------------------------------------------------------------------------
Private Function fStrFromPtrW(pBuf As Long) As String

Dim lngLen As Long
Dim abytBuf() As Byte
 
    ' Get the length of the string at the memory location
    lngLen = apilstrlenW(pBuf) * 2
    ' if it's not a ZLS
    If lngLen Then
    
        ReDim abytBuf(lngLen)
        ' then copy the memory contents
        ' into a temp buffer
        Call sapiCopyMem(abytBuf(0), ByVal pBuf, lngLen)
        ' return the buffer
        fStrFromPtrW = abytBuf
        
    End If
    
End Function
