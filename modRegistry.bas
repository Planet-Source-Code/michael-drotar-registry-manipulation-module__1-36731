Attribute VB_Name = "modRegistry"
Option Explicit

Const ERROR_SUCCESS = 0&
Const ERROR_NO_MORE_ITEMS = 259&

Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H80000005
Const HKEY_DYN_DATA = &H80000006

Const REG_NONE = 0
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_DWORD_LITTLE_ENDIAN = 4
Const REG_DWORD_BIG_ENDIAN = 5
Const REG_LINK = 6
Const REG_MULTI_SZ = 7

Const STANDARD_RIGHTS_ALL = &H1F0000
Const SYNCHRONIZE = &H100000
Const READ_CONTROL = &H20000

Const STANDARD_RIGHTS_READ = READ_CONTROL
Const STANDARD_RIGHTS_WRITE = READ_CONTROL

'dwOptions
Const REG_OPTION_BACKUP_RESTORE = 4     ' open for backup or restore (WinNT)
Const REG_OPTION_VOLATILE = 1           ' Key is not preserved (WinNT)
Const REG_OPTION_NON_VOLATILE = 0       ' Key is preserved (Default)

'samDesired
Const KEY_CREATE_LINK = &H20
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or _
                        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
                     And (Not SYNCHRONIZE))
Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or _
                        KEY_CREATE_SUB_KEY) _
                     And (Not SYNCHRONIZE))
Const KEY_EXECUTE = KEY_READ
Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or _
                        KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or _
                        KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or _
                        KEY_CREATE_LINK) _
                     And (Not SYNCHRONIZE))
                     
'lpdwDisposition
Const REG_CREATED_NEW_KEY As Long = &H1
Const REG_OPENED_EXISTING_KEY As Long = &H2

'lpSecurityAttributes
Public Type SECURITY_ATTRIBUTES
   nLength As Long
   lpSecurityDescriptor As Long
   bInheritHandle As Long
End Type
Private SECURE As SECURITY_ATTRIBUTES


'====================================================================================
'====================================================================================

Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" _
            (ByVal hKey As Long, ByVal lpSubKey As String, _
                ByVal Reserved As Long, ByVal lpClass As String, _
                ByVal dwOptions As Long, ByVal samDesired As Long, _
                lpSecurityAttributes As SECURITY_ATTRIBUTES, _
                phkResult As Long, lpdwDisposition As Long) _
                                                                            As Long
                                                                           
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
            (ByVal hKey As Long, ByVal lpSubKey As String, _
                ByVal ulOptions As Long, ByVal samDesired As Long, _
                phkResult As Long) _
                                                                            As Long
                                                                            
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" _
            (ByVal hKey As Long, ByVal dwIndex As Long, _
                ByVal lpName As String, lpcbName As Long, _
                lpReserved As Long, ByVal lpClass As String, _
                lpcbClass As Long, lpftLastWriteTime As Any) _
                                                                            As Long

Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
            (ByVal hKey As Long, ByVal lpSubKey As String) _
                                                                            As Long
Declare Function RegCloseKey Lib "advapi32.dll" _
            (ByVal hKey As Long) _
                                                                            As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
            (ByVal hKey As Long, ByVal lpValueName As String, _
                ByVal Reserved As Long, ByVal dwType As Long, _
                ByVal lpData As String, ByVal cbData As Long) _
                                                                            As Long
                                                                            
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
            (ByVal hKey As Long, ByVal lpValueName As String, _
                ByVal lpReserved As Long, lpType As Long, _
                lpData As Any, lpcbData As Long) _
                                                                            As Long

'====================================================================================
'====================================================================================


Public Function OpenKey _
            (ByVal lpPath As String, ByRef phkResult As Long) _
                                                                           As Long

   Dim ret As Long, lpdwDisposition As Long
   ret = RegOpenKeyEx(HKEY_CURRENT_USER, "Software\" & lpPath, 0, KEY_ALL_ACCESS, _
                        phkResult)
   If phkResult = 0 Then
        ret = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\" & lpPath, _
                        0, "REG_DWORD", REG_OPTION_NON_VOLATILE, _
                        KEY_ALL_ACCESS, SECURE, phkResult, lpdwDisposition)
    End If
    
    OpenKey = ret
End Function

Public Function CloseKey _
            (ByVal phkResult As Long) _
                                                                           As Long
   CloseKey = RegCloseKey(phkResult)
End Function

Public Function CreateKey _
            (ByVal lpPath As String) _
                                                                           As Long
    Dim ret As Long, phkResult As Long
    ret = OpenKey(lpPath, phkResult)
    If ret = ERROR_SUCCESS Then ret = CloseKey(phkResult)
    CreateKey = ret
End Function

Public Function SetAppKeyValue _
            (ByVal lpKeyName As Variant, lpKeyValue As Variant, _
                Optional dwType = REG_SZ) _
                                                                            As Long
    SetKeyValue App.Title, lpKeyName, lpKeyValue, dwType
End Function

Public Function SetKeyValue _
            (ByVal lpPath As String, lpKeyName As Variant, lpKeyValue As Variant, _
                Optional dwType = REG_SZ) _
                                                                            As Long
            
    Dim cbData As Long, phkResult As Long
    
    Select Case dwType
        Case REG_SZ
            cbData = Len(lpKeyValue)
        Case Else
            cbData = 0
    End Select
            
    OpenKey lpPath, phkResult
    SetKeyValue = RegSetValueEx(phkResult, lpKeyName, 0, dwType, lpKeyValue, cbData)
    CloseKey phkResult
End Function

Public Function GetAppKeyValue _
            (ByVal lpKeyName As Variant, Optional defValue As Variant = 0) _
                                                                            As Variant
    GetAppKeyValue = GetKeyValue(App.Title, lpKeyName, defValue)
End Function
            
Public Function GetKeyValue _
            (ByVal lpPath As String, lpKeyName As Variant, Optional defValue As Variant = 0) _
                                                                            As Variant
    Dim phkResult As Long, lResult As Long, lValueType As Long, _
            strBuf As String, lDataBufSize As Long, ret As Variant
    
    OpenKey lpPath, phkResult
    lResult = RegQueryValueEx(phkResult, lpKeyName, 0, lValueType, ByVal 0, lDataBufSize)
    If lResult = 0 Then
        If lValueType = REG_SZ Then
            strBuf = String(lDataBufSize, Chr$(0))
            lResult = RegQueryValueEx(phkResult, lpKeyName, 0, 0, ByVal strBuf, _
                                        lDataBufSize)
            If lResult = 0 And lDataBufSize > 0 Then
                ret = Left$(strBuf, InStr(1, strBuf, Chr$(0)) - 1)
                If ret = "" Then ret = defValue
                GetKeyValue = ret
            Else
                GetKeyValue = defValue
            End If
        ElseIf lValueType = REG_BINARY Then
            Dim strdata As Integer
            lResult = RegQueryValueEx(phkResult, lpKeyName, 0, 0, strdata, lDataBufSize)
            If lResult = 0 Then
                GetKeyValue = strdata
            Else
                GetKeyValue = defValue
            End If
        End If
    Else
        GetKeyValue = defValue
    End If
    CloseKey phkResult
End Function

Public Function EnumKey _
            (ByVal lpPath As String, ByVal cnt As Long, ByRef sName As String, _
                ByRef sLen As Long) _
                                                                            As Long
    Dim ret As Long, phkResult As Long
    Const BUFFER_SIZE = 255
    
    ret = OpenKey(lpPath, phkResult)
    If ret = ERROR_SUCCESS Then ret = RegEnumKeyEx(phkResult, cnt, sName, sLen, _
            ByVal 0&, vbNullString, ByVal 0&, ByVal 0&)
    CloseKey phkResult
    EnumKey = ret
End Function

Public Function DeleteKeyStruct _
            (ByVal lpPath As String, ByRef phkResult As Long) _
                                                                            As Long
    On Error GoTo ErrHandler
    
    Dim ret As Long
    Dim cnt As Long, sName As String, sLen As Long
    Const BUFFER_SIZE = 255
    
    sLen = BUFFER_SIZE
    sName = Space(BUFFER_SIZE)
    
    cnt = 0     'Since we're deleting the keys as they're found, don't increment the cnt
    While EnumKey(lpPath, cnt, sName, sLen) <> ERROR_NO_MORE_ITEMS
        sName = Left$(sName, sLen)
        ret = DeleteKeyStruct(lpPath & "\" & sName, phkResult)
        If ret <> ERROR_SUCCESS Then GoTo ErrHandler
        
        sName = Space(BUFFER_SIZE)
        sLen = BUFFER_SIZE
    Wend
    
ErrHandler:
    DeleteKeyStruct = RegDeleteKey(HKEY_CURRENT_USER, "software\" & lpPath)
End Function

Public Function DeleteKey _
            (ByVal lpPath As String) _
                                                                           As Long
    On Error GoTo ErrHandler
    
    Dim ret As Long, phkResult As Long

    ret = OpenKey(lpPath, phkResult)
    If ret = ERROR_SUCCESS Then ret = DeleteKeyStruct(lpPath, phkResult)
        

ErrHandler:
    CloseKey phkResult
    DeleteKey = ret
End Function
