<div align="center">

## Registry Manipulation Module


</div>

### Description

Included in this module are several functions to help you in handling with the registry. I have written and tested these as best I can. You may use them at your own risk. I frequently use SetAppKeyValue and GetAppKeyValue without any problems. If you wish to alter or test any of the functions, I strongly advise you to backup your registry first (Start > Run "regedit").

The included functions are:

CloseKey,

CreateKey,

DeleteKey,

DeleteKeyStruct,

EnumKey,

GetAppKeyValue,

GetKeyValue,

OpenKey,

SetAppKeyValue, and

SetKeyValue.

This code covers:

Constants,

Functions,

GoTo,

On Error,

Recursion,

RegCloseKey API,

RegCreateKeyEx API,

RegDeleteKey API,

RegEnumKeyEx API,

RegQueryValueEx API,

RegOpenKeyEx API,

RegSetValueEx API, and

Types.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2002-07-09 09:03:12
**By**             |[Michael Drotar](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/michael-drotar.md)
**Level**          |Advanced
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[Libraries](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/libraries__1-49.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Registry\_M103871792002\.zip](https://github.com/Planet-Source-Code/michael-drotar-registry-manipulation-module__1-36731/archive/master.zip)

### API Declarations

```
Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
```





