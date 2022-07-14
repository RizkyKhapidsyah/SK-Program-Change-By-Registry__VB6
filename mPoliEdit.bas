Attribute VB_Name = "mPoliEdit"
Public Declare Function SHRestartSystem Lib "Shell32" Alias "#59" (ByVal hOwner As Long, ByVal sPrompt As String, ByVal uFlags As Long) As Long
Public Const Restart_Logoff = &H0
Public Const Restart_ShutDown = &H1
Public Const Restart_Reboot = &H2
Public Const Restart_Force = &H4

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Const REG_DWORD As Long = 4
Private Const KEY_ALL_ACCESS As Long = &H3F
Private Const HKEY_CURRENT_USER As Long = &H80000001
Private Const ERROR_SUCCESS As Long = 0

Private Function regValue_Exist(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String) As Boolean
  Dim lKeyHandle As Long
  Dim lRet As Long
  Dim lDataType As Long
  Dim lBufferSize As Long
  lKeyHandle = 0
  lRet = RegOpenKey(hKey, sRegKeyPath, lKeyHandle)
  If lKeyHandle <> 0 Then lRet = RegQueryValueEx(lKeyHandle, sRegSubKey, 0&, lDataType, ByVal 0&, lBufferSize)
  If lRet = ERROR_SUCCESS Then
     regValue_Exist = True
     lRet = RegCloseKey(lKeyHandle)
  Else
     regValue_Exist = False
  End If
End Function

Private Sub regDelete_Sub_Key(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String)
  Dim lKeyHandle As Long
  Dim lRet As Long
  If regValue_Exist(hKey, sRegKeyPath, sRegSubKey) Then
     lRet = RegOpenKey(hKey, sRegKeyPath, lKeyHandle)
     lRet = RegDeleteValue(lKeyHandle, sRegSubKey)
     lRet = RegCloseKey(lKeyHandle)
  End If
End Sub

Private Sub regCreate_LongValue(ByVal hKey As Long, ByVal sRegKeyPath As String, ByVal sRegSubKey As String, lKeyValue As Long)
   Dim lKeyHandle As Long
   Dim lRet As Long
   Dim lDataType As Long
   lRet = RegCreateKey(hKey, sRegKeyPath, lKeyHandle)
   lRet = RegSetValueEx(lKeyHandle, sRegSubKey, 0&, REG_DWORD, lKeyValue, 4&)
   lRet = RegCloseKey(lKeyHandle)
End Sub

Public Sub DisableCPL(sGroup As String, sKey As String)
   Call regCreate_LongValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\" & sGroup, sKey, 1)
End Sub

Public Sub EnableCPL(sGroup As String, sKey As String)
   Call regDelete_Sub_Key(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\" & sGroup, sKey)
End Sub

Public Function IsCPLEnable(sGroup As String, sKey As String) As Boolean
   IsCPLEnable = Not regValue_Exist(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\" & sGroup, sKey)
End Function
