VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const HKEY_CURRENT_USER As Long = &H80000001

Const ERROR_SUCCESS = 0&
Const KEY_ALL_ACCESS = &H3F
Const REG_OPTION_NON_VOLATILE = 0

Const REG_SZ = 1
Const REG_BINARY = 3
Const REG_DWORD = 4

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Function StringFromBuffer(Buffer As String) As String
    Dim lngPos As Long

    lngPos = InStr(Buffer, vbNullChar)
    If lngPos > 0 Then
        StringFromBuffer = Left$(Buffer, lngPos - 1)
    Else
        StringFromBuffer = Buffer
    End If
End Function

Private Function getRegKeyHandle(lngParentKey As Long, strKeyName As String, lngKeyHandle As Long) As Boolean
    If RegOpenKey(lngParentKey, strKeyName, lngKeyHandle) = ERROR_SUCCESS Then getRegKeyHandle = True
End Function

Private Function setRegValueString(ByVal hKeyHandle As Long, ByVal strValueName As String, ByVal strData As String) As Boolean
    If RegSetValueEx(hKeyHandle, strValueName, 0&, REG_SZ, ByVal strData, LenB(StrConv(strData, vbFromUnicode)) + 1) = ERROR_SUCCESS Then setRegValueString = True
End Function

'use this function to read the registry in order to set the footer to ""
'this is to prevent user from printing out the url of the xml file.
Public Sub ResetPrintFooter()
    
    Dim lngHandle As Long
        
    If getRegKeyHandle(HKEY_CURRENT_USER, "Software\Microsoft\Internet Explorer\PageSetup", lngHandle) Then
        Call setRegValueString(lngHandle, "footer", "")
    End If

End Sub

