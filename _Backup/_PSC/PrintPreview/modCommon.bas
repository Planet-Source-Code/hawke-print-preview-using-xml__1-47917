Attribute VB_Name = "modCommon"
Option Explicit

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function CoCreateGuid Lib "ole32" (id As Any) As Long

Public Function CreateGUID() As String
On Error GoTo err_handler

    Dim bytID(0 To 15) As Byte
    Dim lngCount As Long
    
    If CoCreateGuid(bytID(0)) = 0 Then
        For lngCount = 0 To 15
            CreateGUID = CreateGUID + IIf(bytID(lngCount) < 16, "0", "") + Hex$(bytID(lngCount))
        Next
        
        CreateGUID = Left$(CreateGUID, 8) + "-" + Mid$(CreateGUID, 9, 4) + "-" + Mid$(CreateGUID, 13, 4) + "-" + Mid$(CreateGUID, 17, 4) + "-" + Right$(CreateGUID, 12)
    End If
    Exit Function

err_handler:
        CreateGUID = "ERROR"
    
End Function

Public Function IsFileExist(strFileName As String) As Boolean
On Error GoTo err_handler
  
    Call FileLen(strFileName)
    IsFileExist = True
    Exit Function
  
err_handler:
    IsFileExist = False
    
End Function

Public Function GetTempDirectory() As String
    
    Dim strTemp As String
    Dim strUserName As String
        
    strTemp = String(100, Chr$(0))  'Create a buffer
    GetTempPath 100, strTemp
    strTemp = Trim(Left$(strTemp, InStr(strTemp, Chr$(0)) - 1))
    
    If Right(strTemp, 1) <> "\" Then strTemp = strTemp & "\"
    GetTempDirectory = strTemp
    
End Function

'write any error to event log
Public Sub WriteToEventViewer(strFunction As String, strDesc As String, _
                             strNum As String, strSource As String)
    Dim strError As String
    strError = "Function: " & strFunction & vbCrLf & _
                "Description: " & strDesc & vbCrLf & _
                "Number: " & strNum & vbCrLf & _
                "Date/Time: " & Now & vbCrLf & _
                "Source: " & strSource

    App.LogEvent strError, vbLogEventTypeError

End Sub

'Change the database path to suit your needs
Public Function ConstructConnString() As String
    ConstructConnString = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & App.Path & "\DB.mdb"
End Function
