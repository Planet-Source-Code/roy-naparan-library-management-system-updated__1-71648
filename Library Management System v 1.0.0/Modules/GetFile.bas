Attribute VB_Name = "GetFile"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long
Public Sub OpenURL(urlADD As String, sourceHWND As Long)
     Call ShellExecute(sourceHWND, vbNullString, urlADD, "", vbNullString, 1)
End Sub







