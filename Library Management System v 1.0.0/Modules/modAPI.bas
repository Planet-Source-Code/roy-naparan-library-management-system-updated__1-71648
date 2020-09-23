Attribute VB_Name = "modAPI"
Option Explicit

''API Declarations

'[For getting file size]
Public Declare Function lOpen Lib "kernel32" Alias "_lopen" (ByVal lpPathName As String, ByVal iReadWrite As Long) As Long
Public Declare Function lclose Lib "kernel32" Alias "_lclose" (ByVal hFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long

''API Constant
'[For getting file size]
Public Const OF_READ = &H0&
Public lpFSHigh As Long

