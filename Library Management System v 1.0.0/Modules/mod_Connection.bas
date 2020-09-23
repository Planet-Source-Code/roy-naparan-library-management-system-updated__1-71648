Attribute VB_Name = "mod_Connection"
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000


Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public welcometime As Integer
Public Welcome As Boolean
Public view As Integer

Option Explicit

Sub Main()

    frm_Login.Show 1
    Call AssignValue
    de_Connection.RS_Connection.ConnectionString = CN
    
End Sub

Public Function MakeTransparent(ByVal hwnd As Long, Perc As Integer) As Long
Dim Msg As Long
On Error Resume Next
If Perc < 0 Or Perc > 255 Then
  MakeTransparent = 1
Else
  Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
  Msg = Msg Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, Msg
  SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
  MakeTransparent = 0
End If
If err Then
  MakeTransparent = 2
End If
End Function

Public Sub AssignValue()

    Dim RS_Assign As New ADODB.Recordset

    CN4.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    RS_Assign.Open "SELECT * FROM tbl_Settings", CN4, adOpenStatic, adLockOptimistic
    
        With RS_Assign
    
            dayslimit = .Fields("MaxDayHold")
            maxhold = .Fields("MaxHold")
            Fines = .Fields("Fines")
            rateperday = .Fields("FineCharge")
            
        End With
        
End Sub
