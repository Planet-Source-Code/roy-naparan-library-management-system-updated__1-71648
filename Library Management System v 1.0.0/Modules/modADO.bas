Attribute VB_Name = "modADO"
''Set the main ADODB connection
Public Sub connect_to_db()
CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
End Sub

''Function used to check if the record exit or not.
Public Function rec_exist(ByVal sTable As String, ByVal sField As String, ByVal sStr As String, Optional isString As Boolean) As Boolean
Dim rs As New ADODB.Recordset
If isString = False Then
    rs.Open "Select * From " & sTable & " Where " & sField & " = " & sStr, CN, adOpenStatic, adLockOptimistic
Else
    rs.Open "Select * From " & sTable & " Where " & sField & " = '" & sStr & "'", CN, adOpenStatic, adLockOptimistic
End If
If rs.RecordCount < 1 Then
    rec_exist = False
Else
    rec_exist = True
End If
Set rs = Nothing
End Function

''Procedure used to delete record with SQL
Public Sub del_rec_wSQL(ByVal sCN As ADODB.Connection, ByVal sTable As String, ByVal sField As String, ByVal sString As String, ByVal isNumber As Boolean, ByVal snum As Long)
If isNumber = True Then
    sCN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " =" & snum
Else
    sCN.Execute "DELETE FROM " & sTable & " WHERE " & sField & " ='" & sString & "'"
End If
End Sub
''Procedure used to get the generated id
Public Function autoId(ByVal sTable As String) As String
Dim rs As New ADODB.Recordset
With rs
    .Open "Select autoIdGenerator.* From autoIdGenerator Where SourceTable = '" & sTable & "'", CN, adOpenStatic, adLockReadOnly
        autoId = .Fields("InitialLetter") & .Fields("Separator") & Left(.Fields("MZBN"), (Len(.Fields("MZBN")) - Len(.Fields("NextNumber")))) & .Fields("NextNumber")
End With
Set rs = Nothing
End Function

''Procedure used to bind data combo
Public Sub bind_dc(ByVal sSQL As String, ByVal bindField As String, ByVal sCN As ADODB.Connection, ByRef sDC As DataCombo)
Dim rs As New ADODB.Recordset

rs.Open sSQL, CN, adOpenStatic, adLockOptimistic

sDC.ListField = bindField
Set sDC.RowSource = rs
sDC.Tag = rs.RecordCount

Set rs = Nothing
End Sub

Public Function validDB(ByVal srcPath As String) As Boolean
On Error GoTo err
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset

conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & srcPath & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
rs.Open "SELECT * FROM tbl_Identity", conn, adOpenForwardOnly, adLockReadOnly
If rs.Fields("Identity") = "library" Then
    validDB = True
Else
    validDB = False
End If

Set rs = Nothing
Set conn = Nothing
Exit Function
err:
    validDB = False
End Function


