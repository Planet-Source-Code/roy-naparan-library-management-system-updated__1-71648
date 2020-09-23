Attribute VB_Name = "modProcedures"
Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
Dim tmp_listtview As ListItem
Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem + lvwText, lvwPartial, lvwPartial)
If Not tmp_listtview Is Nothing Then
    tmp_listtview.EnsureVisible
    tmp_listtview.Selected = True
End If
End Sub
Public Sub fill_combo(ByRef srcCombo As ComboBox, ByVal srcSQL As String, ByVal srcField As String, Optional srcFirsttext As String)
Dim rs As New ADODB.Recordset
srcCombo.Clear
If srcFirsttext <> "" Then srcCombo.AddItem srcFirsttext

rs.Open srcSQL, CN, adOpenStatic, adLockReadOnly
With rs
    Do While Not .EOF
        srcCombo.AddItem .Fields(srcField)
        .MoveNext
    Loop
End With
Set rs = Nothing
End Sub
Public Sub fillCombo(ByRef srcComboBox As ComboBox, ByVal srcTable As String, ByVal srcField As String, Optional srcFirsttext As String)
On Error GoTo ERR
srcComboBox.Clear
If srcFirsttext <> "" Then srcComboBox.AddItem srcFirsttext
Dim rs As New ADODB.Recordset
With rs
    .Open "SELECT * FROM " & srcTable & " ORDER BY " & srcField, CN, adOpenStatic, adLockOptimistic
    .MoveFirst
    Do While Not .EOF
        srcComboBox.AddItem .Fields(srcField)
        .MoveNext
    Loop
End With
srcTable = ""
srcField = ""
Set rs = Nothing
Exit Sub
ERR:
    Resume Next
End Sub
Public Sub FillListView(ByRef sListView As ListView, ByRef sRecordSource As ADODB.Recordset, ByVal pos_start As Long, ByVal pos_end As Long, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean, Optional match_field As String, Optional match_str As String, Optional match_ico As Byte)
'Optional to be declare as variant
Dim x As Variant
Dim i As Byte
sListView.ListItems.Clear
If sRecordSource.RecordCount < 1 Then Exit Sub
sRecordSource.AbsolutePosition = pos_start
On Error Resume Next
Do
    If match_field = "" Then
        If with_num = True Then
            Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
        Else
            Set x = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
        End If
    Else
        If sRecordSource.Fields(match_field) = match_str Then
            If with_num = True Then
                Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, match_ico, match_ico)
            Else
                Set x = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), match_ico, match_ico)
            End If
        Else
            If with_num = True Then
                Set x = sListView.ListItems.Add(, , "" & sRecordSource.AbsolutePosition, sNumIco, sNumIco)
            Else
                Set x = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
            End If
        End If
    End If
        For i = 1 To sNumOfFields - 1
            If Not sRecordSource.Fields(Val(i)) = "" Then
                If show_first_rec = True Then
                    x.SubItems(i) = "" & Replace(sRecordSource.Fields(Val(i) - 1), vbCrLf, " ", , , vbBinaryCompare)
                Else
                    x.SubItems(i) = "" & Replace(sRecordSource.Fields(Val(i)), vbCrLf, " ", , , vbBinaryCompare)
                End If
            End If
        Next i
    If sRecordSource.AbsolutePosition >= pos_end Then
        Exit Do
    Else
        sRecordSource.MoveNext
    End If
Loop
i = 0
Set x = Nothing
End Sub
'Procedure used to fill combo for sorting record with Data Grid
Public Sub fill_for_sort_wDG(ByRef sCombo As ComboBox, ByVal sDataGrid As DataGrid)
Dim i As Integer
With sCombo
    For i = 0 To sDataGrid.Columns.Count - 1
        If sDataGrid.Columns(i).Visible = True Then
            .AddItem sDataGrid.Columns(i).DataField & " Asc"
            .AddItem sDataGrid.Columns(i).DataField & " Desc"
        End If
    Next i
End With
i = 0
End Sub
'Procedure used to fill combo for sorting record with ADODB recordset
Public Sub fill_for_sort_wRS(ByRef sCombo As ComboBox, ByVal sRS As ADODB.Recordset, ByVal sHaveStopper As Boolean, Optional sStopNum As Byte)
Dim x As Long
With sCombo
    For x = 1 To sRS.Fields.Count - 1
        If sHaveStopper = True Then
            If x = sStopNum Then Exit For
        End If
        .AddItem sRS.Fields.Item(x).Name & " Asc"
        .AddItem sRS.Fields.Item(x).Name & " Desc"
    Next x
End With
End Sub
'Procedure used to highlight text when focus
Public Sub hl_text(ByRef sText)
With sText
    .SelStart = 0
    .SelLength = Len(sText.Text)
End With
End Sub
'Procedure used to set Datagrid Divider Color
Public Sub set_dg_row_col_color(ByRef sDataGrid As DataGrid)
Dim x As Long
With sDataGrid
    .RowDividerStyle = dbgLightGrayLine
    For x = 0 To .Columns.Count - 1
        .Columns(x).DividerStyle = dbgLightGrayLine
    Next x
End With
x = 0
End Sub
'Procedure used to fill combo with Data Grid' data field name
Public Sub fill_wDG(ByRef sCombo As ComboBox, ByVal sDataGrid As DataGrid)
Dim i As Integer
With sCombo
    For i = 0 To sDataGrid.Columns.Count - 1
        If sDataGrid.Columns(i).Visible = True Then
            .AddItem sDataGrid.Columns(i).DataField
        End If
    Next i
End With
i = 0
End Sub

Public Sub FillLWV(ByRef sListView As ListView, ByRef sRecordSource As ADODB.Recordset, ByVal sNumOfFields As Byte, ByVal sNumIco As Byte, ByVal with_num As Boolean, ByVal show_first_rec As Boolean)
Dim x As Variant
Dim i As Byte
On Error Resume Next
sListView.ListItems.Clear
If sRecordSource.RecordCount < 1 Then Exit Sub
sRecordSource.MoveFirst
Do While Not sRecordSource.EOF
    If with_num = True Then
        Set x = sListView.ListItems.Add(, , sRecordSource.AbsolutePosition, sNumIco, sNumIco)
    Else
        Set x = sListView.ListItems.Add(, , "" & sRecordSource.Fields(0), sNumIco, sNumIco)
    End If
        For i = 1 To sNumOfFields - 1
            If Not sRecordSource.Fields(Val(i)) = "" Then
                If show_first_rec = True Then
                    x.SubItems(i) = sRecordSource.Fields(Val(i) - 1)
                Else
                    x.SubItems(i) = sRecordSource.Fields(Val(i))
                End If
            End If
        Next i
    sRecordSource.MoveNext
Loop
i = 0
Set x = Nothing
End Sub

Public Sub customMove(ByRef sRS As ADODB.Recordset, ByVal isNum As Boolean, ByVal findStr As String, ByVal sField As String)
If sRS.RecordCount < 1 Then Exit Sub
Dim old_pos As Long
sRS.MoveFirst
old_pos = sRS.AbsolutePosition
If isNum = True Then
    sRS.Find sField & " = " & findStr
Else
    sRS.Find sField & " = '" & findStr & "'"
End If
If sRS.EOF Then sRS.AbsolutePosition = old_pos
old_pos = 0
End Sub

'Procedure used to promp unexpected errors
Public Sub prompt_err(ByVal sError As ErrObject)
MsgBox sError.Description & vbCrLf & vbCrLf & "Error Number: " & ERR.Number & vbCrLf & vbCrLf & "*Note: See the error information in the help file to get more information about this.", vbExclamation
End Sub
'Procedure used to save file in the database
Public Sub save_file_to_db(ByRef sRS As ADODB.Recordset, ByVal sField As String, ByVal sFilePath As String)
    Dim FILE_ARRAY() As Byte
    Dim FILE_POINTER As Long
    Dim FILE_SIZE As Long
    
    FILE_POINTER = lOpen(sFilePath, OF_READ)
    'Get the file size
    FILE_SIZE = GetFileSize(FILE_POINTER, lpFSHigh)
    lclose Pointer

    'Create a byte array for file
    ReDim FILE_ARRAY(FILE_SIZE)

    Open sFilePath For Binary Access Read As #1
    Get #1, , FILE_ARRAY
    Close #1
    sRS(sField).Value = FILE_ARRAY
    Exit Sub
    
End Sub

Public Sub updateID(ByVal srcTable As String)
On Error GoTo ERR
Dim rs As New ADODB.Recordset
rs.Open "SELECT * FROM ID_GENERATOR WHERE Table = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
rs.Fields("NextAutoNumber") = rs.Fields("NextAutoNumber") + 1
rs.Update

srcTable = ""
Set rs = Nothing
Exit Sub
ERR:
    ''Error when incounter a null value
    If ERR.Number = 94 Then Resume Next
End Sub

''Procedure used to center horizontal
Public Sub centerFormHorizontal(ByRef sForm As Form, ByVal sWidth As Integer)
    sForm.Left = (sWidth - sForm.Width) / 2
End Sub

''Procedure used to clear the text content
Public Sub clearText(ByRef sForm As Form)
Dim CONTROL As CONTROL
For Each CONTROL In sForm.Controls
    If (TypeOf CONTROL Is TextBox) Then CONTROL = vbNullString
Next CONTROL
Set CONTROL = Nothing
End Sub

''Procedure used to clear the data field
Public Sub clearDataField(ByRef sRpt As DataReport)
sRpt.DataMember = vbNullString
Dim i As Integer
For i = 1 To sRpt.Sections(3).Controls.Count
    If (TypeOf sRpt.Sections(3).Controls(i) Is RptTextBox) Then sRpt.Sections(3).Controls(i).DataMember = vbNullString
Next i
i = 0
End Sub

