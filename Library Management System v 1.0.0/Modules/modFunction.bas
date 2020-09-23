Attribute VB_Name = "modFunction"
Option Explicit


''Function used to remove decimal without rounding
Public Function REMOVE_DEC(ByVal SRC_DEC As String) As Long
Dim i As Long
For i = 1 To Len(SRC_DEC)
    If Mid(SRC_DEC, i, 1) = "." Then Exit For
Next i
REMOVE_DEC = Val(Left(SRC_DEC, i)) + 1
i = 0
End Function

Public Function exist_incombo(ByVal src_combo As ComboBox, ByVal search_string As String) As Boolean
Dim i As Long
For i = 0 To src_combo.ListCount
    If src_combo.List(i) = search_string Then
        exist_incombo = True
        Exit For
    End If
Next i
If exist_incombo = True Then Exit Function
MsgBox "Please select a valid record in the list.", vbExclamation, "Information"
exist_incombo = False
i = 0
End Function

Public Function is_empty(ByRef sText As Variant) As Boolean
If sText.Text = "" Then
    is_empty = True
    MsgBox "The following information is required!", vbExclamation, "Information"
    sText.SetFocus
Else
    is_empty = False
End If
End Function
Public Function getIndex(ByVal srcTable As String) As Long
On Error GoTo err
Dim RS As New ADODB.Recordset
Dim RI As Long

RS.Open "SELECT * FROM TBL_GENERATOR WHERE TableName = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic

RI = RS.Fields("NextNo")
RS.Fields("NextNo") = RI + 1
RS.Update

getIndex = RI

srcTable = ""
RI = 0
Set RS = Nothing
Exit Function
err:
    ''Error when incounter a null value
    If err.Number = 94 Then getIndex = 1: Resume Next
End Function

Public Function getID(ByVal srcTable As String) As Long
On Error GoTo err
Dim RS As New ADODB.Recordset
RS.Open "SELECT * FROM ID_GENERATOR WHERE Table = '" & srcTable & "'", CN, adOpenStatic, adLockOptimistic
getID = RS.Fields("NextAutoNumber")

srcTable = ""
Set RS = Nothing
Exit Function
err:
    ''Error when incounter a null value
    If err.Number = 94 Then getID = 1: Resume Next
End Function

Public Function isLoaded(ByVal sForm As Form) As Boolean
If sForm.FORM_LOADED = True Then
    isLoaded = True
End If
End Function


Public Function existInLV(ByRef srcLV As ListView, ByVal strFind As String, ByVal inFirst As Boolean, Optional numCol As Byte) As Boolean
If srcLV.ListItems.Count < 1 Then Exit Function
Dim i As Long
For i = 1 To srcLV.ListItems.Count
    srcLV.ListItems(i).Selected = True
    If inFirst = True Then
        If srcLV.SelectedItem = strFind Then existInLV = True: Exit For
    Else
        If srcLV.SelectedItem.ListSubItems(numCol) = strFind Then existInLV = True: Exit For
    End If
Next i
i = 0
End Function

Public Function If_File_Exists(ByVal sFile As String) As Boolean
Dim FileLength As Long
On Error GoTo err
FileLength = FileLen(sFile)
If_File_Exists = True
Exit Function
err:
  If_File_Exists = False
End Function
Public Function GenerateID(ByVal srcNo As String, ByVal src1stStr As String, ByVal src2ndStr As String) As String
    If Len(src2ndStr) <= Len(srcNo) Then
        GenerateID = src1stStr & srcNo
    Else
        GenerateID = src1stStr & Left$(src2ndStr, Len(src2ndStr) - Len(srcNo)) & srcNo
    End If
End Function
