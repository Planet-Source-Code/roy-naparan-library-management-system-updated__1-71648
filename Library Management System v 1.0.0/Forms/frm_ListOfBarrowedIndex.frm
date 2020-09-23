VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_ListOfBarrowedIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Barrowed Index Records"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10680
   ClipControls    =   0   'False
   Icon            =   "frm_ListOfBarrowedIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   10680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin isCoolButton.isButton btn_Select 
      Height          =   330
      Left            =   8010
      TabIndex        =   0
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_ListOfBarrowedIndex.frx":3482
      Style           =   5
      Caption         =   "&Select"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin isCoolButton.isButton btn_Close 
      Height          =   330
      Left            =   9225
      TabIndex        =   1
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_ListOfBarrowedIndex.frx":349E
      Style           =   5
      Caption         =   "&Close"
      IconAlign       =   1
      iNonThemeStyle  =   0
      Tooltiptitle    =   ""
      ToolTipIcon     =   0
      ToolTipType     =   0
      ttForeColor     =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   4155
      Left            =   225
      TabIndex        =   3
      Top             =   1125
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index ID"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Title"
         Object.Width           =   10583
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Date Barrowed"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Date Return"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "RowIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search non-book materials from barrowed to be returned."
      Height          =   510
      Left            =   1215
      TabIndex        =   2
      Top             =   405
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_ListOfBarrowedIndex.frx":34BA
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frm_ListOfBarrowedIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Recordset used to the display the record
Dim RS_ListofIndex As New ADODB.Recordset
Dim RS_GetListofIndex As New ADODB.Recordset
Dim RS_ReservedBooks As New ADODB.Recordset
''Variable used to page the records
Dim MY_PAGE                         As PAGE_INFO
''Variable that hold the current column (Used to sorting the List View)
Dim CURR_COL                        As Integer
''Variable the current list in the page
Dim CURR_LIST                       As String
Public SRC_RI As Long
Option Explicit





Private Sub btn_Close_Click()

    Unload Me

End Sub

Private Sub btn_Select_Click()

    SRC_RI = Val(lvList.SelectedItem.ListSubItems(4))
    
    
        RS_GetListofIndex.Open "SELECT * FROM qry_BarrowedIndex WHERE tbl_BorrowedIndex.RowIndex =" & SRC_RI & "", CN2, adOpenStatic, adLockOptimistic
        
            frm_ReturnIndex.txt_BookID = RS_GetListofIndex.Fields("IndexID")
            frm_ReturnIndex.txt_BookTitle = RS_GetListofIndex.Fields("Title")
            frm_ReturnIndex.dtp_DateBarrowed = RS_GetListofIndex.Fields("DateBarrowed")
            frm_ReturnIndex.dtp_DateDue = RS_GetListofIndex.Fields("DateReturned")
         
            frm_ReturnIndex.BarrowedIndex = SRC_RI
        RS_GetListofIndex.Close
    
    Unload Me

End Sub

Private Sub Form_Load()

lvList.ColumnHeaders(1).Alignment = lvwColumnLeft

With frm_Main
 
    Set lvList.SmallIcons = .i16x16
    Set lvList.ColumnHeaderIcons = .i16x16
    Set lvList.Icons = .i16x16
    
End With

''Set the record set
CN2.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"


    RS_ListofIndex.Open "SELECT IndexID,Title,DateBarrowed,DateReturned,tbl_BorrowedIndex.RowIndex FROM qry_BarrowedIndex WHERE StudentNo LIKE '" & frm_ReturnIndex.txt_StudentNo.Text & "' AND StatusReturned =" & False & "", CN2, adOpenStatic, adLockOptimistic
    ''Call the procedure to fill the listview
    Call FILL_RECORD(1)
    

End Sub

Private Sub FILL_RECORD(ByVal SRC_PAGE As Long)
On Error Resume Next
Screen.MousePointer = vbHourglass
Dim pos_start As Long, pos_end As Long
With MY_PAGE
        .PAGE_CURRENT = 1
        .PAGE_NEXT = 1
        .PAGE_PREVIOUS = 1
        .PAGE_TOTAL = 1
        .PAGE_CURRENT = SRC_PAGE
        .PAGE_TOTAL = REMOVE_DEC("" & (RS_ListofIndex.RecordCount / 100000))
        If .PAGE_TOTAL > .PAGE_CURRENT Then
            .PAGE_NEXT = .PAGE_CURRENT + 1
        ElseIf .PAGE_CURRENT > 1 Then
            .PAGE_PREVIOUS = .PAGE_CURRENT - 1
        End If
        If .PAGE_TOTAL = 1 Then
            pos_start = 1
            pos_end = RS_ListofIndex.RecordCount
         
        ElseIf .PAGE_CURRENT = 1 Then
            pos_start = 1
            pos_end = 100000
         
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            pos_start = ((.PAGE_CURRENT - 1) * 100000) + 1
            pos_end = RS_ListofIndex.RecordCount
        
        Else
            pos_start = ((.PAGE_CURRENT - 1) * 100000) + 1
            pos_end = (.PAGE_NEXT - 1) * 100000
         
        End If
End With
''Reset the sorting
'lvList.SortOrder = 1
'lvList.Sorted = True

            ''Sort the listview
            If 2 - 1 <> CURR_COL Then
            lvList.SortOrder = 0
            Else
            lvList.SortOrder = Abs(lvList.SortOrder - 1)
            End If
            lvList.SortKey = 2 - 1
            
            lvList.Sorted = True
            CURR_COL = 2 - 1

Call FillListView(lvList, RS_ListofIndex, pos_start, pos_end, 5, 8, False, False)

''Clear variables
pos_start = 0
pos_end = 0
Screen.MousePointer = vbDefault
End Sub



''Filter record
Public Sub FILTER_REC()
If RS_ListofIndex.RecordCount < 1 Then MsgBox "No record match.", vbInformation, "Search Result"
Call FILL_RECORD(1)
End Sub
''Proccess command from other form
Public Sub COMMAND_PASS(ByVal SEL_COMMAND As Integer)
Select Case SEL_COMMAND
    ''Reload record
    Case 1:
       ' RS_ListofIndex.Filter = adFilterNone
        RS_ListofIndex.Requery
        Call FILL_RECORD(1)
        ''Display record status


End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)
''Cleanup variables
RS_ListofIndex.Close
Set RS_ListofIndex = Nothing
Set CN2 = Nothing



CURR_COL = 0
CURR_LIST = ""




End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
''Sort the listview
If ColumnHeader.Index - 1 <> CURR_COL Then
    lvList.SortOrder = 0
Else
    lvList.SortOrder = Abs(lvList.SortOrder - 1)
End If
lvList.SortKey = ColumnHeader.Index - 1

lvList.Sorted = True
CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub lvList_DblClick()
Call COMMAND_PASS(5)
End Sub



Private Sub Command3_Click()
If MY_PAGE.PAGE_CURRENT <> 1 Then Call FILL_RECORD(1)
End Sub

Private Sub Command4_Click()
If MY_PAGE.PAGE_CURRENT <> 1 Then Call FILL_RECORD(MY_PAGE.PAGE_PREVIOUS)
End Sub

Private Sub Command5_Click()
If MY_PAGE.PAGE_CURRENT <> MY_PAGE.PAGE_TOTAL Then Call FILL_RECORD(MY_PAGE.PAGE_NEXT)
End Sub

Private Sub Command6_Click()
If MY_PAGE.PAGE_CURRENT <> MY_PAGE.PAGE_TOTAL Then Call FILL_RECORD(MY_PAGE.PAGE_TOTAL)
End Sub

