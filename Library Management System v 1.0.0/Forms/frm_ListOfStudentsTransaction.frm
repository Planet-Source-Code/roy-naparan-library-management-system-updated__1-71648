VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_ListOfStudentsTransaction 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Records"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12105
   ClipControls    =   0   'False
   Icon            =   "frm_ListOfStudentsTransaction.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleMode       =   0  'User
   ScaleWidth      =   12105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2880
      TabIndex        =   5
      Top             =   5490
      Width           =   2580
   End
   Begin VB.ComboBox cbx_Type 
      Height          =   315
      ItemData        =   "frm_ListOfStudentsTransaction.frx":3482
      Left            =   1125
      List            =   "frm_ListOfStudentsTransaction.frx":3492
      TabIndex        =   4
      Top             =   5490
      Width           =   1725
   End
   Begin isCoolButton.isButton btn_Select 
      Height          =   330
      Left            =   8235
      TabIndex        =   0
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_ListOfStudentsTransaction.frx":34BF
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
   Begin isCoolButton.isButton btn_Search 
      Height          =   345
      Left            =   5490
      TabIndex        =   1
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Icon            =   "frm_ListOfStudentsTransaction.frx":34DB
      Style           =   5
      Caption         =   "&Search"
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
      Left            =   10665
      TabIndex        =   2
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_ListOfStudentsTransaction.frx":34F7
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
      TabIndex        =   7
      Top             =   1125
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   7329
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
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
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Student No."
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Student Name"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Current Address"
         Object.Width           =   7056
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "E-mail Address"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Contact No."
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Course"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Level"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Status"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Comment"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Date Modified"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "RowIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin isCoolButton.isButton btn_Reload 
      Height          =   330
      Left            =   9450
      TabIndex        =   8
      Top             =   5490
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_ListOfStudentsTransaction.frx":3513
      Style           =   5
      Caption         =   "&Reload"
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
   Begin VB.Label Label2 
      Caption         =   "Search by:"
      Height          =   195
      Left            =   225
      TabIndex        =   6
      Top             =   5535
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search library members for the book to be issued."
      Height          =   510
      Left            =   1215
      TabIndex        =   3
      Top             =   405
      Width           =   2715
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_ListOfStudentsTransaction.frx":352F
      Top             =   180
      Width           =   720
   End
End
Attribute VB_Name = "frm_ListOfStudentsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Recordset used to the display the record
Dim RS_Student As New ADODB.Recordset
Dim RS_SelectStudent As New ADODB.Recordset


''Variable used to page the records
Dim MY_PAGE                         As PAGE_INFO
''Variable that hold the current column (Used to sorting the List View)
Dim CURR_COL                        As Integer
''Variable the current list in the page
Dim CURR_LIST                       As String
Public SRC_RI As Long


Public bFilter As String
Public eFilter As String


Option Explicit




Private Sub btn_Add_Click()

End Sub

Private Sub btn_Close_Click()

    Unload Me

End Sub

Private Sub btn_Reload_Click()

    Call COMMAND_PASS(1)

End Sub

Private Sub btn_Search_Click()

    If IsEmpty(Text1) = True Then Exit Sub
    RS_Student.Filter = ""
    RS_Student.Requery
    eFilter = "%'"
    
        If cbx_Type.Text = "Student No." Then
            
'            If isIssue = True Then
        
                bFilter = "StudentNo LIKE '%"
            
'            ElseIf isIssue = False Then
'
'                bFilter = "StudentNo LIKE '%"
'
'            End If
        
        Else
        If cbx_Type.Text = "Full Name" Then
        
            bFilter = "StudentName LIKE '%"
            
        Else
        If cbx_Type.Text = "Address" Then
        
            bFilter = "CurrentAddress LIKE '%"
        
        Else
        If cbx_Type.Text = "Course" Then
        
            bFilter = "Course LIKE '%"
        
        
        End If
        End If
        End If
        End If
        
    
    RS_Student.Filter = bFilter & Text1.Text & eFilter
    Call FILTER_REC

End Sub

Private Sub btn_Select_Click()

    SRC_RI = Val(lvList.SelectedItem.ListSubItems(10))
    
        RS_SelectStudent.Open "SELECT * FROM tbl_Students WHERE RowIndex =" & SRC_RI & "", CN2, adOpenStatic, adLockOptimistic
        
        If isIssue = True Then
        
            frm_IssueBooks.txt_StudentNo = RS_SelectStudent.Fields("StudentNo")
            frm_IssueBooks.txt_StudentName = RS_SelectStudent.Fields("StudentName")
            
        Else
        
            frm_ReturnBooks.txt_StudentNo = RS_SelectStudent.Fields("StudentNo")
            frm_ReturnBooks.txt_StudentName = RS_SelectStudent.Fields("StudentName")
            
        End If
        
        RS_SelectStudent.Close
            
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
    
    If isIssue = True Then
    
        RS_Student.Open "SELECT * FROM tbl_Students", CN2, adOpenStatic, adLockOptimistic
    
    ElseIf isIssue = False Then
        
        RS_Student.Open "SELECT * FROM tbl_Students WHERE BookinHand >=" & 1 & "", CN2, adOpenStatic, adLockOptimistic
    
    End If
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
        .PAGE_TOTAL = REMOVE_DEC("" & (RS_Student.RecordCount / 100000))
        If .PAGE_TOTAL > .PAGE_CURRENT Then
            .PAGE_NEXT = .PAGE_CURRENT + 1
        ElseIf .PAGE_CURRENT > 1 Then
            .PAGE_PREVIOUS = .PAGE_CURRENT - 1
        End If
        If .PAGE_TOTAL = 1 Then
            pos_start = 1
            pos_end = RS_Student.RecordCount
         
        ElseIf .PAGE_CURRENT = 1 Then
            pos_start = 1
            pos_end = 100000
         
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            pos_start = ((.PAGE_CURRENT - 1) * 100000) + 1
            pos_end = RS_Student.RecordCount
        
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

Call FillListView(lvList, RS_Student, pos_start, pos_end, 11, 6, False, False)

''Clear variables
pos_start = 0
pos_end = 0
Screen.MousePointer = vbDefault
End Sub



''Filter record
Public Sub FILTER_REC()
If RS_Student.RecordCount < 1 Then
MsgBox "No record match. Please try again.", vbInformation, "Search Result"
Else
Call FILL_RECORD(1)
End If
End Sub
''Proccess command from other form
Public Sub COMMAND_PASS(ByVal SEL_COMMAND As Integer)
Select Case SEL_COMMAND
    ''Reload record
    Case 1:
        RS_Student.Filter = adFilterNone
        RS_Student.Requery
        Call FILL_RECORD(1)

End Select
End Sub



Private Sub Form_Unload(Cancel As Integer)
''Cleanup variables
RS_Student.Close
Set RS_Student = Nothing
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

