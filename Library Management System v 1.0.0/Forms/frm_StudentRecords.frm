VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_StudentRecords 
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   12690
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   0
      ScaleHeight     =   2805
      ScaleWidth      =   12690
      TabIndex        =   14
      Top             =   5625
      Width           =   12690
      Begin VB.ComboBox cbx_Status 
         Height          =   315
         ItemData        =   "frm_StudentRecords.frx":0000
         Left            =   1395
         List            =   "frm_StudentRecords.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   90
         Width           =   2040
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   8190
         ScaleHeight     =   330
         ScaleWidth      =   4485
         TabIndex        =   17
         Top             =   90
         Width           =   4485
         Begin isCoolButton.isButton btn_hide 
            Height          =   330
            Left            =   3960
            TabIndex        =   18
            Top             =   0
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":002F
            Style           =   5
            Caption         =   "isButton"
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
         Begin isCoolButton.isButton btn_Show 
            Height          =   330
            Left            =   3510
            TabIndex        =   19
            Top             =   0
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":0A41
            Style           =   5
            Caption         =   "isButton"
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
      End
      Begin MSComctlLib.ListView lv_Books 
         Height          =   2265
         Left            =   45
         TabIndex        =   15
         Top             =   495
         Width           =   12615
         _ExtentX        =   22251
         _ExtentY        =   3995
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
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Book ID"
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
            Text            =   "Book Source"
            Object.Width           =   4762
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Fines"
            Object.Width           =   4410
         EndProperty
      End
      Begin isCoolButton.isButton btn_Print 
         Height          =   330
         Left            =   3465
         TabIndex        =   21
         Top             =   90
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":1453
         Style           =   5
         Caption         =   "isButton"
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
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   45
         X2              =   12330
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "frm_StudentRecords.frx":1E65
         Top             =   90
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "List of Books:"
         Height          =   240
         Left            =   360
         TabIndex        =   16
         Top             =   135
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   12690
      TabIndex        =   1
      Top             =   5115
      Width           =   12690
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   8325
         ScaleHeight     =   465
         ScaleWidth      =   4305
         TabIndex        =   2
         Top             =   0
         Width           =   4305
         Begin isCoolButton.isButton command3 
            Height          =   330
            Left            =   1980
            TabIndex        =   3
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":2867
            Style           =   5
            Caption         =   "isButton"
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
         Begin isCoolButton.isButton command4 
            Height          =   330
            Left            =   2565
            TabIndex        =   4
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":3279
            Style           =   5
            Caption         =   "isButton"
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
         Begin isCoolButton.isButton command5 
            Height          =   330
            Left            =   3150
            TabIndex        =   5
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":3C8B
            Style           =   5
            Caption         =   "isButton"
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
         Begin isCoolButton.isButton command6 
            Height          =   330
            Left            =   3735
            TabIndex        =   6
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_StudentRecords.frx":469D
            Style           =   5
            Caption         =   "isButton"
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
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   360
            TabIndex        =   7
            Top             =   180
            Width           =   1545
         End
      End
      Begin isCoolButton.isButton btn_Add 
         Height          =   330
         Left            =   45
         TabIndex        =   8
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":50AF
         Style           =   5
         Caption         =   "&Add"
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
      Begin isCoolButton.isButton btn_Edit 
         Height          =   330
         Left            =   1260
         TabIndex        =   9
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":50CB
         Style           =   5
         Caption         =   "&Edit"
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
         Height          =   330
         Left            =   2475
         TabIndex        =   10
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":50E7
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
      Begin isCoolButton.isButton btn_Delete 
         Height          =   330
         Left            =   4905
         TabIndex        =   11
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":5103
         Style           =   5
         Caption         =   "&Delete"
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
         Left            =   6120
         TabIndex        =   12
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":511F
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
      Begin isCoolButton.isButton btn_Reload 
         Height          =   330
         Left            =   3690
         TabIndex        =   13
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_StudentRecords.frx":513B
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
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5055
      Left            =   -45
      TabIndex        =   0
      Top             =   -45
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   8916
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
End
Attribute VB_Name = "frm_StudentRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
''Recordset used to the display the record
Dim RS As New ADODB.Recordset
Dim RS_BarrowedBooks As New ADODB.Recordset
Dim RS_FindStudentID As New ADODB.Recordset
Private RS_rptListofBarrowedBooks As New ADODB.Recordset

Dim Selected As Integer

Dim StudentNo As String
Dim SName As String
Dim Address As String
Dim ContactNo As String
Dim Course As String
Dim Year As String

''Variable used to page the records
Dim MY_PAGE                         As PAGE_INFO
Dim MY_PAGE2                        As PAGE_INFO
''Variable that hold the current column (Used to sorting the List View)
Dim CURR_COL                        As Integer
''Variable the current list in the page
Dim CURR_LIST                       As String



Private Sub btn_Add_Click()

    COMMAND_PASS (4)
    
End Sub

Private Sub btn_Close_Click()

    Unload Me

End Sub

Private Sub btn_Delete_Click()

    COMMAND_PASS (3)

End Sub

Private Sub btn_Edit_Click()

    COMMAND_PASS (5)

End Sub

Private Sub btn_hide_Click()

    Picture3.Height = 475
    lvList.Height = (Me.ScaleHeight - (Picture1.ScaleHeight + 30 + Picture3.ScaleHeight))

End Sub

Private Sub btn_Print_Click()

    On Error Resume Next
    CN12.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    CN14.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    
    
    StudentNo = lvList.SelectedItem.ListSubItems(10)
    SName = lvList.SelectedItem.ListSubItems(1)
    Address = lvList.SelectedItem.ListSubItems(2)
    ContactNo = lvList.SelectedItem.ListSubItems(4)
    Course = lvList.SelectedItem.ListSubItems(5)
    Year = lvList.SelectedItem.ListSubItems(6)
    
        RS_FindStudentID.Open "SELECT StudentNo FROM tbl_Students WHERE RowIndex LIKE '" & StudentNo & "'", CN14, adOpenStatic, adLockOptimistic
        
    
        rpt_StudentTransaction.Sections("Section2").Controls("lblStudentNo").Caption = RS_FindStudentID.Fields("StudentNo")
        rpt_StudentTransaction.Sections("Section2").Controls("lblName").Caption = SName
        rpt_StudentTransaction.Sections("Section2").Controls("lblAddress").Caption = Address
        rpt_StudentTransaction.Sections("Section2").Controls("lblContactNo").Caption = ContactNo
        rpt_StudentTransaction.Sections("Section2").Controls("lblCourse").Caption = Course
        rpt_StudentTransaction.Sections("Section2").Controls("lblYear").Caption = Year
    
        CN14.Close
    
    If cbx_Status.Text = "Barrowed" Then
 
        RS_rptListofBarrowedBooks.Open "SELECT * FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "' AND  StatusReturned =" & False & "", CN12, adOpenStatic, adLockOptimistic
        
        Set rpt_StudentTransaction.DataSource = RS_rptListofBarrowedBooks
        
        rpt_StudentTransaction.Show 1
        RS_rptListofBarrowedBooks.Close
        CN12.Close
        
        
    ElseIf cbx_Status.Text = "Returned" Then
    
        RS_rptListofBarrowedBooks.Open "SELECT * FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & True & "", CN12, adOpenStatic, adLockOptimistic
        
        Set rpt_StudentTransaction.DataSource = RS_rptListofBarrowedBooks
        
        rpt_StudentTransaction.Show 1
        RS_rptListofBarrowedBooks.Close
        CN12.Close
        
    ElseIf cbx_Status.Text = "View All" Then
    
        RS_rptListofBarrowedBooks.Open "SELECT * FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "'", CN12, adOpenStatic, adLockOptimistic
        
        Set rpt_StudentTransaction.DataSource = RS_rptListofBarrowedBooks
        
        rpt_StudentTransaction.Show 1
        RS_rptListofBarrowedBooks.Close
        CN12.Close
        
    End If

End Sub

Private Sub btn_Reload_Click()

    COMMAND_PASS (1)

End Sub

Private Sub btn_Search_Click()

   With frm_SearchRecords
    
        Set .SRC_FORM = Me
        Set .RS = RS
        .cbx_Type.Clear
        .cbx_Type.AddItem "Name"
        .cbx_Type.AddItem "Student No."
        .eFilter = "%'"
        .Show vbModal
        
    End With

End Sub

Private Sub btn_Show_Click()

    Picture3.Height = 2805
    lvList.Height = (Me.ScaleHeight - (Picture1.ScaleHeight + 30 + Picture3.ScaleHeight))
    
End Sub

Private Sub Combo1_Change()

    

End Sub

Private Sub cbx_Status_Click()

 
    If cbx_Status.Text = "Barrowed" Then

        RS_BarrowedBooks.Open "SELECT BookID,Title,DateBarrowed,DateReturned,BookSource,Fines FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & False & "", CN, adOpenStatic, adLockOptimistic
        
        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedBooks.Close
        
    ElseIf cbx_Status.Text = "Returned" Then
    
        RS_BarrowedBooks.Open "SELECT BookID,Title,DateBarrowed,DateReturned,BookSource,Fines FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & True & "", CN, adOpenStatic, adLockOptimistic
   
        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedBooks.Close
        
    ElseIf cbx_Status.Text = "View All" Then
        
        RS_BarrowedBooks.Open "SELECT BookID,Title,DateBarrowed,DateReturned,BookSource,Fines FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "'", CN, adOpenStatic, adLockOptimistic
   
        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedBooks.Close
 
    End If

    
    
End Sub

Private Sub Form_Activate()

    frm_Main.toolbar_Menu.Buttons(6).Value = tbrPressed

    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Student Records"

End Sub

Private Sub Form_Resize()
On Error Resume Next
lvList.Width = Me.ScaleWidth
lvList.Height = (Me.ScaleHeight - (Picture1.ScaleHeight + 30 + Picture3.ScaleHeight))

End Sub

Private Sub Form_Load()

command3.IconSize = 16
command4.IconSize = 16
command5.IconSize = 16
command6.IconSize = 16

command3.Caption = ""
command4.Caption = ""
command5.Caption = ""
command6.Caption = ""

command3.IconAlign = isbCenter
command4.IconAlign = isbCenter
command5.IconAlign = isbCenter
command6.IconAlign = isbCenter

btn_Show.Caption = ""
btn_hide.Caption = ""
btn_Print.Caption = ""
btn_Show.IconAlign = isbCenter
btn_hide.IconAlign = isbCenter
btn_Print.IconAlign = isbCenter


'Picture3.Height = 475
With frm_Main
 
    Set lvList.SmallIcons = .i16x16
    Set lvList.ColumnHeaderIcons = .i16x16
    Set lvList.Icons = .i16x16
    
    Set lv_Books.SmallIcons = .i16x16
    Set lv_Books.ColumnHeaderIcons = .i16x16
    Set lv_Books.Icons = .i16x16
    
End With

''Set the record set
Call connect_to_db
RS.Open "SELECT * FROM tbl_Students ORDER BY StudentName ASC", CN, adOpenStatic, adLockOptimistic

    
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
        .PAGE_TOTAL = REMOVE_DEC("" & (RS.RecordCount / 100))
        If .PAGE_TOTAL > .PAGE_CURRENT Then
            .PAGE_NEXT = .PAGE_CURRENT + 1
        ElseIf .PAGE_CURRENT > 1 Then
            .PAGE_PREVIOUS = .PAGE_CURRENT - 1
        End If
        If .PAGE_TOTAL = 1 Then
            pos_start = 1
            pos_end = RS.RecordCount
            command3.Enabled = False
            command4.Enabled = False
            command5.Enabled = False
            command6.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            pos_start = 1
            pos_end = 100
            command3.Enabled = False
            command4.Enabled = False
            command5.Enabled = True
            command6.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            pos_start = ((.PAGE_CURRENT - 1) * 100) + 1
            pos_end = RS.RecordCount
            command3.Enabled = True
            command4.Enabled = True
            command5.Enabled = False
            command6.Enabled = False
        Else
            pos_start = ((.PAGE_CURRENT - 1) * 100) + 1
            pos_end = (.PAGE_NEXT - 1) * 100
            command3.Enabled = True
            command4.Enabled = True
            command5.Enabled = True
            command6.Enabled = True
        End If
End With


        ''Sort the listview
        If 2 - 1 <> CURR_COL Then
        lvList.SortOrder = 0
        Else
        lvList.SortOrder = Abs(lvList.SortOrder - 1)
        End If
        lvList.SortKey = 2 - 1
        
        lvList.Sorted = True
        'CURR_COL = 2 - 1


Call FillListView(lvList, RS, pos_start, pos_end, 11, 6, False, False)
Call TRACK_LIST
''Clear variables
pos_start = 0
pos_end = 0
Screen.MousePointer = vbDefault
End Sub

Private Sub TRACK_LIST()
Label1.Caption = "0 - 0 of 0"
CURR_LIST = "0 to 0"
If lvList.ListItems.Count < 1 Then Exit Sub
''Display the page information
With MY_PAGE
        If .PAGE_TOTAL = 1 Then
            Label1.Caption = lvList.SelectedItem.Index & " - " & lvList.ListItems.Count & " of " & RS.RecordCount
            CURR_LIST = lvList.SelectedItem.Index & " to " & lvList.ListItems.Count
        ElseIf .PAGE_CURRENT = 1 Then
            Label1.Caption = lvList.SelectedItem.Index & " - " & lvList.ListItems.Count & " of " & RS.RecordCount
            CURR_LIST = lvList.SelectedItem.Index & " to " & lvList.ListItems.Count
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            Label1.Caption = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index & " - " & RS.RecordCount & " of " & RS.RecordCount
            CURR_LIST = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index & " to " & RS.RecordCount
        Else
            Label1.Caption = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index & " - " & ((.PAGE_CURRENT - 1) * 100) + 100 & " of " & RS.RecordCount
            CURR_LIST = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index & " to " & ((.PAGE_CURRENT - 1) * 100) + 100
        End If
End With
End Sub

''Filter record
Public Sub FILTER_REC()
If RS.RecordCount < 1 Then
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
        RS.Filter = adFilterNone
        RS.Requery
        Call FILL_RECORD(1)
    ''Display record status
    Case 2:
        MsgBox "Selected Record:          " & GET_SELECTED_RECORD_NUM & vbCrLf & _
               "Current List:                 " & CURR_LIST & vbCrLf & vbCrLf & _
               "Total Records:              " & RS.RecordCount, vbInformation
    ''Delete record
    Case 3:
        If RS.RecordCount < 1 Then MsgBox "No record to delete.", vbExclamation: Exit Sub
        Dim ANS As Integer
        ANS = MsgBox("Are you sure you want to delete the selected record?", vbCritical + vbYesNo, "Confirm Record Delete")
        Me.MousePointer = vbHourglass
        If ANS = vbYes Then
            With RS
                .MoveFirst
                .Find "RowIndex = " & Val(lvList.SelectedItem.ListSubItems(10))
                .Delete
                .Requery
                Call FILL_RECORD(1)
                MsgBox "Record has been successfully deleted.", vbInformation, "Confirm"
            End With
        End If
        ANS = 0
        Me.MousePointer = vbDefault
    ''Add New
    Case 4:
        frm_StudentProfile.ADD_STATE = True
        frm_StudentProfile.Show
        Unload Me
    ''Edit
    Case 5:
        If RS.RecordCount < 1 Then MsgBox "No record to edit.", vbExclamation, "Information": Exit Sub
        frm_StudentProfile.SRC_RI = Val(lvList.SelectedItem.ListSubItems(10))
        frm_StudentProfile.ADD_STATE = False
        frm_StudentProfile.Show
        Unload Me
End Select
End Sub

Private Function GET_SELECTED_RECORD_NUM() As Integer
GET_SELECTED_RECORD_NUM = 0
If lvList.ListItems.Count < 1 Then Exit Function
With MY_PAGE
        If .PAGE_TOTAL = 1 Then
            GET_SELECTED_RECORD_NUM = lvList.SelectedItem.Index
        ElseIf .PAGE_CURRENT = 1 Then
            GET_SELECTED_RECORD_NUM = lvList.SelectedItem.Index
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            GET_SELECTED_RECORD_NUM = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index
        Else
            GET_SELECTED_RECORD_NUM = ((.PAGE_CURRENT - 1) * 100) + lvList.SelectedItem.Index
        End If
End With
End Function

Private Sub Form_Unload(Cancel As Integer)
''Cleanup variables
RS.Close
Set RS = Nothing
Set CN = Nothing

CURR_COL = 0
CURR_LIST = ""

frm_Main.RESET_STATUS

frm_Main.RESTORE_BUTTON_VALUE
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

Private Sub lvList_ItemClick(ByVal Item As MSComctlLib.ListItem)

 On Error Resume Next
 cbx_Status.Text = "Barrowed"
 Selected = Val(lvList.SelectedItem.ListSubItems(10))
 RS_BarrowedBooks.Open "SELECT BookID,Title,DateBarrowed,DateReturned,BookSource,Fines FROM qry_BarrowedBooks WHERE tbl_Students.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & False & "", CN, adOpenStatic, adLockOptimistic

 Call FILL_BARROWEDBOOK(1)
 RS_BarrowedBooks.Close

End Sub

Private Sub Picture1_Resize()
Picture2.Left = Picture1.ScaleWidth - Picture2.ScaleWidth
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



Private Sub Picture3_Resize()

    Line1.X2 = Me.ScaleWidth
    lv_Books.Width = Picture3.ScaleWidth
    Picture4.Left = Picture3.ScaleWidth - Picture4.ScaleWidth

End Sub
Private Sub FILL_BARROWEDBOOK(ByVal SRC_PAGE As Long)

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim pos_start2 As Long, pos_end2 As Long
    With MY_PAGE2
    
            .PAGE_CURRENT = 1
            .PAGE_NEXT = 1
            .PAGE_PREVIOUS = 1
            .PAGE_TOTAL = 1
            .PAGE_CURRENT = SRC_PAGE
            .PAGE_TOTAL = REMOVE_DEC("" & (RS_BarrowedBooks.RecordCount / 100))
            If .PAGE_TOTAL > .PAGE_CURRENT Then
                .PAGE_NEXT = .PAGE_CURRENT + 1
            ElseIf .PAGE_CURRENT > 1 Then
                .PAGE_PREVIOUS = .PAGE_CURRENT - 1
            End If
            If .PAGE_TOTAL = 1 Then
                pos_start2 = 1
                pos_end2 = RS_BarrowedBooks.RecordCount
    
            ElseIf .PAGE_CURRENT = 1 Then
                pos_start2 = 1
                pos_end2 = 100
    
            ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
                pos_start2 = ((.PAGE_CURRENT - 1) * 100) + 1
                pos_end2 = RS_BarrowedBooks.RecordCount
    
            Else
                pos_start2 = ((.PAGE_CURRENT - 1) * 100) + 1
                pos_end2 = (.PAGE_NEXT - 1) * 100
    
            End If
            
    End With
    
    
            ''Sort the listview
            If 2 - 1 <> CURR_COL Then
            lv_Books.SortOrder = 0
            Else
            lv_Books.SortOrder = Abs(lv_Books.SortOrder - 1)
            End If
            lv_Books.SortKey = 2 - 1
            
            lv_Books.Sorted = True
            'CURR_COL = 2 - 1
    
    
    Call FillListView(lv_Books, RS_BarrowedBooks, pos_start2, pos_end2, 6, 12, False, False)

    ''Clear variables
    pos_start2 = 0
    pos_end2 = 0
    Screen.MousePointer = vbDefault

End Sub
