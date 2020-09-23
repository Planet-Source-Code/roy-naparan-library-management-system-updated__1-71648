VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_Index 
   ClientHeight    =   10005
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   12690
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10005
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
      Top             =   7200
      Width           =   12690
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   330
         Left            =   8280
         ScaleHeight     =   330
         ScaleWidth      =   4395
         TabIndex        =   16
         Top             =   90
         Width           =   4395
         Begin isCoolButton.isButton btn_hide 
            Height          =   330
            Left            =   3960
            TabIndex        =   17
            Top             =   0
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":0000
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
            TabIndex        =   18
            Top             =   0
            Width           =   420
            _ExtentX        =   741
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":0A12
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
      Begin VB.ComboBox cbx_Status 
         Height          =   315
         ItemData        =   "frm_Index.frx":1424
         Left            =   1575
         List            =   "frm_Index.frx":1431
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   90
         Width           =   2040
      End
      Begin MSComctlLib.ListView lv_Books 
         Height          =   2265
         Left            =   45
         TabIndex        =   19
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Student no."
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
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
      End
      Begin isCoolButton.isButton btn_Print 
         Height          =   330
         Left            =   3645
         TabIndex        =   21
         Top             =   90
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":1453
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
      Begin VB.Label Label2 
         Caption         =   "List of Students:"
         Height          =   240
         Left            =   360
         TabIndex        =   20
         Top             =   135
         Width           =   1590
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   90
         Picture         =   "frm_Index.frx":1E65
         Top             =   90
         Width           =   240
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   45
         X2              =   12330
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   6225
      Left            =   -45
      TabIndex        =   1
      Top             =   -45
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   10980
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
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Index ID"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Subject"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Title"
         Object.Width           =   8819
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Author"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Publisher"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Edition"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Page"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Date Modified"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "RowIndex"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   0
      ScaleHeight     =   510
      ScaleWidth      =   12690
      TabIndex        =   0
      Top             =   6690
      Width           =   12690
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   8325
         ScaleHeight     =   465
         ScaleWidth      =   4305
         TabIndex        =   6
         Top             =   0
         Width           =   4305
         Begin isCoolButton.isButton command3 
            Height          =   330
            Left            =   1980
            TabIndex        =   8
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":2867
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
            TabIndex        =   9
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":3279
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
            TabIndex        =   10
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":3C8B
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
            TabIndex        =   11
            Top             =   90
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   582
            Icon            =   "frm_Index.frx":469D
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
            TabIndex        =   12
            Top             =   180
            Width           =   1545
         End
      End
      Begin isCoolButton.isButton btn_Add 
         Height          =   330
         Left            =   45
         TabIndex        =   2
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":50AF
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
         TabIndex        =   3
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":50CB
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
         TabIndex        =   4
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":50E7
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
         TabIndex        =   5
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":5103
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
         TabIndex        =   7
         Top             =   90
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   582
         Icon            =   "frm_Index.frx":511F
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
         Icon            =   "frm_Index.frx":513B
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
End
Attribute VB_Name = "frm_Index"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Recordset used to the display the record
Dim RS As New ADODB.Recordset

Dim RS_BarrowedIndex As New ADODB.Recordset

Dim RS_FindIndexID As New ADODB.Recordset
Dim RS_rptListofStudentBarrowedIndex As New ADODB.Recordset

Dim IndexID As String
Dim Title As String
Dim Subject As String
Dim Author As String
Dim Publisher As String
Dim Edition As String

Dim Selected As Integer
''Variable used to page the records
Dim MY_PAGE                         As PAGE_INFO
Dim MY_PAGE2                        As PAGE_INFO

''Variable that hold the current column (Used to sorting the List View)
Dim CURR_COL                        As Integer
''Variable the current list in the page
Dim CURR_LIST                       As String

Option Explicit

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
    CN17.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    CN18.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
    
    
    IndexID = lvList.SelectedItem.ListSubItems(8)
    Title = lvList.SelectedItem.ListSubItems(2)
    Subject = lvList.SelectedItem.ListSubItems(1)
    Author = lvList.SelectedItem.ListSubItems(3)
    Publisher = lvList.SelectedItem.ListSubItems(4)
    Edition = lvList.SelectedItem.ListSubItems(5)

    
        RS_FindIndexID.Open "SELECT IndexID FROM tbl_Index WHERE RowIndex LIKE '" & IndexID & "'", CN17, adOpenStatic, adLockOptimistic
        
    
        rpt_IndexTransaction.Sections("Section2").Controls("lblIndexID").Caption = RS_FindIndexID.Fields("IndexID")
        rpt_IndexTransaction.Sections("Section2").Controls("lblTitle").Caption = Title
        rpt_IndexTransaction.Sections("Section2").Controls("lblSubject").Caption = Subject
        rpt_IndexTransaction.Sections("Section2").Controls("lblAuthor").Caption = Author
        rpt_IndexTransaction.Sections("Section2").Controls("lblPublisher").Caption = Publisher
        rpt_IndexTransaction.Sections("Section2").Controls("lblEdition").Caption = Edition

    
        CN17.Close
    
    If cbx_Status.Text = "Barrowed" Then
 
        RS_rptListofStudentBarrowedIndex.Open "SELECT * FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "' AND  StatusReturned =" & False & "", CN18, adOpenStatic, adLockOptimistic
        
        Set rpt_IndexTransaction.DataSource = RS_rptListofStudentBarrowedIndex
        
        rpt_IndexTransaction.Show 1
        RS_rptListofStudentBarrowedIndex.Close
        
        
    ElseIf cbx_Status.Text = "Returned" Then
    
        RS_rptListofStudentBarrowedIndex.Open "SELECT * FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & True & "", CN18, adOpenStatic, adLockOptimistic
        
        Set rpt_IndexTransaction.DataSource = RS_rptListofStudentBarrowedIndex
        
        rpt_IndexTransaction.Show 1
        RS_rptListofStudentBarrowedIndex.Close
        
    ElseIf cbx_Status.Text = "View All" Then
    
        RS_rptListofStudentBarrowedIndex.Open "SELECT * FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "'", CN18, adOpenStatic, adLockOptimistic
        
        Set rpt_IndexTransaction.DataSource = RS_rptListofStudentBarrowedIndex
        
        rpt_IndexTransaction.Show 1
        RS_rptListofStudentBarrowedIndex.Close
        
    End If

End Sub

Private Sub btn_Reload_Click()

    COMMAND_PASS (1)

End Sub

Private Sub btn_Search_Click()

    With frm_SearchIndexRecords
    
        Set .SRC_FORM = Me
        Set .RS = RS
        .eFilter = "%'"
        .Show vbModal
        
    End With

End Sub

Private Sub btn_Show_Click()

    Picture3.Height = 2805
    lvList.Height = (Me.ScaleHeight - (Picture1.ScaleHeight + 30 + Picture3.ScaleHeight))

End Sub

Private Sub cbx_Status_Click()

     If cbx_Status.Text = "Barrowed" Then

        RS_BarrowedIndex.Open "SELECT StudentNo,StudentName,DateBarrowed,DateReturned FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & False & "", CN, adOpenStatic, adLockOptimistic
 
        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedIndex.Close

    ElseIf cbx_Status.Text = "Returned" Then

        RS_BarrowedIndex.Open "SELECT StudentNo,StudentName,DateBarrowed,DateReturned FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & True & "", CN, adOpenStatic, adLockOptimistic

        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedIndex.Close

    ElseIf cbx_Status.Text = "View All" Then

        RS_BarrowedIndex.Open "SELECT StudentNo,StudentName,DateBarrowed,DateReturned FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "'", CN, adOpenStatic, adLockOptimistic

        Call FILL_BARROWEDBOOK(1)
        RS_BarrowedIndex.Close

    End If


End Sub

Private Sub Form_Activate()

    frm_Main.toolbar_Menu.Buttons(5).Value = tbrPressed

    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Index"

    Kill (App.Path & "\ActiveForm.dat")
    Call savePath
    
End Sub

Private Sub savePath()
    On Error Resume Next
    Open App.Path & "\ActiveForm.dat" For Output As #1
        Print #1, frm_Main.StatusBar1.Panels(2).Text
    Close #1
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

lvList.ColumnHeaders(1).Alignment = lvwColumnLeft
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
RS.Open "SELECT * FROM tbl_Index ORDER BY IndexID ASC", CN, adOpenStatic, adLockOptimistic
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
           ' CURR_COL = 2 - 1

Call FillListView(lvList, RS, pos_start, pos_end, 9, 8, False, False)
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
If RS.RecordCount < 1 Then MsgBox "No record match.", vbInformation, "Search Result"
Call FILL_RECORD(1)
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
                .Find "RowIndex = " & Val(lvList.SelectedItem.ListSubItems(8))
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
        frm_IndexDetails.ADD_STATE = True
        frm_IndexDetails.Show
        Unload Me
    ''Edit
    Case 5:
        If RS.RecordCount < 1 Then MsgBox "No record to edit.", vbExclamation, "Information": Exit Sub
        frm_IndexDetails.SRC_RI = Val(lvList.SelectedItem.ListSubItems(8))
        frm_IndexDetails.ADD_STATE = False
        frm_IndexDetails.Show
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

 'On Error Resume Next
 cbx_Status.Text = "Barrowed"
 Selected = Val(lvList.SelectedItem.ListSubItems(8))
 RS_BarrowedIndex.Open "SELECT StudentNo,StudentName,DateBarrowed,DateReturned FROM qry_BarrowedIndex WHERE tbl_Index.RowIndex LIKE '" & Selected & "' AND StatusReturned =" & False & "", CN, adOpenStatic, adLockOptimistic

 Call FILL_BARROWEDBOOK(1)
 RS_BarrowedIndex.Close

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
Private Sub FILL_BARROWEDBOOK(ByVal SRC_PAGE2 As Long)

    On Error Resume Next
    Screen.MousePointer = vbHourglass
    Dim pos_start2 As Long, pos_end2 As Long
    With MY_PAGE2
    
            .PAGE_CURRENT = 1
            .PAGE_NEXT = 1
            .PAGE_PREVIOUS = 1
            .PAGE_TOTAL = 1
            .PAGE_CURRENT = SRC_PAGE2
            .PAGE_TOTAL = REMOVE_DEC("" & (RS_BarrowedIndex.RecordCount / 100))
            If .PAGE_TOTAL > .PAGE_CURRENT Then
                .PAGE_NEXT = .PAGE_CURRENT + 1
            ElseIf .PAGE_CURRENT > 1 Then
                .PAGE_PREVIOUS = .PAGE_CURRENT - 1
            End If
            If .PAGE_TOTAL = 1 Then
                pos_start2 = 1
                pos_end2 = RS_BarrowedIndex.RecordCount
    
            ElseIf .PAGE_CURRENT = 1 Then
                pos_start2 = 1
                pos_end2 = 100
    
            ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
                pos_start2 = ((.PAGE_CURRENT - 1) * 100) + 1
                pos_end2 = RS_BarrowedIndex.RecordCount
    
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
    
    
    Call FillListView(lv_Books, RS_BarrowedIndex, pos_start2, pos_end2, 4, 6, False, False)

    ''Clear variables
    pos_start2 = 0
    pos_end2 = 0
    Screen.MousePointer = vbDefault

End Sub

