VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_BookDetails 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "shpDot"
   ClientHeight    =   10140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEntryID 
      Height          =   330
      Left            =   1485
      Locked          =   -1  'True
      TabIndex        =   35
      Top             =   1845
      Width           =   1995
   End
   Begin VB.ComboBox dc_Edition 
      Height          =   315
      ItemData        =   "frm_BookDetails.frx":0000
      Left            =   1485
      List            =   "frm_BookDetails.frx":001F
      TabIndex        =   32
      Top             =   5355
      Width           =   1995
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   7
      Left            =   1485
      TabIndex        =   11
      Top             =   6120
      Width           =   1995
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   6
      Left            =   1485
      TabIndex        =   10
      Top             =   5715
      Width           =   1995
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   9
      Left            =   1485
      TabIndex        =   9
      Top             =   6930
      Width           =   1590
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   8
      Left            =   1485
      TabIndex        =   8
      Top             =   6525
      Width           =   1590
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000E&
      Caption         =   "Copy Info."
      Height          =   1410
      Left            =   315
      TabIndex        =   7
      Top             =   7560
      Width           =   5910
      Begin VB.TextBox txtEntry 
         Enabled         =   0   'False
         Height          =   330
         Index           =   13
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "0"
         Top             =   765
         Width           =   1590
      End
      Begin VB.TextBox txtEntry 
         Enabled         =   0   'False
         Height          =   330
         Index           =   12
         Left            =   1170
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   765
         Width           =   1590
      End
      Begin VB.TextBox txtEntry 
         Enabled         =   0   'False
         Height          =   330
         Index           =   11
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         Top             =   405
         Width           =   1590
      End
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   10
         Left            =   1170
         TabIndex        =   15
         Top             =   405
         Width           =   1590
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Reserved:"
         Height          =   240
         Left            =   3150
         TabIndex        =   34
         Top             =   855
         Width           =   735
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Available:"
         Height          =   240
         Left            =   180
         TabIndex        =   23
         Top             =   855
         Width           =   960
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Issued:"
         Height          =   240
         Left            =   3150
         TabIndex        =   22
         Top             =   495
         Width           =   735
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Copies:"
         Height          =   240
         Left            =   135
         TabIndex        =   21
         Top             =   495
         Width           =   960
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Caption         =   "Author(s)"
      Height          =   1770
      Left            =   315
      TabIndex        =   6
      Top             =   3285
      Width           =   9420
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   5
         Left            =   1170
         TabIndex        =   14
         Top             =   1125
         Width           =   7980
      End
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   4
         Left            =   1170
         TabIndex        =   13
         Top             =   765
         Width           =   7980
      End
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   3
         Left            =   1170
         TabIndex        =   12
         Top             =   405
         Width           =   7980
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Author 3:"
         Height          =   240
         Left            =   270
         TabIndex        =   20
         Top             =   1170
         Width           =   960
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Author 2:"
         Height          =   240
         Left            =   270
         TabIndex        =   19
         Top             =   810
         Width           =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Author 1:"
         Height          =   240
         Left            =   270
         TabIndex        =   18
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   2
      Left            =   1485
      TabIndex        =   3
      Top             =   2655
      Width           =   8250
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   1
      Left            =   1485
      TabIndex        =   2
      Top             =   2250
      Width           =   8250
   End
   Begin isCoolButton.isButton btn_Save 
      Height          =   330
      Left            =   6480
      TabIndex        =   0
      Top             =   9495
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      Icon            =   "frm_BookDetails.frx":0050
      Style           =   5
      Caption         =   "&Save Changes"
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
   Begin isCoolButton.isButton btn_Cancel 
      Height          =   330
      Left            =   8145
      TabIndex        =   1
      Top             =   9495
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      Icon            =   "frm_BookDetails.frx":006C
      Style           =   5
      Caption         =   "&Cancel"
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
   Begin VB.Label lblFlag 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   31
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Shape shpDot 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   45
      Index           =   1
      Left            =   0
      Top             =   1395
      Width           =   45
   End
   Begin VB.Shape shpDot 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      Height          =   45
      Index           =   0
      Left            =   1395
      Top             =   0
      Width           =   45
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00000000&
      BorderWidth     =   2
      Height          =   1455
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frm_BookDetails.frx":0088
      Top             =   360
      Width           =   720
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F0E0DF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   60
      Left            =   315
      Top             =   9315
      Width           =   9420
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00F0E0DF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   300
      Left            =   315
      Top             =   1140
      Width           =   9465
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   1155
      TabIndex        =   30
      Top             =   630
      Width           =   5895
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "ISBN No."
      Height          =   240
      Left            =   360
      TabIndex        =   29
      Top             =   6210
      Width           =   960
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   240
      Left            =   360
      TabIndex        =   28
      Top             =   5805
      Width           =   960
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Edition:"
      Height          =   240
      Left            =   360
      TabIndex        =   27
      Top             =   5400
      Width           =   960
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Price:"
      Height          =   240
      Left            =   360
      TabIndex        =   26
      Top             =   7020
      Width           =   960
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Pages:"
      Height          =   240
      Left            =   360
      TabIndex        =   25
      Top             =   6615
      Width           =   960
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Book ID:"
      Height          =   240
      Left            =   315
      TabIndex        =   24
      Top             =   1935
      Width           =   960
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Publication:"
      Height          =   240
      Left            =   315
      TabIndex        =   5
      Top             =   2745
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Book Title:"
      Height          =   240
      Left            =   315
      TabIndex        =   4
      Top             =   2340
      Width           =   960
   End
End
Attribute VB_Name = "frm_BookDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Use to check if the for use for adding or editing
Public ADD_STATE As Boolean
''Get the source Row Index (Use only for editing)
Public SRC_RI As Long
Public PK                   As Long
''Recordset for updating
Dim RS As New ADODB.Recordset

Option Explicit

Private Sub btn_Cancel_Click()

    Unload Me
    frm_BooksRecord.Show

End Sub

Private Sub btn_Save_Click()
On Error Resume Next
    ''Check the emportant field if empty
If is_empty(txtEntryID) = True Then Exit Sub
If is_empty(dc_Edition) = True Then Exit Sub

''If edit then we will check if the record is still exist or not
'If ADD_STATE = False Then
'    If rec_exist("tbl_Book", "RowIndex", rs.Fields("RowIndex")) = False Then
'        MsgBox "This record has been removed by other user.", vbExclamation, "Information"
'        Unload Me
'        frm_BooksRecord.Show
'        Exit Sub
'    End If
'End If

'On Error GoTo err
With RS
    If ADD_STATE = True Then .AddNew
        ![BookID] = txtEntryID.Text
        ![Title] = txtEntry(1).Text
        ![Publication] = txtEntry(2).Text
        ![Author1] = txtEntry(3).Text
        ![Author2] = txtEntry(4).Text
        ![Author3] = txtEntry(5).Text
        ![Edition] = dc_Edition.Text
        ![SubjectCategory] = txtEntry(6).Text
        ![ISSBNNo] = txtEntry(7).Text
        ![Pages] = txtEntry(8).Text
        ![Price] = txtEntry(9).Text
        ![TotalNo] = txtEntry(10).Text
        ![IssuedNo] = txtEntry(11).Text
        ![AvailNo] = txtEntry(10).Text
        ![DateModified] = Now
    .Update
End With

If ADD_STATE = True Then
    MsgBox "New record has been successfully saved.", vbInformation, "Information"
    Dim Reply As Integer
    Reply = MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo, "Confirmation")
    If Reply = vbYes Then
        Call RESET_FIELD
        txtEntry(11).Text = 0
        GeneratePK
    Else
        Unload Me
        frm_BooksRecord.Show
    End If
    Reply = 0
Else:
    MsgBox "Changes record has been successfully saved.", vbInformation, "Information"
    Unload Me
    frm_BooksRecord.Show
End If


Exit Sub
err:
    Call prompt_err(err)

End Sub

Private Sub Form_Activate()
frm_Main.toolbar_Menu.Buttons(2).Value = tbrPressed
''Display info in the status bar
If ADD_STATE = True Then
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Book Details - Create New Record"
Else
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Book Details - Edit Existing Record"
End If
End Sub

Private Sub Form_Load()
Me.Top = 100
Call centerFormHorizontal(Me, Screen.Width)


If ADD_STATE = True Then
    lblFlag.Caption = "Create New Record"
    Me.Caption = "Create New Record"
    RS.Open "SELECT * FROM tbl_Book", CN, adOpenStatic, adLockOptimistic
    GeneratePK
    txtEntry(10).Locked = False
Else
    lblFlag.Caption = "Edit Existing Record"
    Me.Caption = "Edit Existing Record"
    RS.Open "SELECT * FROM tbl_Book WHERE RowIndex =" & SRC_RI, CN, adOpenStatic, adLockOptimistic
    Call FILL_FIELDS
    txtEntry(10).Locked = True
End If

End Sub

Private Sub Form_Resize()
shpBorder(0).Width = Me.ScaleWidth
shpBorder(0).Height = Me.ScaleHeight

shpBorder(1).Width = Me.ScaleWidth
shpBorder(1).Height = Me.ScaleHeight

shpDot(0).Left = Me.ScaleWidth - shpDot(0).Width + 20
shpDot(1).Top = Me.ScaleHeight - shpDot(1).Height + 20
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    RS.Close
    Set RS = Nothing


    frm_Main.RESET_STATUS

    frm_Main.RESTORE_BUTTON_VALUE

End Sub

Private Sub RESET_FIELD()
''Clear the entry fields
Call clearText(Me)
dc_Edition.Text = ""
txtEntryID.SetFocus
End Sub

Private Sub FILL_FIELDS()
On Error GoTo err
''Display records from database
With RS
        txtEntryID.Text = ![BookID]
        txtEntry(1).Text = ![Title]
        txtEntry(2).Text = ![Publication]
        txtEntry(3).Text = ![Author1]
        txtEntry(4).Text = ![Author2]
        txtEntry(5).Text = ![Author3]
        dc_Edition.Text = ![Edition]
        txtEntry(6).Text = ![SubjectCategory]
        txtEntry(7).Text = ![ISSBNNo]
        txtEntry(8).Text = ![Pages]
        txtEntry(9).Text = ![Price]
        txtEntry(10).Text = ![TotalNo]
        txtEntry(11).Text = ![IssuedNo]
        txtEntry(12).Text = ![AvailNo]

End With
Exit Sub
err:
    If err.Number = 94 Then
        ''Error when encounter null value
        Resume Next
    Else
        Call prompt_err(err)
    End If
End Sub
Private Sub GeneratePK()
    PK = getIndex("tbl_Book")
    txtEntryID.Text = GenerateID(PK, "BID-", "00000")
End Sub


Private Sub txtEntry_Change(Index As Integer)

    If ADD_STATE = True Then
    
        txtEntry(12).Text = txtEntry(10).Text
        
    End If

End Sub
