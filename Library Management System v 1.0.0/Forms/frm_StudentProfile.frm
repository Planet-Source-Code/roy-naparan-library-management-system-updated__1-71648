VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_StudentProfile 
   BackColor       =   &H8000000E&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10140
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   10140
   Begin MSComDlg.CommonDialog cmddlgBrowse 
      Left            =   7515
      Top             =   1845
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtEntry 
      Height          =   1905
      Index           =   5
      Left            =   405
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   6525
      Width           =   6270
   End
   Begin VB.ComboBox cbx_Status 
      Height          =   315
      ItemData        =   "frm_StudentProfile.frx":0000
      Left            =   1665
      List            =   "frm_StudentProfile.frx":000D
      TabIndex        =   21
      Top             =   5490
      Width           =   2445
   End
   Begin VB.ComboBox cbx_Level 
      Height          =   315
      ItemData        =   "frm_StudentProfile.frx":0027
      Left            =   1665
      List            =   "frm_StudentProfile.frx":003D
      TabIndex        =   20
      Top             =   5130
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo dc_Course 
      Height          =   315
      Left            =   1665
      TabIndex        =   19
      Top             =   4770
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   4
      Left            =   1665
      TabIndex        =   18
      Top             =   4095
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   3
      Left            =   1665
      TabIndex        =   17
      Top             =   3690
      Width           =   3030
   End
   Begin VB.TextBox txtEntry 
      Height          =   690
      Index           =   2
      Left            =   1665
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2925
      Width           =   4920
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   1
      Left            =   1665
      TabIndex        =   6
      Top             =   2250
      Width           =   4920
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   0
      Left            =   1665
      TabIndex        =   2
      Top             =   1845
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_Save 
      Height          =   330
      Left            =   6480
      TabIndex        =   7
      Top             =   9180
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      Icon            =   "frm_StudentProfile.frx":0060
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
      TabIndex        =   8
      Top             =   9180
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   582
      Icon            =   "frm_StudentProfile.frx":007C
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
   Begin isCoolButton.isButton cmdChangePic 
      Height          =   330
      Left            =   8010
      TabIndex        =   23
      Top             =   3780
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      Icon            =   "frm_StudentProfile.frx":0098
      Style           =   5
      Caption         =   "&Change Image"
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
   Begin isCoolButton.isButton btn_Remove 
      Height          =   330
      Left            =   8010
      TabIndex        =   24
      Top             =   4140
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   582
      Icon            =   "frm_StudentProfile.frx":00B4
      Style           =   5
      Caption         =   "&Remove Image"
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
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   8055
      Picture         =   "frm_StudentProfile.frx":00D0
      Stretch         =   -1  'True
      Top             =   4860
      Visible         =   0   'False
      Width           =   1650
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
   Begin VB.Image imgStudent 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   8055
      Picture         =   "frm_StudentProfile.frx":3A1D
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1650
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00F0E0DF&
      BorderWidth     =   2
      Height          =   1875
      Left            =   8010
      Top             =   1845
      Width           =   1740
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Comment (Optional):"
      Height          =   240
      Left            =   405
      TabIndex        =   15
      Top             =   6255
      Width           =   1545
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Status:"
      Height          =   240
      Left            =   405
      TabIndex        =   14
      Top             =   5535
      Width           =   1140
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Year:"
      Height          =   240
      Left            =   405
      TabIndex        =   13
      Top             =   5175
      Width           =   1140
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Course:"
      Height          =   240
      Left            =   405
      TabIndex        =   12
      Top             =   4815
      Width           =   1140
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No.:"
      Height          =   240
      Left            =   405
      TabIndex        =   11
      Top             =   4185
      Width           =   1140
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   240
      Left            =   405
      TabIndex        =   10
      Top             =   3735
      Width           =   1140
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
      Height          =   240
      Left            =   405
      TabIndex        =   9
      Top             =   2925
      Width           =   1140
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00F0E0DF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   60
      Left            =   405
      Top             =   9000
      Width           =   9330
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "(Last Name, First Name Middle Initial)"
      Height          =   240
      Left            =   1665
      TabIndex        =   5
      Top             =   2655
      Width           =   3705
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name:"
      Height          =   240
      Left            =   360
      TabIndex        =   4
      Top             =   2385
      Width           =   1140
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Student No.:"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   1935
      Width           =   960
   End
   Begin VB.Shape shpBorder 
      BorderColor     =   &H00000000&
      Height          =   1455
      Index           =   0
      Left            =   0
      Top             =   0
      Width           =   1455
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
      Left            =   405
      TabIndex        =   1
      Top             =   1170
      Width           =   1815
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Profile"
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
      TabIndex        =   0
      Top             =   630
      Width           =   5895
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
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frm_StudentProfile.frx":736A
      Top             =   360
      Width           =   720
   End
End
Attribute VB_Name = "frm_StudentProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Use to check if the for use for adding or editing
Public ADD_STATE As Boolean
''Get the source Row Index (Use only for editing)
Public SRC_RI As Long
Dim OLD_NAME As String
''Recordset for updating
Public RS As New ADODB.Recordset

Dim IMG_PATH As String
Option Explicit

Private Sub btn_Cancel_Click()

    Unload Me
    frm_StudentRecords.Show

End Sub

Private Sub btn_Remove_Click()

    If MsgBox("Are you do you want to remove this image?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
    
        RS.Fields("HavePic") = 0
        RS.Fields("Picture") = ""
        imgStudent.Picture = Image2.Picture
        
    End If

End Sub

Private Sub btn_Save_Click()

    ''Check the emportant field if empty
'If is_empty(txtEntry(0)) = True Then Exit Sub

''If edit then we will check if the record is still exist or not
'If ADD_STATE = False Then
'    If rec_exist("tbl_Students", "RowIndex", RS.Fields("RowIndex")) = False Then
'        MsgBox "This record has been removed by other user.", vbExclamation
'        Unload Me
'        Exit Sub
'    End If
'    If LCase$(OLD_NAME) <> LCase$(txtEntry(0).Text) Then
'        If rec_exist("tbl_Students", "StudentName", txtEntry(0).Text, True) = True Then
'            MsgBox "Student name already exist in the records.Please try to have another Student name.", vbExclamation
'            txtEntry(0).SetFocus
'            Exit Sub
'        End If
'    End If
'Else
'    If rec_exist("tbl_Students", "StudentName", txtEntry(0).Text, True) = True Then
'        MsgBox "Student name already exist in the records.Please try to have another Student name.", vbExclamation
'        txtEntry(0).SetFocus
'        Exit Sub
'    End If
'
'End If

On Error Resume Next
With RS
    If ADD_STATE = True Then .AddNew
    
        ![StudentNo] = txtEntry(0).Text
        ![StudentName] = txtEntry(1).Text
        ![CurrentAddress] = txtEntry(2).Text
        ![E-mail] = txtEntry(3).Text
        ![ContactNo] = txtEntry(4).Text
        ![Course] = dc_Course.Text
        ![Level] = cbx_Level.Text
        ![Status] = cbx_Status.Text
        ![Comment] = txtEntry(5).Text
        ![DateModified] = Now

        If ADD_STATE = True Then
            If IMG_PATH = "" Then
                .Fields("HavePic") = 0
                .Fields("Picture") = ""
            Else
                .Fields("HavePic") = 1
                Call save_file_to_db(RS, "Picture", IMG_PATH)
            End If
        ElseIf IMG_PATH <> "NONE" Then
            If IMG_PATH <> "" Then
                .Fields("HavePic") = 1
                Call save_file_to_db(RS, "Picture", IMG_PATH)
            Else
                .Fields("HavePic") = 0
                .Fields("Picture") = ""
            End If
        End If
    .Update
End With

If ADD_STATE = True Then
    MsgBox "New record has been successfully saved.", vbInformation, "Information"
    Dim Reply As Integer
    Reply = MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo, "Confirmation")
    If Reply = vbYes Then
        Call RESET_FIELD
        dc_Course.Text = ""
        cbx_Level.Text = ""
        cbx_Status.Text = ""
    Else
        Unload Me
        frm_StudentRecords.Show
    End If
    Reply = 0
Else
    MsgBox "Changes record has been successfully saved.", vbInformation, "Information"
    Unload Me
    frm_StudentRecords.Show
End If


Exit Sub
'ERR:
'    Call prompt_err(ERR)

End Sub

Private Sub cmdChangePic_Click()

    cmddlgBrowse.ShowOpen
    If cmddlgBrowse.FileName <> "" Then
    IMG_PATH = cmddlgBrowse.FileName
    imgStudent.Picture = LoadPicture(IMG_PATH)
End If

End Sub

Private Sub Form_Activate()

frm_Main.toolbar_Menu.Buttons(5).Value = tbrPressed
''Display info in the status bar
If ADD_STATE = True Then
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Student Profile - Create New Record"
Else
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Student Profile - Edit Existing Record"
End If

End Sub


Private Sub Form_Load()

Me.Top = 100
Call centerFormHorizontal(Me, Screen.Width)

Dim rsClientID As New Recordset
        rsClientID.Open "SELECT * FROM tbl_Course", CN, adOpenStatic, adLockOptimistic
        
            Set dc_Course.RowSource = rsClientID
                dc_Course.ListField = "CourseName"
                dc_Course.BoundColumn = "CourseID"


'With frm_Main
'    imgStudent.Picture = .img_Pic.ListImages(1).Picture
'End With

cmddlgBrowse.Filter = "Graphics format (*.bmp,*.jpg,*.gif)|*.bmp;*.jpg;*.gif;*.jpeg"

If ADD_STATE = True Then
    lblFlag.Caption = "Create New Record"
    Me.Caption = "Create New Record"
    RS.Open "SELECT * FROM tbl_Students", CN, adOpenStatic, adLockOptimistic
Else
    lblFlag.Caption = "Edit Existing Record"
    Me.Caption = "Edit Existing Record"
    RS.Open "SELECT * FROM tbl_Students WHERE RowIndex =" & SRC_RI, CN, adOpenStatic, adLockOptimistic
    Call FILL_FIELDS
End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

    RS.Close
    Set RS = Nothing


    frm_Main.RESET_STATUS

    frm_Main.RESTORE_BUTTON_VALUE

End Sub
Private Sub Form_Resize()
shpBorder(0).Width = Me.ScaleWidth
shpBorder(0).Height = Me.ScaleHeight

shpBorder(1).Width = Me.ScaleWidth
shpBorder(1).Height = Me.ScaleHeight

shpDot(0).Left = Me.ScaleWidth - shpDot(0).Width + 20
shpDot(1).Top = Me.ScaleHeight - shpDot(1).Height + 20
End Sub

Private Sub FILL_FIELDS()
On Error GoTo err
''Display records from database
With RS
        txtEntry(0).Text = ![StudentNo]
        txtEntry(1).Text = ![StudentName]
        txtEntry(2).Text = ![CurrentAddress]
        txtEntry(3).Text = ![E-mail]
        txtEntry(4).Text = ![ContactNo]
        dc_Course.Text = ![Course]
        cbx_Level.Text = ![Level]
        cbx_Status.Text = ![Status]
        txtEntry(5).Text = ![Comment]
      

        If .Fields("HavePic") = 1 Then
            ''Assign NONE to variable to unchange the picture
            IMG_PATH = "NONE"
            ''Bind the image control to get the picture
            imgStudent.DataField = "Picture"
            Set imgStudent.DataSource = RS
            ''Unbind the image control
            Set imgStudent.DataSource = Nothing
            imgStudent.DataField = ""
        End If
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

Private Sub RESET_FIELD()
''Clear the entry fields
Call clearText(Me)


''Reset the pirture
cmdRemPic_Click

''Set the focus to the Ship name
txtEntry(0).SetFocus
End Sub

Private Sub cmdRemPic_Click()
IMG_PATH = ""
imgStudent.Picture = Image2.Picture
End Sub

