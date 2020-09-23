VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_UserAddEdit 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5775
   Icon            =   "frm_UserAddEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   5775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4500
      Top             =   405
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserAddEdit.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_UserAddEdit.frx":6914
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "User Information"
      Height          =   2985
      Left            =   315
      TabIndex        =   6
      Top             =   2880
      Width           =   5145
      Begin VB.ComboBox cmb_Usertype 
         Height          =   315
         ItemData        =   "frm_UserAddEdit.frx":9DA6
         Left            =   1665
         List            =   "frm_UserAddEdit.frx":9DB3
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2295
         Width           =   3255
      End
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   5
         Left            =   1665
         TabIndex        =   12
         Top             =   1890
         Width           =   3255
      End
      Begin VB.TextBox txtEntry 
         Height          =   915
         Index           =   4
         Left            =   1665
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   900
         Width           =   3255
      End
      Begin VB.TextBox txtEntry 
         Height          =   330
         Index           =   3
         Left            =   1665
         TabIndex        =   7
         Top             =   495
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "User Type:"
         Height          =   240
         Left            =   180
         TabIndex        =   14
         Top             =   2385
         Width           =   960
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Contact No.:"
         Height          =   240
         Left            =   180
         TabIndex        =   13
         Top             =   1980
         Width           =   960
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Address:"
         Height          =   240
         Left            =   180
         TabIndex        =   10
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Full Name:"
         Height          =   240
         Left            =   180
         TabIndex        =   8
         Top             =   540
         Width           =   960
      End
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   1980
      PasswordChar    =   "•"
      TabIndex        =   4
      Top             =   2250
      Width           =   3255
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1980
      PasswordChar    =   "•"
      TabIndex        =   2
      Top             =   1845
      Width           =   3255
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   0
      Left            =   1980
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
   Begin isCoolButton.isButton btn_Add 
      Height          =   345
      Left            =   3060
      TabIndex        =   15
      Top             =   6075
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Icon            =   "frm_UserAddEdit.frx":9DE2
      Style           =   5
      Caption         =   "&Save"
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
      Height          =   345
      Left            =   4275
      TabIndex        =   16
      Top             =   6075
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   609
      Icon            =   "frm_UserAddEdit.frx":9DFE
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   270
      X2              =   5445
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   270
      X2              =   5445
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Add new user in the system by providing correct information."
      Height          =   510
      Left            =   1125
      TabIndex        =   11
      Top             =   495
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   270
      Top             =   180
      Width           =   825
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Re-type Password:"
      Height          =   240
      Left            =   360
      TabIndex        =   5
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   1935
      Width           =   960
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   1530
      Width           =   960
   End
End
Attribute VB_Name = "frm_UserAddEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Use to check if the for use for adding or editing
Public ADD_STATE As Boolean
''Get the source Row Index (Use only for editing)
Public SRC_RI As Long

''Recordset for updating
Public RS As New ADODB.Recordset

Option Explicit

Private Sub btn_Add_Click()

    On Error Resume Next
    
    If txtEntry(2).Text = txtEntry(1).Text Then
    
            With RS
                If ADD_STATE = True Then .AddNew
                
                    ![UserName] = txtEntry(0).Text
                    ![Password] = Encode(txtEntry(2).Text)
                    ![FullName] = txtEntry(3).Text
                    ![Address] = txtEntry(4).Text
                    ![ContactNo] = txtEntry(5).Text
                    ![UserType] = cmb_Usertype
                    
                .Update
                
            End With
            
            If ADD_STATE = True Then
                
                    MsgBox "New record has been successfully saved.", vbInformation, "Information"
                    Dim Reply As Integer
                    Reply = MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo, "Confirmation")
                    If Reply = vbYes Then
                        Call RESET_FIELD
                        cmb_Usertype.Text = ""
                    Else
                        Unload Me
                        frm_UserManagement.Show
                    End If
                    Reply = 0
                             
            Else
            
                    MsgBox "Changes of record has been successfully saved.", vbInformation, "Information"
                    Unload Me
                    frm_UserManagement.Show
                
            End If
    
    
    Else
        
            MsgBox "Password did not match. Please type it again.", vbInformation, "Confirmation"
            
    End If

End Sub

Private Sub btn_Close_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    If ADD_STATE = True Then
        Image1.Picture = frm_UserAddEdit.ImageList1.ListImages(1).Picture
        Me.Caption = "Add New User"
        RS.Open "SELECT * FROM tbl_User", CN2, adOpenStatic, adLockOptimistic
    Else
        Image1.Picture = frm_UserAddEdit.ImageList1.ListImages(2).Picture
        Label5.Caption = "Updating existing user in the system by providing correct information."
        Me.Caption = "Edit Existing User"
        RS.Open "SELECT * FROM tbl_User WHERE RowIndex =" & SRC_RI, CN2, adOpenStatic, adLockOptimistic
        Call FILL_FIELDS
    End If

End Sub
Private Sub FILL_FIELDS()
On Error GoTo err
''Display records from database
With RS
        txtEntry(0).Text = ![UserName]
        txtEntry(1).Text = DeCode(![Password])
        txtEntry(3).Text = ![FullName]
        txtEntry(4).Text = ![Address]
        txtEntry(5).Text = ![ContactNo]
        cmb_Usertype = ![UserType]
        
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

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next

    RS.Close
    Set RS = Nothing


End Sub
Private Sub RESET_FIELD()

''Clear the entry fields
Call clearText(Me)

txtEntry(0).SetFocus

End Sub
