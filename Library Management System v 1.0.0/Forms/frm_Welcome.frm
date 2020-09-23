VERSION 5.00
Begin VB.Form frm_Welcome 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Login Time:"
   ClientHeight    =   1590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   2730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   1500
      Left            =   45
      ScaleHeight     =   1470
      ScaleWidth      =   2595
      TabIndex        =   0
      Top             =   45
      Width           =   2625
      Begin VB.Label lbl_User 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   1350
         TabIndex        =   7
         Top             =   495
         Width           =   1365
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Login User:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   495
         Width           =   1230
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Login Date:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   135
         TabIndex        =   5
         Top             =   1125
         Width           =   1005
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Login Time:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   135
         TabIndex        =   4
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "USER LOG"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   240
         Left            =   360
         TabIndex        =   3
         Top             =   90
         Width           =   2040
      End
      Begin VB.Label lbl_Time 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1350
         TabIndex        =   2
         Top             =   810
         Width           =   1365
      End
      Begin VB.Label lbl_Day 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   285
         Left            =   1350
         TabIndex        =   1
         Top             =   1125
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frm_Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2
Const SWP_NOACTIVATE = &H10
Const SWP_SHOWWINDOW = &H40

Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim i As Integer

Option Explicit
Private Sub Form_Load()

    On Error Resume Next
    Me.Left = Screen.Width - (Me.Width + 50)
    Me.Top = Screen.Height - 450
    Picture1.Visible = False
    MakeTransparent Me.hwnd, 180

End Sub


Private Sub popup()
On Error Resume Next
    Picture1.Visible = True
    i = Me.Height
    Me.Height = 0
    While Me.Height < i
        Me.Height = Me.Height + 2
        Me.Top = Me.Top - 2
        DoEvents
    Wend
End Sub
Private Sub popdown()
On Error Resume Next
    i = Me.Height
    While Me.Height > 500
        Me.Height = Me.Height - 2
        Me.Top = Me.Top + 2
        DoEvents
    Wend
End Sub
Private Sub Form_Activate()
On Error Resume Next
    
    frm_Main.Enabled = False
    lbl_User.Caption = frm_Main.StatusBar1.Panels(4)
    lbl_Time.Caption = Format$(Now, "hh:mm:ss AM/PM")
    lbl_Day.Caption = Format$(Date, "dd-MMM-yyyy")
    Call popup
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
    Sleep welcometime
    Call popdown
frm_Main.Enabled = True
Unload Me
End Sub

