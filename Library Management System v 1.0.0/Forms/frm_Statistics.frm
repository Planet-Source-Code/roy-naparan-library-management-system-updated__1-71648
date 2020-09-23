VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frm_Statistics 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Library Statistics"
   ClientHeight    =   7545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10830
   Icon            =   "frm_Statistics.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5820
      Left            =   270
      TabIndex        =   2
      Top             =   990
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   10266
      _Version        =   393216
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Book Statistics"
      TabPicture(0)   =   "frm_Statistics.frx":3482
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "MSChart1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Picture1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Student Statistics"
      TabPicture(1)   =   "frm_Statistics.frx":349E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "MSChart2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Index Statistics"
      TabPicture(2)   =   "frm_Statistics.frx":34BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSChart3"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox Picture1 
         Height          =   600
         Left            =   90
         ScaleHeight     =   540
         ScaleWidth      =   810
         TabIndex        =   7
         Top             =   5175
         Visible         =   0   'False
         Width           =   870
      End
      Begin MSChart20Lib.MSChart MSChart3 
         Height          =   4740
         Left            =   -74685
         OleObjectBlob   =   "frm_Statistics.frx":34D6
         TabIndex        =   5
         Top             =   720
         Width           =   9510
      End
      Begin MSChart20Lib.MSChart MSChart2 
         Height          =   4695
         Left            =   -74640
         OleObjectBlob   =   "frm_Statistics.frx":582C
         TabIndex        =   4
         Top             =   765
         Width           =   9465
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   4830
         Left            =   315
         OleObjectBlob   =   "frm_Statistics.frx":7B80
         TabIndex        =   3
         Top             =   630
         Width           =   9555
      End
   End
   Begin isCoolButton.isButton btn_Print 
      Height          =   330
      Left            =   8100
      TabIndex        =   0
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_Statistics.frx":9ED6
      Style           =   5
      Caption         =   "&Print"
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
      Left            =   9315
      TabIndex        =   1
      Top             =   6930
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_Statistics.frx":9EF2
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "View statistical graph of Books, Students, Index and other library materials."
      Height          =   465
      Left            =   1080
      TabIndex        =   6
      Top             =   360
      Width           =   2985
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_Statistics.frx":9F0E
      Top             =   135
      Width           =   720
   End
End
Attribute VB_Name = "frm_Statistics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btn_Close_Click()

    Unload Me

End Sub

Private Sub btn_Print_Click()

    MSChart1.EditCopy
      DoEvents
       Printer.Print
      Picture1.PaintPicture Clipboard.GetData(), 0, 0
     


End Sub



Private Sub Form_Activate()

frm_Main.toolbar_Menu.Buttons(15).Value = tbrPressed


End Sub

Private Sub Form_Unload(Cancel As Integer)

frm_Main.toolbar_Menu.Buttons(15).Value = tbrUnpressed

End Sub
