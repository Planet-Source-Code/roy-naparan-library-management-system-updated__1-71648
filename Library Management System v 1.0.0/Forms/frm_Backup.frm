VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_Backup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Back-up Database"
   ClientHeight    =   6210
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   6660
   Icon            =   "frm_Backup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6210
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Previous Back-up"
      Height          =   1770
      Left            =   270
      TabIndex        =   1
      Top             =   4140
      Width           =   6090
      Begin VB.Label lbl_Date 
         BackStyle       =   0  'Transparent
         Caption         =   "Date   :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lbl_time 
         BackStyle       =   0  'Transparent
         Caption         =   "Time   :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lbl_apath 
         BackStyle       =   0  'Transparent
         Caption         =   "Path   :"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lbl_lastdate 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup date"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   990
         TabIndex        =   8
         Top             =   360
         Width           =   3105
      End
      Begin VB.Label lbl_lasttime 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup time"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   990
         TabIndex        =   7
         Top             =   720
         Width           =   3105
      End
      Begin VB.Label lbl_path 
         BackStyle       =   0  'Transparent
         Caption         =   "Last backup path"
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   990
         TabIndex        =   6
         Top             =   1080
         Width           =   4845
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current Back-up"
      Height          =   2850
      Left            =   270
      TabIndex        =   0
      Top             =   1170
      Width           =   6090
      Begin VB.DirListBox Dir1 
         Height          =   1665
         Left            =   360
         TabIndex        =   13
         ToolTipText     =   "Select folder"
         Top             =   900
         Width           =   3825
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   360
         TabIndex        =   12
         ToolTipText     =   "Select diskdrive"
         Top             =   540
         Width           =   3825
      End
      Begin VB.FileListBox File1 
         Height          =   1650
         Left            =   135
         Pattern         =   "*.mdb"
         TabIndex        =   3
         Top             =   900
         Width           =   255
      End
      Begin isCoolButton.isButton btn_Save 
         Height          =   330
         Left            =   4365
         TabIndex        =   4
         Top             =   1845
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         Icon            =   "frm_Backup.frx":1982
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
      Begin isCoolButton.isButton btn_Cancel 
         Height          =   330
         Left            =   4365
         TabIndex        =   5
         Top             =   2205
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         Icon            =   "frm_Backup.frx":199E
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Make back-up on database for reference and to secure data from loss."
      Height          =   600
      Left            =   1125
      TabIndex        =   2
      Top             =   360
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_Backup.frx":19BA
      Top             =   225
      Width           =   720
   End
End
Attribute VB_Name = "frm_Backup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim FileSys As New FileSystemObject
Dim bckupFile As File

Option Explicit

Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Save_Click()

    On Error Resume Next

    Dim finalpath As String
    Dim destination As String
    Dim Source As String
    Dim currDate, currTime As String
    
    currDate = Format$(Now, "dd, mmm, yyyy")
    currTime = Format$(Now, "hh:mm:ss AM/PM")
    
    destination = File1.Path & "\" & "db_CLMS [Back-up].mdb"
    Source = App.Path & "\Database\db_CLMS.mdb"
   
    Set bckupFile = FileSys.GetFile(finalpath)
    bckupFile.Attributes = Compressed
    FileSys.CopyFile Source, destination, True
    
    SaveSetting App.Title, "Settings", "BackupPath", destination
    SaveSetting App.Title, "Settings", "BackupDate", currDate
    SaveSetting App.Title, "Settings", "BackupTime", currTime

    MsgBox "Database were back-up successfully!", vbInformation, "Back-up Detail"


End Sub

Private Sub Form_Load()
    
    Dim lastPath As String
    Dim lastDate As String
    Dim lastTime As String
    File1.Visible = False
    
     
    'Read Registry for previous settings stored
    lastPath = GetSetting(App.Title, "Settings", "BackupPath")
    lastDate = GetSetting(App.Title, "Settings", "BackupDate")
    lastTime = GetSetting(App.Title, "Settings", "BackupTime")
    


    If lastPath = "" Then
    
        lbl_path.Caption = "No Backup made previously..."
        lbl_lastdate.Caption = " "
        lbl_lasttime.Caption = " "
        
    Else
    
        lbl_path.Caption = lastPath
        lbl_lastdate.Caption = lastDate
        lbl_lasttime.Caption = lastTime
        
    End If

End Sub
Private Sub Drive1_Change()

    Dir1.Path = Drive1.Drive
    
End Sub
Private Sub Dir1_Change()

    File1.Path = Dir1.Path
    
End Sub
