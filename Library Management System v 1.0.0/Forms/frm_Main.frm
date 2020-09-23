VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248850FC-2BAF-48AF-99D6-220E54FE68CA}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm frm_Main 
   BackColor       =   &H8000000C&
   Caption         =   "Library Management System v 1.0.0"
   ClientHeight    =   8610
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   12615
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList img_Pic3 
      Left            =   11340
      Top             =   1440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   64
      ImageHeight     =   64
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3482
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_Pic2 
      Left            =   11970
      Top             =   1440
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
            Picture         =   "frm_Main.frx":66DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":9B6E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList img_Pic 
      Left            =   10125
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   72
      ImageHeight     =   72
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":D000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   60
      Left            =   0
      ScaleHeight     =   60
      ScaleWidth      =   12615
      TabIndex        =   2
      Top             =   600
      Width           =   12615
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   12015
      Top             =   765
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   30
      Bmp:1           =   "frm_Main.frx":E0B8
      Key:1           =   "#mnu_UserManagement"
      Bmp:2           =   "frm_Main.frx":E4E0
      Key:2           =   "#mnu_ChangeUser"
      Bmp:3           =   "frm_Main.frx":E908
      Key:3           =   "#mnu_ExitProgram"
      Bmp:4           =   "frm_Main.frx":ED30
      Key:4           =   "#mnu_Books"
      Bmp:5           =   "frm_Main.frx":F158
      Key:5           =   "#mnu_CardCatalog"
      Bmp:6           =   "frm_Main.frx":F580
      Key:6           =   "#mnu_Index"
      Bmp:7           =   "frm_Main.frx":F9A8
      Key:7           =   "#mnu_Students"
      Bmp:8           =   "frm_Main.frx":FDD0
      Key:8           =   "#mnu_IssueBooks"
      Bmp:9           =   "frm_Main.frx":101F8
      Key:9           =   "#mnu_ReturnBooks"
      Bmp:10          =   "frm_Main.frx":10620
      Key:10          =   "#mnu_Notepad"
      Bmp:11          =   "frm_Main.frx":10A48
      Key:11          =   "#mnu_Calculator"
      Bmp:12          =   "frm_Main.frx":10E70
      Key:12          =   "#mnu_About"
      Bmp:13          =   "frm_Main.frx":11298
      Key:13          =   "#mnu_BackupDatabase"
      Bmp:14          =   "frm_Main.frx":116C0
      Key:14          =   "#mnu_SystemSettings"
      Bmp:15          =   "frm_Main.frx":11AE8
      Key:15          =   "#mnu_Manual"
      Bmp:16          =   "frm_Main.frx":11F10
      Key:16          =   "#mnu_Calendar"
      Bmp:17          =   "frm_Main.frx":12338
      Key:17          =   "#mnu_IssueIndex"
      Bmp:18          =   "frm_Main.frx":12760
      Key:18          =   "#mnu_ReturnIndex"
      Bmp:19          =   "frm_Main.frx":12B88
      Key:19          =   "#mnu_SearcCardCatalog"
      Bmp:20          =   "frm_Main.frx":12FB0
      Key:20          =   "#mnu_SearchIndex"
      Bmp:21          =   "frm_Main.frx":133D8
      Key:21          =   "#mnu_BookRecords"
      Bmp:22          =   "frm_Main.frx":13800
      Key:22          =   "#mnu_Circulation"
      Bmp:23          =   "frm_Main.frx":13C28
      Key:23          =   "#mnu_Reserved"
      Bmp:24          =   "frm_Main.frx":14050
      Key:24          =   "#mnu_AllBooks"
      Bmp:25          =   "frm_Main.frx":14478
      Key:25          =   "#mnu_StudentRecords"
      Bmp:26          =   "frm_Main.frx":148A0
      Key:26          =   "#mnu_CardCatalogRecords"
      Bmp:27          =   "frm_Main.frx":14CC8
      Key:27          =   "#mnu_IndexRecords"
      Bmp:28          =   "frm_Main.frx":150F0
      Key:28          =   "#mnu_BarrowedBooks"
      Bmp:29          =   "frm_Main.frx":15518
      Key:29          =   "#mnu_Indexes"
      Bmp:30          =   "frm_Main.frx":15940
      Key:30          =   "#mnu_BarrowedIndexes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList imglst_Toolbar 
      Left            =   11385
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":15D68
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":176FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1908C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1AA1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1C3B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1DD42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":1F6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":21066
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":229F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2438A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":25D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":276AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":29040
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2A9D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2C364
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":2DCF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   8235
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   10
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   706
            MinWidth        =   706
            Picture         =   "frm_Main.frx":2F688
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Text            =   "Active Form: None"
            TextSave        =   "Active Form: None"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Picture         =   "frm_Main.frx":3071A
            Text            =   "User:"
            TextSave        =   "User:"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2469
            MinWidth        =   2469
            Picture         =   "frm_Main.frx":317AC
            Text            =   "Date & Time:"
            TextSave        =   "Date & Time:"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "17/01/2009"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            TextSave        =   "12:44 AM"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel10 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "INS"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar toolbar_Menu 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "imglst_Toolbar"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_Books"
            Object.ToolTipText     =   "Book Records"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_ReservedBooks"
            Object.ToolTipText     =   "Reserved Book Records"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_CardCatalog"
            Object.ToolTipText     =   "Card Catalog"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_Index"
            Object.ToolTipText     =   "Index"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_StudentRecords"
            Object.ToolTipText     =   "Student Records"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_IssueBook"
            Object.ToolTipText     =   "Issue Book"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_ReturnBook"
            Object.ToolTipText     =   "Return Book"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_IssueIndex"
            Object.ToolTipText     =   "Issue Index"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_ReturnIndex"
            Object.ToolTipText     =   "Return Index"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_SearchCardCatalog"
            Object.ToolTipText     =   "Search Card Catalog"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_SearchIndex"
            Object.ToolTipText     =   "Search Index"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_Chart"
            Object.ToolTipText     =   "Library Statistics"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_GlobalRecords"
            Object.ToolTipText     =   "Global Records"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btn_Settings"
            Object.ToolTipText     =   "Settings"
            ImageIndex      =   12
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   10755
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":3283E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":33250
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":33C62
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":34674
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":35086
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":35A98
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":364AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":36EBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":378CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":382E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":38CF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm_Main.frx":39704
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Index           =   1
      Begin VB.Menu mnu_UserManagement 
         Caption         =   "User Management"
         Shortcut        =   ^U
      End
      Begin VB.Menu mnu_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ChangeUser 
         Caption         =   "Sign Out"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnu_ExitProgram 
         Caption         =   "Exit Program"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnu_Records 
      Caption         =   "R&ecords"
      Index           =   2
      Begin VB.Menu mnu_Books 
         Caption         =   "Books"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnu_CardCatalog 
         Caption         =   "Card Catalog"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnu_Index 
         Caption         =   "Index"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnu_Students 
         Caption         =   "Students"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnu_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SearcCardCatalog 
         Caption         =   "Search Card Catalog"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnu_SearchIndex 
         Caption         =   "Search Index"
         Shortcut        =   {F6}
      End
   End
   Begin VB.Menu mnu_Transaction 
      Caption         =   "Tran&saction"
      Index           =   3
      Begin VB.Menu mnu_IssueBooks 
         Caption         =   "Issue Books"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnu_ReturnBooks 
         Caption         =   "Return Books"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnu_Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_IssueIndex 
         Caption         =   "Issue Index"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnu_ReturnIndex 
         Caption         =   "Return Index"
         Shortcut        =   {F4}
      End
   End
   Begin VB.Menu mnu_Report 
      Caption         =   "&Reports"
      Index           =   4
      Begin VB.Menu mnu_BookRecords 
         Caption         =   "Book Records"
         Begin VB.Menu mnu_Circulation 
            Caption         =   "Circulation"
            Shortcut        =   +^{F1}
         End
         Begin VB.Menu mnu_Reserved 
            Caption         =   "Reserved"
            Shortcut        =   +^{F2}
         End
         Begin VB.Menu mnu_Sep7 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_BarrowedBooks 
            Caption         =   "Barrowed Books"
            Shortcut        =   +^{F3}
         End
      End
      Begin VB.Menu mnu_StudentRecords 
         Caption         =   "Student Records"
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnu_CardCatalogRecords 
         Caption         =   "Card Catalog Records"
         Shortcut        =   +{F2}
      End
      Begin VB.Menu mnu_IndexRecords 
         Caption         =   "Index Records"
         Begin VB.Menu mnu_Indexes 
            Caption         =   "Indexes"
            Shortcut        =   +^{F4}
         End
         Begin VB.Menu mnu_Sep8 
            Caption         =   "-"
         End
         Begin VB.Menu mnu_BarrowedIndexes 
            Caption         =   "Barrowed Indexes"
            Shortcut        =   +^{F5}
         End
      End
   End
   Begin VB.Menu mnu_View 
      Caption         =   "&View"
      Index           =   5
      Begin VB.Menu mnu_ShowToolbar 
         Caption         =   "Show Toolbar"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnu_Status 
         Caption         =   "Show Status Bar"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "&Tools"
      Index           =   6
      Begin VB.Menu mnu_Notepad 
         Caption         =   "Notepad"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnu_Calculator 
         Caption         =   "Calculator"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnu_Calendar 
         Caption         =   "Calendar"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnu_Sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_BackupDatabase 
         Caption         =   "Back-up Database"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnu_Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_SystemSettings 
         Caption         =   "System Settings"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnu_Help 
      Caption         =   "&Help"
      Index           =   7
      Begin VB.Menu mnu_Manual 
         Caption         =   "LMS Help"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnu_Sep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "About LMS"
         Shortcut        =   {F9}
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Out As Boolean
Private UserLogRowIndex As Integer
Private RS_UserLogUpdate As New ADODB.Recordset
Private RS_AdminLogUpdate As New ADODB.Recordset
Private RS_rptBooks As New ADODB.Recordset
Private RS_rptStudents As New ADODB.Recordset
Private RS_rptCardCatalog As New ADODB.Recordset
Private RS_rptIndex As New ADODB.Recordset
Private RS_rptReservedBooks As New ADODB.Recordset
Private RS_rptBarrowedBooks As New ADODB.Recordset
Private RS_rptBarrowedIndexes As New ADODB.Recordset
Option Explicit
Private Sub MDIForm_Initialize()

    Call GET_LAYOUT

End Sub
Private Function getUserIndex() As String
On Error Resume Next
Dim t As String
Open App.Path & "\UserLog.dat" For Input As #1
    Input #1, t
Close #1
getUserIndex = Trim$(t)
t = vbNullString
End Function
Private Function getAdminIndex() As String
On Error Resume Next
Dim t As String
Open App.Path & "\IsActive.dat" For Input As #1
    Input #1, t
Close #1
getAdminIndex = Trim$(t)
t = vbNullString
End Function
Private Sub MDIForm_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    On Error Resume Next
    If Button = vbRightButton Then
    
        PopupMenu mnu_Records
        
    End If

End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim repp As Integer

        On Error Resume Next
   
        repp = MsgBox("This will terminate the application. Do you want to proceed?", vbQuestion + vbYesNo, "Terminate")
        If repp = vbNo Then
        Cancel = 1
        Else
        
        CN6.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
        RS_UserLogUpdate.Open "SELECT * FROM tbl_UserLog WHERE UserLogNo LIKE '" & getUserIndex & "'", CN6, adOpenStatic, adLockOptimistic
    
        With RS_UserLogUpdate
        
            ![DateLogOut] = Date
            ![TimeLogOut] = Time
            .Update
            .Close
            
        End With
        
        RS_AdminLogUpdate.Open "SELECT * FROM tbl_User WHERE RowIndex LIKE '" & getAdminIndex & "'", CN6, adOpenStatic, adLockOptimistic
        
            With RS_AdminLogUpdate
            
                ![Active] = False
                .Update
                .Close
            
            End With
        
        Kill (App.Path & "\IsActive.dat")
        Kill (App.Path & "\UserLog.dat")
        Set CN6 = Nothing
        
        End If

        

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
  Call SAVE_LAYOUT
  
End Sub

Private Sub mnu_About_Click()

    frm_About.Show 1

End Sub



Private Sub mnu_BackupDatabase_Click()

    frm_Backup.Show 1

End Sub

Private Sub mnu_BarrowedBooks_Click()

    On Error Resume Next
    CN22.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptBarrowedBooks.Open "SELECT * FROM qry_BarrowedBooks WHERE StatusReturned =" & False & "", CN22, adOpenStatic, adLockOptimistic
    
    Set rpt_ListsofBarrowedBooks.DataSource = RS_rptBarrowedBooks
    
    rpt_ListsofBarrowedBooks.Show 1
    CN22.Close

End Sub

Private Sub mnu_BarrowedIndexes_Click()

    On Error Resume Next
    CN23.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptBarrowedIndexes.Open "SELECT * FROM qry_BarrowedIndex WHERE StatusReturned =" & False & "", CN23, adOpenStatic, adLockOptimistic
    
    Set rpt_ListsofBarrowedIndexes.DataSource = RS_rptBarrowedIndexes
    
    rpt_ListsofBarrowedIndexes.Show 1
    CN23.Close

End Sub

Private Sub mnu_Books_Click()

    Call UNLOAD_CHILDS
            frm_BooksRecord.Show
            frm_BooksRecord.SetFocus

End Sub

Private Sub mnu_Calculator_Click()

    On Error GoTo err
    Shell "Calc.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "No application installed!", vbInformation, "Error"

End Sub

Private Sub mnu_Calendar_Click()

    frm_Calendar.Show
    frm_Calendar.SetFocus

End Sub

Private Sub mnu_CardCatalog_Click()

    Call UNLOAD_CHILDS
            frm_CardCatalog.Show
            frm_CardCatalog.SetFocus

End Sub

Private Sub mnu_CardCatalogRecords_Click()

    On Error Resume Next
    CN9.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptCardCatalog.Open "SELECT * FROM tbl_Catalog", CN9, adOpenStatic, adLockOptimistic
    
    Set rpt_CardCatalog.DataSource = RS_rptCardCatalog
    
    rpt_CardCatalog.Show 1
    CN9.Close

End Sub

Private Sub mnu_ChangeUser_Click()

    On Error Resume Next
    
    If MsgBox("All active forms will be closed. Do you want to proceed?", vbQuestion + vbYesNo, "Sign Out") = vbYes Then
    
    
        CN6.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
        RS_UserLogUpdate.Open "SELECT * FROM tbl_UserLog WHERE UserLogNo LIKE '" & getUserIndex & "'", CN6, adOpenStatic, adLockOptimistic
    
        With RS_UserLogUpdate
        
            ![DateLogOut] = Date
            ![TimeLogOut] = Time
            .Update
            .Close
            
        End With
        
        RS_AdminLogUpdate.Open "SELECT * FROM tbl_User WHERE RowIndex LIKE '" & getAdminIndex & "'", CN6, adOpenStatic, adLockOptimistic
        
            With RS_AdminLogUpdate
            
                ![Active] = False
                .Update
                .Close
            
        End With
        
        Kill (App.Path & "\IsActive.dat")
        Kill (App.Path & "\UserLog.dat")
        Set CN6 = Nothing
        
        Call UNLOAD_CHILDS
        Me.Hide
        frm_Login.Show 1
        
   End If

End Sub

Private Sub mnu_Circulation_Click()

    On Error Resume Next
    CN7.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptBooks.Open "SELECT * FROM tbl_Book", CN7, adOpenStatic, adLockOptimistic
    
    Set rpt_Books.DataSource = RS_rptBooks
    
    rpt_Books.Show 1
    CN7.Close

End Sub

Private Sub mnu_ExitProgram_Click()

    On Error Resume Next
    
    Dim repp As Integer
    Dim Cancel As Integer
    repp = MsgBox("This will terminate the application. Do you want to proceed?", vbExclamation + vbYesNo, "Terminate")
    If repp = vbYes Then
    
        CN6.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
        RS_UserLogUpdate.Open "SELECT * FROM tbl_UserLog WHERE UserLogNo LIKE '" & getUserIndex & "'", CN6, adOpenStatic, adLockOptimistic
    
        With RS_UserLogUpdate
        
            ![DateLogOut] = Date
            ![TimeLogOut] = Time
            .Update
            .Close
            
        End With
        
        RS_AdminLogUpdate.Open "SELECT * FROM tbl_User WHERE RowIndex LIKE '" & getAdminIndex & "'", CN6, adOpenStatic, adLockOptimistic
        
            With RS_AdminLogUpdate
            
                ![Active] = False
                .Update
                .Close
            
        End With
        
        Kill (App.Path & "\IsActive.dat")
        Kill (App.Path & "\UserLog.dat")
        Set CN6 = Nothing
        
    End
    End If

End Sub

Private Sub mnu_Index_Click()

    Call UNLOAD_CHILDS
            frm_Index.Show
            frm_Index.SetFocus

End Sub

Private Sub mnu_Indexes_Click()

    On Error Resume Next
    CN10.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptIndex.Open "SELECT * FROM tbl_Index", CN10, adOpenStatic, adLockOptimistic
    
    Set rpt_Index.DataSource = RS_rptIndex
    
    rpt_Index.Show 1
    CN10.Close

End Sub

Private Sub mnu_IssueBooks_Click()
    
    frm_IssueBooks.Show 1

End Sub

Private Sub mnu_IssueIndex_Click()

    frm_IssueIndex.Show 1

End Sub

Private Sub mnu_Manual_Click()

    OpenURL App.Path & "\Manual\Manual.htm", frm_Main.hwnd

End Sub

Private Sub mnu_Notepad_Click()

    On Error GoTo err
    Shell "Notepad.exe", vbNormalFocus
    Exit Sub
err:
    MsgBox "No application installed!", vbInformation, "Error"

End Sub



Private Sub mnu_Reserved_Click()

    On Error Resume Next
    CN11.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptReservedBooks.Open "SELECT * FROM tbl_ReservedBook", CN11, adOpenStatic, adLockOptimistic
    
    Set rpt_ReservedBooks.DataSource = RS_rptReservedBooks
    
    rpt_ReservedBooks.Show 1
    CN11.Close

End Sub

Private Sub mnu_ReturnBooks_Click()

    frm_ReturnBooks.Show 1

End Sub

Private Sub mnu_ReturnIndex_Click()

    frm_ReturnIndex.Show 1

End Sub

Private Sub mnu_SearcCardCatalog_Click()

    Call UNLOAD_CHILDS
    frm_SearchCardCatalog.Show
    frm_SearchCardCatalog.lvList.SetFocus

End Sub

Private Sub mnu_SearchIndex_Click()

    Call UNLOAD_CHILDS
    frm_SearchIndex.Show
    frm_SearchIndex.lvList.SetFocus

End Sub

Private Sub mnu_ShowToolbar_Click()

    mnu_ShowToolbar.Checked = Not mnu_ShowToolbar.Checked
    toolbar_Menu.Visible = mnu_ShowToolbar.Checked

End Sub

Private Sub mnu_Status_Click()

    mnu_Status.Checked = Not mnu_Status.Checked
    StatusBar1.Visible = mnu_Status.Checked

End Sub

Private Sub mnu_StudentRecords_Click()

    On Error Resume Next
    CN8.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_rptStudents.Open "SELECT * FROM tbl_Students", CN8, adOpenStatic, adLockOptimistic
    
    Set rpt_Students.DataSource = RS_rptStudents
    
    rpt_Students.Show 1
    CN8.Close

End Sub

Private Sub mnu_Students_Click()

    Call UNLOAD_CHILDS
            frm_StudentRecords.Show
            frm_StudentRecords.SetFocus

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

End Sub

Private Sub tmrMemStatus_Timer()

End Sub

Private Sub mnu_SystemSettings_Click()

    frm_Settings.Show 1

End Sub

Private Sub mnu_UserManagement_Click()

    frm_UserManagement.Show 1

End Sub

Private Sub toolbar_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.Key
    
        Case "btn_Books"
        
            Call UNLOAD_CHILDS
            frm_BooksRecord.Show
            frm_BooksRecord.lvList.SetFocus
            
        Case "btn_StudentRecords"
        
            Call UNLOAD_CHILDS
            frm_StudentRecords.Show
            frm_StudentRecords.lvList.SetFocus
            
        Case "btn_CardCatalog"
        
            Call UNLOAD_CHILDS
            frm_CardCatalog.Show
            frm_CardCatalog.lvList.SetFocus
        
        Case "btn_Index"
        
            Call UNLOAD_CHILDS
            frm_Index.Show
            frm_Index.lvList.SetFocus
            
        Case "btn_IssueBook"
        
            'Call UNLOAD_CHILDS
            frm_IssueBooks.Show 1
        
        Case "btn_ReturnBook"
        
            'Call UNLOAD_CHILDS
            frm_ReturnBooks.Show 1
            
        Case "btn_IssueIndex"
        
            'Call UNLOAD_CHILDS
            frm_IssueIndex.Show 1
        
        Case "btn_ReturnIndex"
        
            'Call UNLOAD_CHILDS
            frm_ReturnIndex.Show 1
        
        Case "btn_SearchCardCatalog"
            
            Call UNLOAD_CHILDS
            frm_SearchCardCatalog.Show
            frm_SearchCardCatalog.lvList.SetFocus
            
        Case "btn_ReservedBooks"
        
            Call UNLOAD_CHILDS
            frm_ReservedBooksRecord.Show
            frm_ReservedBooksRecord.lvList.SetFocus
            
        Case "btn_SearchIndex"
        
            Call UNLOAD_CHILDS
            frm_SearchIndex.Show
            frm_SearchIndex.lvList.SetFocus
            
        Case "btn_Chart"
        
            frm_Statistics.Show 1
            
        Case "btn_Settings"
        
             frm_Settings.Show 1
            
    End Select

End Sub

Public Sub UNLOAD_CHILDS()

    Dim Form As Form
       For Each Form In Forms
          If Form.Name <> Me.Name Then Unload Form
       Next Form
    Set Form = Nothing

End Sub

Public Sub RESET_STATUS()
StatusBar1.Panels(2).Text = "Active Form: None"
End Sub

Public Sub RESTORE_BUTTON_VALUE()
Dim i As Byte
For i = 1 To toolbar_Menu.Buttons.Count
    toolbar_Menu.Buttons(i).Value = tbrUnpressed
Next i
i = 0

toolbar_Menu.Refresh
End Sub

Private Sub MDIForm_Resize()
If Me.WindowState <> vbMinimized Then
    If Me.Width < 11805 Then Me.Width = 11805
    If Me.Height < 8280 Then Me.Height = 8280
   
End If
End Sub

Private Sub GET_LAYOUT()
On Error Resume Next
Dim T1 As String, T2 As String, T3 As String
''Get the layout.dat contents
Open App.Path & "\Layout.dat" For Input As #1
    Input #1, T1
    Input #1, T2

Close #1

''Remove white space
T1 = Trim$(T1)
T2 = Trim$(T2)


''Assign the value
mnu_ShowToolbar.Checked = T1
mnu_Status.Checked = T2


''Perform the action
mnu_ShowToolbar_Click
mnu_Status_Click


''Clear the variables
T1 = ""
T2 = ""
T3 = ""
End Sub
Private Sub SAVE_LAYOUT()
On Error Resume Next
''Save the layout to file
Open App.Path & "\Layout.dat" For Output As #1
    Print #1, Not mnu_ShowToolbar.Checked
    Print #1, Not mnu_Status.Checked

Close #1
End Sub
