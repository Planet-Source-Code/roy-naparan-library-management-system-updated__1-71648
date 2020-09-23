VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_LoginSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log-in Settings"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   4560
   Icon            =   "frm_LoginSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chk_MaskUsername 
      Caption         =   "Mask User Name"
      Height          =   375
      Left            =   2430
      TabIndex        =   5
      Top             =   180
      Width           =   1950
   End
   Begin VB.CheckBox chk_UnmaskPassword 
      Caption         =   "Unmask Password"
      Height          =   375
      Left            =   2430
      TabIndex        =   4
      Top             =   630
      Width           =   1950
   End
   Begin VB.CheckBox chk_SavePassword 
      Caption         =   "Save Password"
      Height          =   375
      Left            =   180
      TabIndex        =   3
      Top             =   630
      Width           =   1950
   End
   Begin VB.CheckBox chk_SaveUsername 
      Caption         =   "Save User Name"
      Height          =   375
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   1950
   End
   Begin isCoolButton.isButton btn_Ok 
      Height          =   300
      Left            =   2250
      TabIndex        =   0
      Top             =   1260
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      Icon            =   "frm_LoginSettings.frx":1982
      Style           =   5
      Caption         =   "&Ok"
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
      Height          =   300
      Left            =   3330
      TabIndex        =   1
      Top             =   1260
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   529
      Icon            =   "frm_LoginSettings.frx":199E
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
Attribute VB_Name = "frm_LoginSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Ok_Click()

    If chk_UnmaskPassword.Value = Checked Then
    
        frm_Login.txt_Password.PasswordChar = ""
        
    ElseIf chk_UnmaskPassword.Value = Unchecked Then
    
        frm_Login.txt_Password.PasswordChar = "•"
        
    End If
    
    Unload Me

End Sub
