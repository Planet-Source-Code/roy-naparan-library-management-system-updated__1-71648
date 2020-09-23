VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Begin VB.Form frm_Calendar 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   2655
   Icon            =   "frm_Calendar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   2655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2310
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2640
      _ExtentX        =   4657
      _ExtentY        =   4075
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   0
      StartOfWeek     =   62783489
      TitleBackColor  =   33023
      CurrentDate     =   39656
   End
End
Attribute VB_Name = "frm_Calendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
MonthView1.Value = Now
End Sub

