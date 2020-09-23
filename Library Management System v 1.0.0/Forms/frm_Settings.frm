VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_Settings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4035
   Icon            =   "frm_Settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   4035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   4
      Left            =   2385
      TabIndex        =   6
      Top             =   2835
      Width           =   1365
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   3
      Left            =   2385
      TabIndex        =   4
      Top             =   2430
      Width           =   1365
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   1
      Left            =   2385
      TabIndex        =   1
      Top             =   1800
      Width           =   1365
   End
   Begin VB.TextBox txtEntry 
      Height          =   330
      Index           =   0
      Left            =   2385
      TabIndex        =   0
      Top             =   1395
      Width           =   1365
   End
   Begin isCoolButton.isButton btn_Save 
      Height          =   330
      Left            =   1665
      TabIndex        =   8
      Top             =   3600
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_Settings.frx":1982
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
      Left            =   2745
      TabIndex        =   9
      Top             =   3600
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_Settings.frx":199E
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
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   180
      X2              =   3760
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   180
      X2              =   3760
      Y1              =   1170
      Y2              =   1170
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_Settings.frx":19BA
      Top             =   225
      Width           =   720
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Assign a default value of the following transaction details."
      Height          =   600
      Left            =   1125
      TabIndex        =   10
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Fine charge per day:"
      Height          =   285
      Left            =   270
      TabIndex        =   7
      Top             =   2880
      Width           =   2040
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Fines:"
      Height          =   285
      Left            =   270
      TabIndex        =   5
      Top             =   2475
      Width           =   2040
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Max. books to hold:"
      Height          =   285
      Left            =   270
      TabIndex        =   3
      Top             =   1845
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Max.days to hold a books:"
      Height          =   285
      Left            =   270
      TabIndex        =   2
      Top             =   1440
      Width           =   2040
   End
End
Attribute VB_Name = "frm_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_AssignGet As New ADODB.Recordset
Dim RS_AssignUpdate As New ADODB.Recordset
Option Explicit

Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Save_Click()

    RS_AssignUpdate.Open "SELECT * FROM tbl_Settings", CN5, adOpenStatic, adLockOptimistic
    
        With RS_AssignUpdate
    
            .Fields("MaxDayHold") = txtEntry(0).Text
            .Fields("MaxHold") = txtEntry(1).Text
            .Fields("Fines") = txtEntry(3).Text
            .Fields("FineCharge") = txtEntry(4).Text
            .Update
            
            
            dayslimit = .Fields("MaxDayHold")
            maxhold = .Fields("MaxHold")
            Fines = .Fields("Fines")
            rateperday = .Fields("FineCharge")
            
        End With
        
    MsgBox "New settings updated successfully.", vbInformation, "Information"
    RS_AssignUpdate.Close
    Unload Me

End Sub

Private Sub Form_Load()

    CN5.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"
 
    RS_AssignGet.Open "SELECT * FROM tbl_Settings", CN5, adOpenStatic, adLockOptimistic
    
        With RS_AssignGet
    
            txtEntry(0) = .Fields("MaxDayHold")
            txtEntry(1) = .Fields("MaxHold")
            txtEntry(3) = Format(.Fields("Fines"), "#,##0.00")
            txtEntry(4) = Format(.Fields("FineCharge"), "#,##0.00")
            
        End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    RS_AssignGet.Close
    Set RS_AssignGet = Nothing
    Set CN5 = Nothing

End Sub
