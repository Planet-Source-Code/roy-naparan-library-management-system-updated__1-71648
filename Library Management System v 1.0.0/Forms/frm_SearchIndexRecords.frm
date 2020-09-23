VERSION 5.00
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_SearchIndexRecords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Record"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7335
   Icon            =   "frm_SearchIndexRecords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbx_Type 
      Height          =   315
      ItemData        =   "frm_SearchIndexRecords.frx":1982
      Left            =   945
      List            =   "frm_SearchIndexRecords.frx":1998
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   990
      Width           =   2220
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3195
      TabIndex        =   2
      Top             =   990
      Width           =   3885
   End
   Begin isCoolButton.isButton btn_Search 
      Height          =   330
      Left            =   5895
      TabIndex        =   0
      Top             =   135
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_SearchIndexRecords.frx":19D2
      Style           =   5
      Caption         =   "&Search"
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
      Left            =   5895
      TabIndex        =   1
      Top             =   540
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_SearchIndexRecords.frx":19EE
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
   Begin VB.Label Label1 
      Caption         =   "Search records and select what record type to be search."
      Height          =   510
      Left            =   1035
      TabIndex        =   5
      Top             =   270
      Width           =   2760
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   180
      Picture         =   "frm_SearchIndexRecords.frx":1A0A
      Top             =   90
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Search by:"
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   1035
      Width           =   1005
   End
End
Attribute VB_Name = "frm_SearchIndexRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public wFilter As String
Public bFilter As String
Public eFilter As String
Public RS As ADODB.Recordset
Public SRC_FORM As Form

Option Explicit

Private Sub btn_Close_Click()

    Unload Me

End Sub



Private Sub btn_Search_Click()

    If IsEmpty(Text1) = True Then Exit Sub
    RS.Filter = ""
    RS.Requery
    
        If cbx_Type.Text = "Index ID" Then
        
            bFilter = "IndexID LIKE '%"
        
        Else
        If cbx_Type.Text = "Subject" Then
        
            bFilter = "Subject LIKE '%"
            
        Else
        If cbx_Type.Text = "Title" Then
        
            bFilter = "Title LIKE '%"
        
        Else
        If cbx_Type.Text = "Author" Then
        
            bFilter = "Author LIKE '%"
            
        Else
        If cbx_Type.Text = "Publisher" Then
        
            bFilter = "Publisher LIKE '%"
            
        Else
        If cbx_Type.Text = "Edition" Then
        
            bFilter = "Edition LIKE '%"
            
        End If
        End If
        End If
        End If
        End If
        End If
    
    RS.Filter = bFilter & Text1.Text & eFilter
    SRC_FORM.FILTER_REC
    Unload Me

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set RS = Nothing
    Set SRC_FORM = Nothing
    bFilter = ""
    eFilter = ""

End Sub
