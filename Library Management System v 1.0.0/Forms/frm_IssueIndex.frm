VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_IssueIndex 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Index Transaction"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5610
   Icon            =   "frm_IssueIndex.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_StudentName 
      Height          =   330
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker dtp_DateIssued 
      Height          =   330
      Left            =   1575
      TabIndex        =   9
      Top             =   3420
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16449536
      CurrentDate     =   39711
   End
   Begin VB.TextBox txt_StudentNo 
      Height          =   330
      Left            =   1575
      TabIndex        =   7
      Top             =   2475
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_StudentNo 
      Height          =   330
      Left            =   3645
      TabIndex        =   6
      Top             =   1440
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_IssueIndex.frx":1082
      Style           =   5
      Caption         =   "isButton1"
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
   Begin VB.TextBox txt_BookTitle 
      Height          =   555
      Left            =   1575
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1845
      Width           =   3705
   End
   Begin VB.TextBox txt_BookID 
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Text            =   "IID-"
      Top             =   1440
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_BookID 
      Height          =   330
      Left            =   3645
      TabIndex        =   8
      Top             =   2475
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_IssueIndex.frx":4514
      Style           =   5
      Caption         =   "isButton1"
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
   Begin MSComCtl2.DTPicker dtp_DateReturned 
      Height          =   330
      Left            =   1575
      TabIndex        =   10
      Top             =   3825
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16449536
      CurrentDate     =   39711
   End
   Begin isCoolButton.isButton btn_Issue 
      Height          =   330
      Left            =   2070
      TabIndex        =   12
      Top             =   4635
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueIndex.frx":79A6
      Style           =   5
      Caption         =   "&Issue"
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
   Begin isCoolButton.isButton btn_Reset 
      Height          =   330
      Left            =   3150
      TabIndex        =   13
      Top             =   4635
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueIndex.frx":79C2
      Style           =   5
      Caption         =   "&Reset"
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
      Left            =   4230
      TabIndex        =   14
      Top             =   4635
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueIndex.frx":79DE
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
   Begin VB.Line Line3 
      BorderColor     =   &H80000010&
      X1              =   315
      X2              =   5265
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   315
      X2              =   5265
      Y1              =   1125
      Y2              =   1125
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Issue a non-book materials to library members by providing correct information."
      Height          =   645
      Left            =   1170
      TabIndex        =   17
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Index ID:"
      Height          =   240
      Left            =   360
      TabIndex        =   16
      Top             =   1485
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Index Title:"
      Height          =   240
      Left            =   360
      TabIndex        =   15
      Top             =   1890
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_IssueIndex.frx":79FA
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Return:"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   3825
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Issued:"
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   3465
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Student Name:"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   2925
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Student No.:"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   2520
      Width           =   1140
   End
End
Attribute VB_Name = "frm_IssueIndex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_Issue As New ADODB.Recordset


Dim RS_BookIn As New ADODB.Recordset
Dim RS_StudentIn As New ADODB.Recordset
Dim RS_IndexinHand As New ADODB.Recordset

Option Explicit

Private Sub btn_BookID_Click()

    isIssueIndex = True
    frm_ListOfStudentsIndexTransaction.Show 1

End Sub

Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Issue_Click()

On Error Resume Next
    
                
            'Add book to Barrowed Book
            
            RS_Issue.Open "SELECT * FROM tbl_BorrowedIndex", CN3, adOpenStatic, adLockOptimistic
            
                With RS_Issue
                
                    .AddNew
                    ![StudentNo] = txt_StudentNo.Text
                    ![IndexID] = txt_BookID.Text
                    ![DateBarrowed] = dtp_DateIssued.Value
                    ![DateReturned] = dtp_DateReturned.Value
                    ![StatusReturned] = False
                    .Update
                
                End With
             
                'Update the Index in Hand
                
                RS_IndexinHand.Open "SELECT * FROM tbl_Students WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN3, adOpenStatic, adLockOptimistic
            
                IndexinHand = RS_IndexinHand.Fields("IndexinHand")
                
                With RS_IndexinHand
                
                    ![IndexinHand] = IndexinHand + 1
                    .Update
                
                End With
                
                RS_IndexinHand.Close
                
                MsgBox "Index was issued successfully", vbInformation, "Information"
                
                RS_Issue.Close
                RS_BookIn.Close
                RS_IndexinHand.Close
                
                If MsgBox("Do you want to issue another Index?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                
                    txt_BookID.Text = "IID-"
                    txt_BookTitle.Text = ""
                    txt_StudentNo.Text = ""
                    txt_StudentName.Text = ""
                    dtp_DateIssued.Value = Date
                    dtp_DateReturned.Value = Date
             
                    
                Else
                
                    Unload Me
                    
                End If
            
    

End Sub

Private Sub btn_Reset_Click()

    On Error Resume Next
    
   
    txt_BookID.Text = "IID-"
    txt_BookTitle.Text = ""
    txt_StudentNo.Text = ""
    txt_StudentName.Text = ""
    dtp_DateIssued.Value = Date
    dtp_DateReturned.Value = Date


End Sub

Private Sub btn_StudentNo_Click()
    

    
        'isIssueIndex = True
        frm_ListOfIndexTransaction.Show 1
    
 

End Sub

Private Sub Form_Activate()

    'On Error Resume Next
    
    
    
    frm_Main.toolbar_Menu.Buttons(10).Value = tbrPressed
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Index Transactions - Issue Index"

End Sub

Private Sub Form_Load()
    
    dtp_DateIssued.Value = Date
    dtp_DateReturned.Value = Date
    'dtp_DateReturned.Value = dtp_DateReturned.Value + dayslimit
    btn_StudentNo.Caption = ""
    btn_BookID.Caption = ""
    
    btn_StudentNo.IconAlign = isbCenter
    btn_BookID.IconAlign = isbCenter
    
    CN3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frm_Main.toolbar_Menu.Buttons(10).Value = tbrUnpressed
    frm_Main.StatusBar1.Panels(2).Text = getPath
    Set CN3 = Nothing

End Sub
Private Function getPath() As String
On Error Resume Next
Dim t As String
Open App.Path & "\ActiveForm.dat" For Input As #1
    Input #1, t
Close #1
getPath = Trim$(t)
t = vbNullString
End Function

Private Sub isButton1_Click()

End Sub

Private Sub isButton3_Click()

    Unload Me

End Sub

Private Sub txt_BookID_Change()

    On Error Resume Next

    RS_BookIn.Open "SELECT * FROM tbl_Index", CN3, adOpenStatic, adLockOptimistic
          
    
    With RS_BookIn
    
        .MoveFirst
        .Find "IndexID LIKE '" & Trim(txt_BookID.Text) & "'"
        
            txt_BookTitle.Text = .Fields("Title")

        .Close
        
     End With

End Sub

Private Sub txt_StudentNo_Change()

    On Error Resume Next

    RS_StudentIn.Open "SELECT * FROM tbl_Students", CN3, adOpenStatic, adLockOptimistic
          
    
    With RS_StudentIn
    
        .MoveFirst
        .Find "StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'"
        
            txt_StudentName.Text = .Fields("StudentName")

        .Close
        
     End With

End Sub
