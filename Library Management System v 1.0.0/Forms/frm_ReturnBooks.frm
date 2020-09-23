VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_ReturnBooks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Transaction"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5610
   Icon            =   "frm_ReturnBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Source 
      Height          =   330
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4185
      Width           =   1995
   End
   Begin VB.TextBox txt_Fines 
      Height          =   330
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4590
      Width           =   1995
   End
   Begin VB.TextBox txt_StudentNo 
      Height          =   330
      Left            =   1575
      TabIndex        =   13
      Top             =   1395
      Width           =   1995
   End
   Begin VB.TextBox txt_StudentName 
      Height          =   330
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   1800
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker dtp_DateBarrowed 
      Height          =   330
      Left            =   1575
      TabIndex        =   5
      Top             =   3375
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16449536
      CurrentDate     =   39711
   End
   Begin isCoolButton.isButton btn_StudentNo 
      Height          =   330
      Left            =   3645
      TabIndex        =   4
      Top             =   2205
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_ReturnBooks.frx":3482
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
      TabIndex        =   3
      Top             =   2610
      Width           =   3705
   End
   Begin VB.TextBox txt_BookID 
      Height          =   330
      Left            =   1575
      TabIndex        =   2
      Text            =   "BID-"
      Top             =   2205
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_Ok 
      Height          =   330
      Left            =   2070
      TabIndex        =   6
      Top             =   5355
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_ReturnBooks.frx":6914
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
   Begin isCoolButton.isButton btn_Reset 
      Height          =   330
      Left            =   3150
      TabIndex        =   7
      Top             =   5355
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_ReturnBooks.frx":6930
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
      TabIndex        =   8
      Top             =   5355
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_ReturnBooks.frx":694C
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
   Begin isCoolButton.isButton btn_BookID 
      Height          =   330
      Left            =   3645
      TabIndex        =   11
      Top             =   1395
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_ReturnBooks.frx":6968
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
   Begin MSComCtl2.DTPicker dtp_DateDue 
      Height          =   330
      Left            =   1575
      TabIndex        =   17
      Top             =   3780
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   16449536
      CurrentDate     =   39711
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   315
      X2              =   5265
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000E&
      BorderWidth     =   2
      X1              =   315
      X2              =   5265
      Y1              =   1035
      Y2              =   1035
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   315
      Picture         =   "frm_ReturnBooks.frx":9DFA
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Return books from a library member by providing correct information."
      Height          =   510
      Left            =   1215
      TabIndex        =   21
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Source:"
      Height          =   240
      Left            =   360
      TabIndex        =   20
      Top             =   4230
      Width           =   1140
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Due Date:"
      Height          =   240
      Left            =   360
      TabIndex        =   18
      Top             =   3825
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Student No.:"
      Height          =   240
      Left            =   360
      TabIndex        =   15
      Top             =   1440
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Student Name:"
      Height          =   240
      Left            =   360
      TabIndex        =   14
      Top             =   1845
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Book ID:"
      Height          =   240
      Left            =   360
      TabIndex        =   10
      Top             =   2250
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Title:"
      Height          =   240
      Left            =   360
      TabIndex        =   9
      Top             =   2655
      Width           =   1140
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Fines:"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   4635
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Barrowed:"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   3420
      Width           =   1140
   End
End
Attribute VB_Name = "frm_ReturnBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RS_ReturnBook As New ADODB.Recordset
Public RS_BookUpdate As New ADODB.Recordset
Public RS_RetCheckStudent As New ADODB.Recordset
Public RS_RetCheckBooks As New ADODB.Recordset
Public RS_BookinHand As New ADODB.Recordset

Public RS_StudentIn As New ADODB.Recordset
Public RS_BookIn As New ADODB.Recordset

Public BarrowedIndex As Integer
Dim SRC_RI As String
Option Explicit

Private Sub btn_BookID_Click()

    isIssue = False
    frm_ListOfStudentsTransaction.Show 1

End Sub

Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Ok_Click()

    'On Error Resume Next
    

    'Check if Student is exist & the Book
    
    RS_RetCheckStudent.Open "SELECT * FROM tbl_Students WHERE StudentNo LIKE '" & txt_StudentNo.Text & "'", CN13, adOpenStatic, adLockOptimistic
             
             If RS_RetCheckStudent.RecordCount < 1 Then
             
                MsgBox "Student no. was not found, no record found!", vbInformation, "Result"
                RS_RetCheckStudent.Close
                Exit Sub
                
            End If
            
    RS_RetCheckBooks.Open "SELECT * FROM tbl_BorrowedBooks where BookID LIKE '" & txt_BookID.Text & "'", CN13, adOpenStatic, adLockOptimistic
        
            If RS_RetCheckBooks.RecordCount < 1 Then
        
                MsgBox "Book ID was not found, no record found!", vbInformation, "Result"
                RS_RetCheckBooks.Close
                Exit Sub
            
            End If
            
    'Update the status of the book
    
    RS_ReturnBook.Open "SELECT * FROM tbl_BorrowedBooks WHERE BookID LIKE '" & txt_BookID.Text & "' AND StudentNo LIKE '" & txt_StudentNo.Text & "' AND RowIndex LIKE '" & BarrowedIndex & "'", CN13, adOpenStatic, adLockOptimistic
    
        With RS_ReturnBook
        
            ![StatusReturned] = True
            .Update
            .Close
            
        End With
        
    If txt_Source.Text = "Reserved" Then
    
        'Update book availability RESERVED
    
        RS_BookUpdate.Open "SELECT * FROM tbl_ReservedBook WHERE ReservedBookID LIKE '" & txt_BookID.Text & "'", CN13, adOpenStatic, adLockOptimistic
        
        AvailNo = RS_BookUpdate.Fields("AvailNo")
        IssuedNo = RS_BookUpdate.Fields("IssuedNo")
        
        With RS_BookUpdate
        
            ![AvailNo] = AvailNo + 1
            ![IssuedNo] = IssuedNo - 1
            .Update
            .Close
            
        End With
    
    ElseIf txt_Source.Text = "Circulation" Then
    
        'Update book availability CIRCULATION
        
        RS_BookUpdate.Open "SELECT * FROM tbl_Book WHERE BookID LIKE '" & txt_BookID.Text & "'", CN13, adOpenStatic, adLockOptimistic
        
        AvailNo = RS_BookUpdate.Fields("AvailNo")
        IssuedNo = RS_BookUpdate.Fields("IssuedNo")
        
            With RS_BookUpdate
            
                ![AvailNo] = AvailNo + 1
                ![IssuedNo] = IssuedNo - 1
                .Update
                .Close
                
            End With
        
    End If
    
    'Update the Book in Hand
                
            RS_BookinHand.Open "SELECT * FROM tbl_Students WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN13, adOpenStatic, adLockOptimistic
            
                BookinHand = RS_BookinHand.Fields("BookinHand")
                
                With RS_BookinHand
                
                    ![BookinHand] = BookinHand - 1
                    .Update
                    .Close
                    
                End With
                

    MsgBox "Book was successfully return.", vbInformation, "information"
    
    If MsgBox("Do you want to return another book?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
    
        txt_StudentNo.Text = ""
        txt_StudentName.Text = ""
        txt_BookID.Text = "BID-"
        txt_BookTitle.Text = ""
        dtp_DateBarrowed.Value = Date
        dtp_DateDue.Value = Date
        txt_Source.Text = ""
        txt_Fines.Text = ""
        
        RS_RetCheckStudent.Close
        RS_RetCheckBooks.Close
        
    Else
    
        Unload Me
        
    End If

End Sub

Private Sub btn_Reset_Click()

    On Error Resume Next
    
    txt_StudentNo.Text = ""
    txt_StudentName.Text = ""
    txt_BookID.Text = "BID-"
    txt_BookTitle.Text = ""
    dtp_DateBarrowed.Value = Date
    dtp_DateDue.Value = Date
    txt_Source.Text = ""
    txt_Fines.Text = ""

End Sub

Private Sub btn_StudentNo_Click()

    frm_ListOfBarrowedBooks.Show 1

End Sub

Private Sub Form_Activate()

    frm_Main.toolbar_Menu.Buttons(9).Value = tbrPressed
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Book Transactions - Returned Book"
    
    

End Sub

Private Sub Form_Load()
    
    btn_StudentNo.Caption = ""
    btn_BookID.Caption = ""
    
    btn_StudentNo.IconAlign = isbCenter
    btn_BookID.IconAlign = isbCenter
    
    CN13.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"


End Sub

Private Sub Image2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next

    Set CN13 = Nothing
    
    frm_Main.toolbar_Menu.Buttons(9).Value = tbrUnpressed
    frm_Main.StatusBar1.Panels(2).Text = getPath
    
    RS_RetCheckStudent.Close
    RS_RetCheckBooks.Close

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

    RS_BookIn.Open "SELECT * FROM qry_BarrowedBooks WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN13, adOpenStatic, adLockOptimistic
          
    
    With RS_BookIn
    
        .MoveFirst
        .Find "BookID LIKE '" & Trim(txt_BookID.Text) & "'"
        
            txt_BookTitle.Text = .Fields("Title")
            dtp_DateBarrowed.Value = .Fields("DateBarrowed")
            dtp_DateDue.Value = .Fields("DateReturned")
            txt_Source.Text = .Fields("BookSource")
            txt_Fines.Text = .Fields("Fines")

        .Close
        
     End With

End Sub

Private Sub txt_StudentNo_Change()

    On Error Resume Next

    RS_StudentIn.Open "SELECT * FROM tbl_Students WHERE BookinHand >=" & 1 & "", CN13, adOpenStatic, adLockOptimistic
          
    
    With RS_StudentIn
    
        .MoveFirst
        .Find "StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'"
        
            txt_StudentName.Text = .Fields("StudentName")

        .Close
        
     End With

End Sub
