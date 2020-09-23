VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Begin VB.Form frm_IssueBooks 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Books Transaction"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   5610
   Icon            =   "frm_IssueBooks.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cbx_Source 
      Height          =   315
      ItemData        =   "frm_IssueBooks.frx":3482
      Left            =   1575
      List            =   "frm_IssueBooks.frx":348C
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   1485
      Width           =   1995
   End
   Begin VB.TextBox txt_StudentName 
      Height          =   330
      Left            =   1575
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   3330
      Width           =   3705
   End
   Begin MSComCtl2.DTPicker dtp_DateIssued 
      Height          =   330
      Left            =   1575
      TabIndex        =   9
      Top             =   3870
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   62521344
      CurrentDate     =   39711
   End
   Begin VB.TextBox txt_StudentNo 
      Height          =   330
      Left            =   1575
      TabIndex        =   7
      Top             =   2925
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_StudentNo 
      Height          =   330
      Left            =   3645
      TabIndex        =   6
      Top             =   1890
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_IssueBooks.frx":34A7
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
      Top             =   2295
      Width           =   3705
   End
   Begin VB.TextBox txt_BookID 
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Text            =   "BID-"
      Top             =   1890
      Width           =   1995
   End
   Begin isCoolButton.isButton btn_BookID 
      Height          =   330
      Left            =   3645
      TabIndex        =   8
      Top             =   2925
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   582
      Icon            =   "frm_IssueBooks.frx":6939
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
      Top             =   4275
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   582
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   62521344
      CurrentDate     =   39711
   End
   Begin isCoolButton.isButton btn_Issue 
      Height          =   330
      Left            =   2070
      TabIndex        =   12
      Top             =   5040
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueBooks.frx":9DCB
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
      Top             =   5040
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueBooks.frx":9DE7
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
      Top             =   5040
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   582
      Icon            =   "frm_IssueBooks.frx":9E03
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
      Caption         =   "Issue a book to library members by providing correct information."
      Height          =   510
      Left            =   1170
      TabIndex        =   19
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Source:"
      Height          =   240
      Left            =   360
      TabIndex        =   18
      Top             =   1530
      Width           =   1140
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Book ID:"
      Height          =   240
      Left            =   360
      TabIndex        =   16
      Top             =   1935
      Width           =   1140
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Book Title:"
      Height          =   240
      Left            =   360
      TabIndex        =   15
      Top             =   2340
      Width           =   1140
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   270
      Picture         =   "frm_IssueBooks.frx":9E1F
      Top             =   135
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Return:"
      Height          =   240
      Left            =   360
      TabIndex        =   3
      Top             =   4275
      Width           =   1140
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Issued:"
      Height          =   240
      Left            =   360
      TabIndex        =   2
      Top             =   3915
      Width           =   1140
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Student Name:"
      Height          =   240
      Left            =   360
      TabIndex        =   1
      Top             =   3375
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Student No.:"
      Height          =   240
      Left            =   360
      TabIndex        =   0
      Top             =   2970
      Width           =   1140
   End
End
Attribute VB_Name = "frm_IssueBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RS_Issue As New ADODB.Recordset
Dim RS_CheckBooks As New ADODB.Recordset
Dim RS_CheckStudent As New ADODB.Recordset
Dim RS_CheckIssue As New ADODB.Recordset
Dim RS_BookUpdate As New ADODB.Recordset
Dim RS_BookinHand As New ADODB.Recordset

Dim RS_BookIn As New ADODB.Recordset
Dim RS_StudentIn As New ADODB.Recordset

Option Explicit

Private Sub btn_BookID_Click()

    isIssue = True
    frm_ListOfStudentsTransaction.Show 1

End Sub

Private Sub btn_Cancel_Click()

    Unload Me

End Sub

Private Sub btn_Issue_Click()

On Error Resume Next
    If cbx_Source.Text = "" Then
        
            MsgBox "Please specify the source of the book.", vbInformation, "Information"
        
        Else
            
            If cbx_Source.Text = "Reserved" Then
                
                    'Checking of existing of books and available no. RESERVED
                
                    RS_CheckBooks.Open "SELECT * FROM tbl_ReservedBook where ReservedBookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                    If RS_CheckBooks.RecordCount < 1 Then
                
                        MsgBox "Book ID was not found, no record found!", vbInformation, "Result"
                        RS_CheckBooks.Close
                        Exit Sub
                
                    Else
                
                        RS_CheckBooks.Close
                        RS_CheckBooks.Open "SELECT AvailNo FROM tbl_ReservedBook WHERE ReservedBookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                            If RS_CheckBooks(0) < 1 Then
                
                                MsgBox "Sorry no available copy of this book.", vbInformation, "Result"
                                RS_CheckBooks.Close
                                Exit Sub
                                
                            End If
                
                    End If
                    
                    'Update the availability of the book
                    
                    RS_BookUpdate.Open "SELECT * FROM tbl_ReservedBook WHERE ReservedBookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                    AvailNo = RS_BookUpdate.Fields("AvailNo")
                    IssuedNo = RS_BookUpdate.Fields("IssuedNo")
                    
                    With RS_BookUpdate
                    
                        ![AvailNo] = AvailNo - 1
                        ![IssuedNo] = IssuedNo + 1
                        .Update
                    
                    End With
                    
                    RS_BookUpdate.Close
                
            ElseIf cbx_Source.Text = "Circulation" Then
        
                    'Checking of existing of books and available no. CIRCULATION
                
                    RS_CheckBooks.Open "SELECT * FROM tbl_Book where BookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                    If RS_CheckBooks.RecordCount < 1 Then
                
                        MsgBox "Book ID was not found, no record found!", vbInformation, "Result"
                        RS_CheckBooks.Close
                        Exit Sub
                
                    Else
                
                        RS_CheckBooks.Close
                        RS_CheckBooks.Open "SELECT AvailNo FROM tbl_Book WHERE BookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                            If RS_CheckBooks(0) < 1 Then
                
                                MsgBox "Sorry no available copy of this book.", vbInformation, "Result"
                                RS_CheckBooks.Close
                                Exit Sub
                                
                            End If
                
                    End If
                    
                    'Update the availability of the book
                    
                    RS_BookUpdate.Open "SELECT * FROM tbl_Book WHERE BookID LIKE '" & Trim(txt_BookID.Text) & "'", CN3, adOpenStatic, adLockOptimistic
                
                    AvailNo = RS_BookUpdate.Fields("AvailNo")
                    IssuedNo = RS_BookUpdate.Fields("IssuedNo")
                    
                    With RS_BookUpdate
                    
                        ![AvailNo] = AvailNo - 1
                        ![IssuedNo] = IssuedNo + 1
                        .Update
                    
                    End With
                    
                    RS_BookUpdate.Close
                
            End If
                    
             'Checking if Student Exist and maximum book hold
                    
             RS_CheckStudent.Open "SELECT * FROM tbl_Students WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN3, adOpenStatic, adLockOptimistic
             
             If RS_CheckStudent.RecordCount < 1 Then
             
                MsgBox "Student no. was not found, no record found!", vbInformation, "Result"
                RS_CheckStudent.Close
                Exit Sub
                
             Else
      
                RS_CheckStudent.Close
                RS_CheckStudent.Open "SELECT BookinHand FROM tbl_Students WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN3, adOpenStatic, adLockOptimistic
           
                    If RS_CheckStudent(0) = maxhold Then
                    
                        MsgBox "You reach the maximum books allowed on hand.", vbInformation, "Result"
                        RS_CheckStudent.Close
                        Exit Sub
                    
                    End If
                    
            End If
            
            'Check if the book already barrowed
            
            RS_CheckIssue.Open "SELECT * FROM tbl_BorrowedBooks WHERE BookID LIKE '" & Trim(txt_BookID.Text) & "' AND StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "' AND StatusReturned =" & False & "", CN3, adOpenStatic, adLockOptimistic
            
                If RS_CheckIssue.RecordCount > 0 Then
                
                    MsgBox "This book was already barrowed. Barrower cannot barrow the same book.", vbInformation, "Information"
                    RS_CheckIssue.Close
                    Exit Sub
                    
                End If
                
            'Add book to Barrowed Book
            
            RS_Issue.Open "SELECT * FROM tbl_BorrowedBooks", CN3, adOpenStatic, adLockOptimistic
            
                With RS_Issue
                
                    .AddNew
                    ![StudentNo] = txt_StudentNo.Text
                    ![BookID] = txt_BookID.Text
                    ![DateBarrowed] = dtp_DateIssued.Value
                    ![DateReturned] = dtp_DateReturned.Value
                    ![BookSource] = cbx_Source.Text
                    ![Fines] = "0"
                    ![StatusReturned] = False
                    .Update
                
                End With
             
            'Update the Book in Hand
                
            RS_BookinHand.Open "SELECT * FROM tbl_Students WHERE StudentNo LIKE '" & Trim(txt_StudentNo.Text) & "'", CN3, adOpenStatic, adLockOptimistic
            
                BookinHand = RS_BookinHand.Fields("BookinHand")
                
                With RS_BookinHand
                
                    ![BookinHand] = BookinHand + 1
                    .Update
                
                End With
                
                RS_BookinHand.Close
                
                MsgBox "Book was issued successfully", vbInformation, "Information"
                
                RS_Issue.Close
                RS_BookIn.Close
                
                If MsgBox("Do you want to issue another book?", vbQuestion + vbYesNo, "Confirmation") = vbYes Then
                
                    cbx_Source.Text = ""
                    txt_BookID.Text = "BID-"
                    txt_BookTitle.Text = ""
                    txt_StudentNo.Text = ""
                    txt_StudentName.Text = ""
                    dtp_DateIssued.Value = Date
                    dtp_DateReturned.Value = Date
                    dtp_DateReturned.Day = dtp_DateReturned.Day + dayslimit
                    
                Else
                
                    Unload Me
                    
                End If
            
    End If

End Sub

Private Sub btn_Reset_Click()

    On Error Resume Next
    
    cbx_Source.Text = ""
    txt_BookID.Text = "BID-"
    txt_BookTitle.Text = ""
    txt_StudentNo.Text = ""
    txt_StudentName.Text = ""
    dtp_DateIssued.Value = Date
    dtp_DateReturned.Value = Date
    dtp_DateReturned.Day = dtp_DateReturned.Day + dayslimit

End Sub

Private Sub btn_StudentNo_Click()
    
    If cbx_Source.Text = "" Then
    
        MsgBox "Please specify the source of the book.", vbInformation, "Information"
    
    Else
    
        isIssue = True
        frm_ListOfBooksTransaction.Show 1
    
    End If

End Sub

Private Sub Form_Activate()

    'On Error Resume Next
    
    
    
    frm_Main.toolbar_Menu.Buttons(8).Value = tbrPressed
    frm_Main.StatusBar1.Panels(2).Text = "Active Form: Book Transactions - Issue Book"

End Sub

Private Sub Form_Load()
    
    dtp_DateIssued.Value = Date
    dtp_DateReturned.Value = Date
    dtp_DateReturned.Value = dtp_DateReturned.Value + dayslimit
    btn_StudentNo.Caption = ""
    btn_BookID.Caption = ""
    
    btn_StudentNo.IconAlign = isbCenter
    btn_BookID.IconAlign = isbCenter
    
    CN3.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & DB_LOCATION & ";Persist Security Info=False;Jet OLEDB:Database Password=admin"

End Sub

Private Sub Image2_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)

    frm_Main.toolbar_Menu.Buttons(8).Value = tbrUnpressed
    frm_Main.StatusBar1.Panels(2).Text = getPath
    Set CN3 = Nothing
    
    frm_Main.RESET_STATUS
    frm_Main.RESTORE_BUTTON_VALUE

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

    RS_BookIn.Open "SELECT * FROM tbl_Book", CN3, adOpenStatic, adLockOptimistic
          
    
    With RS_BookIn
    
        .MoveFirst
        .Find "BookID LIKE '" & Trim(txt_BookID.Text) & "'"
        
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
