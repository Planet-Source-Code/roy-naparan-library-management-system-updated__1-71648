VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{AE48887E-94B0-429D-9EB0-D65524AD13B3}#1.0#0"; "isCoolButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frm_Login 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   8475
   ClientLeft      =   4260
   ClientTop       =   4215
   ClientWidth     =   11925
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Moveable        =   0   'False
   Picture         =   "frm_Login.frx":0000
   ScaleHeight     =   8475
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Loader_Timer 
      Left            =   9585
      Top             =   3015
   End
   Begin MSComctlLib.ProgressBar pb_Loader 
      Height          =   70
      Left            =   8190
      TabIndex        =   9
      Top             =   4410
      Visible         =   0   'False
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   132
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.ComboBox cmb_Usertype 
      Height          =   315
      ItemData        =   "frm_Login.frx":26E49
      Left            =   2655
      List            =   "frm_Login.frx":26E59
      TabIndex        =   0
      Text            =   "Administrator"
      Top             =   3105
      Width           =   2445
   End
   Begin VB.TextBox txt_Username 
      BackColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2655
      TabIndex        =   1
      Text            =   "admin"
      Top             =   3510
      Width           =   2445
   End
   Begin VB.TextBox txt_Password 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2655
      PasswordChar    =   "•"
      TabIndex        =   2
      Text            =   "admin"
      Top             =   3915
      Width           =   2445
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1980
      Top             =   4950
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin isCoolButton.isButton btn_Ok 
      Height          =   330
      Left            =   2655
      TabIndex        =   3
      Top             =   4320
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_Login.frx":26E91
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
      Height          =   330
      Left            =   3915
      TabIndex        =   4
      Top             =   4320
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   582
      Icon            =   "frm_Login.frx":26EAD
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "User Type:"
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   1665
      TabIndex        =   8
      Top             =   3195
      Width           =   825
   End
   Begin VB.Label lbl_Username 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Username:"
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   1395
      TabIndex        =   7
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label lbl_Password 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H80000009&
      Height          =   240
      Left            =   1620
      TabIndex        =   5
      Top             =   4005
      Width           =   870
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Connection Settings"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   240
      Left            =   3060
      MouseIcon       =   "frm_Login.frx":26EC9
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   4995
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   2655
      Picture         =   "frm_Login.frx":2701B
      Top             =   4950
      Width           =   240
   End
End
Attribute VB_Name = "frm_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   x As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 7.0"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long

Private RS_UserLog As New ADODB.Recordset
Private RS_AdminLog As New ADODB.Recordset
Private UserRowIndex As Integer
Private AdminRowIndex As Integer

Dim RS As New ADODB.Recordset

Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, Counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)




    ObjectRegion = CreateRoundRectRgn(72 * ScaleX * 15 / STPPX + OffsetX, 89 * ScaleY * 15 / STPPY + OffsetY, 711 * ScaleX * 15 / STPPX + OffsetX, 461 * ScaleY * 15 / STPPY + OffsetY, 62 * ScaleX * 15 / STPPX, 62 * ScaleY * 15 / STPPY)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion

    ObjectRegion = CreateRoundRectRgn(239 * ScaleX * 15 / STPPX + OffsetX, 183 * ScaleY * 15 / STPPY + OffsetY, 451 * ScaleX * 15 / STPPX + OffsetX, 328 * ScaleY * 15 / STPPY + OffsetY, 212 * ScaleX * 15 / STPPX, 144 * ScaleY * 15 / STPPY)
    nRet = CombineRgn(HolderRegion, ResultRegion, ResultRegion, RGN_COPY)
    nRet = CombineRgn(ResultRegion, HolderRegion, ObjectRegion, 2)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function



Private Sub cmb_Usertype_Click()

      If cmb_Usertype.Text = "Student" Then
    
        lbl_Username.Caption = "Student No.:"
        txt_Password.Visible = False
        lbl_Password.Visible = False
        
    Else
    
        lbl_Username.Caption = "Username:"
        txt_Password.Visible = True
        lbl_Password.Visible = True
    
    End If

End Sub

Private Sub Command1_Click()

End Sub



Private Sub cmb_Usertype_GotFocus()

    cmb_Usertype.BackColor = &HFBE6FB

End Sub

Private Sub cmb_Usertype_LostFocus()

    cmb_Usertype.BackColor = &H80000005

End Sub

Private Sub Form_Load()
 
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hwnd, CreateFormRegion(1, 1, 0, 0), True)

    
    welcometime = 2500
    Welcome = True
    
    If If_File_Exists(App.Path & "\DBPath.dat") = True Then
    If If_File_Exists(App.Path & "\DBPath.dat") = True Then
    
    If validDB(getPath) = True Then
    
        DB_LOCATION = getPath: Exit Sub
        
    Else
    
        MsgBox "The database location has been change or the database has been renamed. Please locate again the database.", vbExclamation, "Error Connection"
           
    End If
    End If
    End If

    CommonDialog1.DialogTitle = "Locate the CLMS Database File"
    CommonDialog1.Filter = "CLMS Database (db_CLMS.mdb)|*.mdb"
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
    
End Sub
Private Sub Form_Unload(Cancel As Integer)

    DeleteObject ResultRegion

End Sub
Private Sub btn_Cancel_Click()
    
    End

End Sub
Private Sub btn_Ok_Click()


        On Error Resume Next
    
        pb_Loader.Visible = True
        Loader_Timer.Interval = 5
        
End Sub
Private Function getAdminIndex() As String
On Error Resume Next
Dim t As String
Open App.Path & "\IsActive.dat" For Input As #1
    Input #1, t
Close #1
getAdminIndex = Trim$(t)
t = vbNullString
End Function
Private Function getUserIndex() As String
On Error Resume Next
Dim t As String
Open App.Path & "\UserLog.dat" For Input As #1
    Input #1, t
Close #1
getUserIndex = Trim$(t)
t = vbNullString
End Function
Private Sub savePathUser()
On Error Resume Next
Open App.Path & "\UserLog.dat" For Output As #1
    Print #1, Trim(UserRowIndex)
Close #1
End Sub
Private Sub saveAdminUser()
On Error Resume Next
Open App.Path & "\IsActive.dat" For Output As #1
    Print #1, Trim(AdminRowIndex)
Close #1
End Sub

Private Sub Label4_Click()
    
    On Error Resume Next
               
    If If_File_Exists(App.Path & "\DBPath.dat") = True Then
    Dim rep As Integer
    rep = MsgBox("Do you want to change Database connection?", vbQuestion + vbYesNo, "Connection Settings")
    If rep = vbYes Then
        
    Else
        Exit Sub
    End If
    End If
    
        CommonDialog1.ShowOpen
        CommonDialog1.CancelError = False
        
        If CommonDialog1.FileName <> "" Then

            If validDB(CommonDialog1.FileName) = True Then
         
               DB_LOCATION = CommonDialog1.FileName
               MsgBox "You are now connected to " & DB_LOCATION & ".", vbInformation
               Kill (App.Path & "\DBPath.dat")
               Call savePath
               
               
                
            Else
        
                MsgBox "The selected file is not a valid database for this application.", vbExclamation, "Invalid"
            
            End If
    
        End If
        
        If validDB(getPath) = True Then
               DB_LOCATION = getPath
               End If

End Sub
Private Sub savePath()
On Error Resume Next
Open App.Path & "\DBPath.dat" For Output As #1
    Print #1, CommonDialog1.FileName
Close #1
End Sub
Private Function getPath() As String
On Error Resume Next
Dim t As String
Open App.Path & "\DBPath.dat" For Input As #1
    Input #1, t
Close #1
getPath = Trim$(t)
t = vbNullString
End Function



Private Sub Loader_Timer_Timer()

    On Error Resume Next
    pb_Loader.Value = pb_Loader.Value + 1
    
        If pb_Loader.Value = pb_Loader.Max Then
                
                    
                    Loader_Timer.Interval = 0
                    pb_Loader.Value = 0
                    pb_Loader.Visible = False
                    
                    Set CN = New ADODB.Connection
                    Set RS = New ADODB.Recordset
                    
                    Call connect_to_db
                    
        If cmb_Usertype.Text = "Student" Then

              'Check users in 1 PC

                If getUserIndex = "" Then

                Else

                        MsgBox "Only 1 user can use the system in a single computer.", vbInformation, "Information"
                        Exit Sub

                End If

              RS.Open "SELECT * from tbl_Students where StudentNo = '" & Replace(Trim(txt_Username.Text), "'", "") & "'", CN, adOpenDynamic, adLockOptimistic

              If RS.EOF Then

                   MsgBox "Invalid student number.Please try again.", vbExclamation, "Information"
                   txt_Username.Text = ""
                   txt_Username.SetFocus

              Else


                  MsgBox "Welcome " & RS.Fields("StudentName") & " to Libarary Management System.", vbInformation, "Information"
                  frm_Main.StatusBar1.Panels(4).Text = RS.Fields("StudentName") & " " & "(" & "Student" & ")"

                    RS_UserLog.Open "SELECT * FROM tbl_UserLog", CN, adOpenStatic, adLockOptimistic

                        With RS_UserLog

                            .AddNew
                            ![UserName] = Trim(txt_Username.Text)
                            ![UserType] = cmb_Usertype.Text
                            ![FullName] = RS.Fields("StudentName")
                            ![DateLogIn] = Date
                            ![TimeLogIn] = Time
                            ![DateLogOut] = Date
                            ![TimeLogOut] = Time
                            .Update

                                UserRowIndex = ![UserLogNo]
                                Call savePathUser

                            .Close

                        End With


                  RS.Close
                  Set CN = Nothing
                  Unload Me

                  frm_Main.mnu_UserManagement.Visible = False
                  frm_Main.mnu_Records(2).Visible = False
                  frm_Main.mnu_Transaction(3).Visible = False
                  frm_Main.mnu_Report(4).Visible = False
                  frm_Main.mnu_BackupDatabase.Visible = False
                  frm_Main.mnu_SystemSettings.Visible = False
                  frm_Main.mnu_Sep3.Visible = False
                  frm_Main.mnu_Sep4.Visible = False

                  frm_Main.toolbar_Menu.Buttons(2).Visible = False
                  frm_Main.toolbar_Menu.Buttons(3).Visible = False
                  frm_Main.toolbar_Menu.Buttons(4).Visible = False
                  frm_Main.toolbar_Menu.Buttons(5).Visible = False
                  frm_Main.toolbar_Menu.Buttons(6).Visible = False
                  frm_Main.toolbar_Menu.Buttons(7).Visible = False
                  frm_Main.toolbar_Menu.Buttons(8).Visible = False
                  frm_Main.toolbar_Menu.Buttons(9).Visible = False
                  frm_Main.toolbar_Menu.Buttons(10).Visible = False
                  frm_Main.toolbar_Menu.Buttons(11).Visible = False

                  'frm_Main.toolbar_Menu.Buttons(12).Visible = True
                  'frm_Main.toolbar_Menu.Buttons(13).Visible = True

                  frm_Main.toolbar_Menu.Buttons(14).Visible = False
                  frm_Main.toolbar_Menu.Buttons(15).Visible = False
                  frm_Main.toolbar_Menu.Buttons(16).Visible = False
                  frm_Main.toolbar_Menu.Buttons(17).Visible = False
                  frm_Main.toolbar_Menu.Buttons(18).Visible = False


                     frm_Main.Show
                     DoEvents
                     If (Welcome = True) Then
                     Load frm_Welcome
                     frm_Welcome.Show
                     End If


              End If

        ElseIf cmb_Usertype.Text = "Administrator" Then

              'Check users in 1 PC

                If getAdminIndex = "" Then

                Else

                    MsgBox "Only 1 user can use the system in a single computer.", vbInformation, "Information"
                    Exit Sub

                End If

              'Checking if the user already Login

              RS_AdminLog.Open "SELECT * FROM tbl_User WHERE Username LIKE '" & Trim(txt_Username.Text) & "' AND Active =" & True & "", CN, adOpenStatic, adLockOptimistic

                If RS_AdminLog.RecordCount >= 1 Then

                    MsgBox "Username is already used by the other user.", vbCritical, "Login Failed"
                    txt_Password.Text = ""
                    txt_Username.SetFocus
                    RS_AdminLog.Close
                    Exit Sub

                End If


              'User Login
              
              RS.Open "SELECT * from tbl_User where USERNAME = '" & Replace(Trim(txt_Username.Text), "'", "") & "' and PASSWORD = '" & Replace(Trim(Encode(txt_Password.Text)), "'", "") & "'", CN, adOpenStatic, adLockOptimistic
              

              If RS.EOF Then
              
                   MsgBox "Invalid username or password.Please try again.", vbExclamation, "Information"
                   txt_Password.Text = ""
                   txt_Username.Text = ""
                   txt_Username.SetFocus
                   
                    
             Else
                  
                          frm_Main.mnu_UserManagement.Visible = True
                          frm_Main.mnu_Records(2).Visible = True
                          frm_Main.mnu_Transaction(3).Visible = True
                          frm_Main.mnu_Report(4).Visible = True
                          frm_Main.mnu_BackupDatabase.Visible = True
                          frm_Main.mnu_SystemSettings.Visible = True
                          frm_Main.mnu_Sep3.Visible = True
                          frm_Main.mnu_Sep4.Visible = True
                          
                          frm_Main.toolbar_Menu.Buttons(2).Visible = True
                          frm_Main.toolbar_Menu.Buttons(3).Visible = True
                          frm_Main.toolbar_Menu.Buttons(4).Visible = True
                          frm_Main.toolbar_Menu.Buttons(5).Visible = True
                          frm_Main.toolbar_Menu.Buttons(6).Visible = True
                          frm_Main.toolbar_Menu.Buttons(7).Visible = True
                          frm_Main.toolbar_Menu.Buttons(8).Visible = True
                          frm_Main.toolbar_Menu.Buttons(9).Visible = True
                          frm_Main.toolbar_Menu.Buttons(10).Visible = True
                          frm_Main.toolbar_Menu.Buttons(11).Visible = True
                          
                          'frm_Main.toolbar_Menu.Buttons(12).Visible = False
                          'frm_Main.toolbar_Menu.Buttons(13).Visible = False
                          
                          frm_Main.toolbar_Menu.Buttons(14).Visible = True
                          frm_Main.toolbar_Menu.Buttons(15).Visible = True
                          frm_Main.toolbar_Menu.Buttons(16).Visible = True
                          frm_Main.toolbar_Menu.Buttons(17).Visible = True
                          frm_Main.toolbar_Menu.Buttons(18).Visible = True
                
                      
                          MsgBox "Welcome " & RS.Fields("FULLNAME") & " to Libarary Management System.", vbInformation, "Information"
                          frm_Main.StatusBar1.Panels(4).Text = RS.Fields("FULLNAME") & " " & "(" & RS.Fields("UserType") & ")"
                          
                          RS_UserLog.Open "SELECT * FROM tbl_UserLog", CN, adOpenStatic, adLockOptimistic
                          
                            With RS_UserLog
                                      
                                  .AddNew
                                  ![UserName] = Trim(txt_Username.Text)
                                  ![UserType] = cmb_Usertype.Text
                                  ![FullName] = RS.Fields("FULLNAME")
                                  ![DateLogIn] = Date
                                  ![TimeLogIn] = Time
                                  ![DateLogOut] = Date
                                  ![TimeLogOut] = Time
                                  .Update
                                  
                                    UserRowIndex = ![UserLogNo]
                                    Call savePathUser
                                  
                                  .Close
                                      
                            End With
                            
                            With RS
                            
                                ![Active] = True
                                .Update
                                    
                                    AdminRowIndex = ![RowIndex]
                                    Call saveAdminUser
                                
                            End With
                            
                            
                          
                          RS.Close
                          Set CN = Nothing
                          Unload Me
                          
                          
                          
                          frm_Main.Show
                           
                             DoEvents
                             If (Welcome = True) Then
                             Load frm_Welcome
                             frm_Welcome.Show
                             End If
               
                      End If
              
        End If
                    
        End If

End Sub

Private Sub txt_Password_GotFocus()

    txt_Password.BackColor = &HFBE6FB

End Sub

Private Sub txt_Password_LostFocus()

    txt_Password.BackColor = &H80000005

End Sub

Private Sub txt_Username_GotFocus()

    txt_Username.BackColor = &HFBE6FB

End Sub

Private Sub txt_Username_LostFocus()

    txt_Username.BackColor = &H80000005

End Sub
