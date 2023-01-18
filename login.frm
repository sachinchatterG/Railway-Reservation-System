VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   9240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17535
   Icon            =   "login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "login.frx":FD31
   ScaleHeight     =   9240
   ScaleWidth      =   17535
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "REGISTER"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   720
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FF8080&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   3015
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   7560
      ScaleHeight     =   6855
      ScaleWidth      =   6135
      TabIndex        =   0
      Top             =   1560
      Width           =   6135
      Begin VB.Frame Frame2 
         Height          =   6375
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   5655
         Begin VB.TextBox Text9 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2760
            TabIndex        =   25
            Top             =   4200
            Width           =   2415
         End
         Begin VB.TextBox Text8 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            IMEMode         =   3  'DISABLE
            Left            =   2760
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   3360
            Width           =   2415
         End
         Begin VB.TextBox Text7 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   2760
            PasswordChar    =   "*"
            TabIndex        =   22
            Top             =   2400
            Width           =   2415
         End
         Begin VB.TextBox Text6 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   21
            Top             =   1680
            Width           =   2415
         End
         Begin VB.TextBox Text5 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2760
            TabIndex        =   20
            Top             =   960
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            TabIndex        =   19
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label Label15 
            Caption         =   "Already In Use!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3120
            TabIndex        =   28
            Top             =   2040
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label Label14 
            Caption         =   "Incorrect Password!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3000
            TabIndex        =   27
            Top             =   3840
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Label Label13 
            Caption         =   "Alphabets Only!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   3000
            TabIndex        =   26
            Top             =   600
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.Image Image5 
            Height          =   1695
            Left            =   2760
            Picture         =   "login.frx":8CB41
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   1815
         End
         Begin VB.Image Image4 
            Height          =   1695
            Left            =   720
            Picture         =   "login.frx":9D5DD
            Stretch         =   -1  'True
            Top             =   4680
            Width           =   1935
         End
         Begin VB.Label Label12 
            Caption         =   "Email"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   24
            Top             =   4200
            Width           =   1455
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Confirm Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   18
            Top             =   3120
            Width           =   1695
         End
         Begin VB.Label Label10 
            Caption         =   "Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   17
            Top             =   2400
            Width           =   2055
         End
         Begin VB.Label Label9 
            Caption         =   "Username"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   16
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label Label8 
            Caption         =   "Last Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   15
            Top             =   960
            Width           =   1695
         End
         Begin VB.Label Label7 
            Caption         =   "First Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   14
            Top             =   240
            Width           =   1935
         End
      End
      Begin VB.Frame Frame1 
         ForeColor       =   &H00000000&
         Height          =   6375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5655
         Begin VB.TextBox Text3 
            Alignment       =   2  'Center
            BackColor       =   &H00C0FFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1920
            TabIndex        =   10
            Top             =   4320
            Width           =   1815
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            IMEMode         =   3  'DISABLE
            Left            =   2280
            PasswordChar    =   "*"
            TabIndex        =   7
            Top             =   2040
            Width           =   3015
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2280
            TabIndex        =   5
            Top             =   840
            Width           =   3015
         End
         Begin VB.Label Label6 
            Caption         =   "Try Again!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3960
            TabIndex        =   12
            Top             =   4440
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Invalid  Credentials!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   1680
            TabIndex        =   11
            Top             =   240
            Visible         =   0   'False
            Width           =   2055
         End
         Begin VB.Image Image3 
            Height          =   1335
            Left            =   3000
            Picture         =   "login.frx":B34CF
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Image Image2 
            Height          =   1365
            Left            =   1320
            Picture         =   "login.frx":C3F6B
            Stretch         =   -1  'True
            Top             =   5040
            Width           =   1455
         End
         Begin VB.Image Image1 
            Height          =   585
            Left            =   4320
            Picture         =   "login.frx":C5C3A
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   645
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            FillStyle       =   5  'Downward Diagonal
            Height          =   855
            Left            =   1440
            Shape           =   4  'Rounded Rectangle
            Top             =   3240
            Width           =   2775
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Persia BT"
               Size            =   26.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   1440
            TabIndex        =   9
            Top             =   3360
            Width           =   2895
         End
         Begin VB.Label Label3 
            Caption         =   "CAPTCHA"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   8
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "PASSWORD"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   6
            Top             =   2040
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "USERNAME"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   4
            Top             =   840
            Width           =   1815
         End
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   3000
      Top             =   4920
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\dell\Desktop\vb project\login.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\dell\Desktop\vb project\login.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Image Image6 
      Height          =   9255
      Left            =   0
      Picture         =   "login.frx":CA27A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20055
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Maxvalue, Minvalue, a, b, c, d As Integer
Dim rsUser As New ADODB.Recordset
Dim s As String
Private Sub Command1_Click()
Command1.BackColor = &HFF8080
Command2.BackColor = &H8000000F
Frame2.Visible = False
Call Image3_Click
End Sub

Private Sub Command2_Click()
Command1.BackColor = &H8000000F
Command2.BackColor = &HFF8080
Frame2.Visible = True
Call Image5_Click
End Sub

Private Sub Form_Load()
Maxvalue = 90: Minvalue = 65: s = ""
Call random
Image6.Height = Screen.Height
Image6.Width = Screen.Width
End Sub

Private Sub Image1_Click()
Call random
End Sub

Private Sub random()
Randomize
a = Int((Maxvalue - Minvalue + 1) * Rnd) + Minvalue
Randomize
b = Int((Maxvalue - Minvalue + 1) * Rnd) + Minvalue
Randomize
c = Int((Maxvalue - Minvalue + 1) * Rnd) + Minvalue
Randomize
d = Int((Maxvalue - Minvalue + 1) * Rnd) + Minvalue
Label4.Caption = Chr(a) + " " + Chr(b) + " " + Chr(c) + " " + Chr(d)
s = Chr(a) + Chr(b) + Chr(c) + Chr(d)
End Sub
Private Sub Image2_Click()
 n1 = 0: n2 = 0
If s = Text3.Text Then
Label6.Visible = False: n1 = 1
Else: n1 = 0
Text3.Text = ""
Label6.Visible = True
Call random
End If
a = isValidUser(Text1.Text, Text2.Text)
If a Then
n2 = 1
un = Text1.Text
Label5.Visible = False
Else: n2 = 0
Label5.Visible = True
Text1.Text = "": Text2.Text = "": Text3.Text = ""
Call random
End If
If n1 = 1 And n2 = 1 Then
railway.Label1.Caption = "WELCOME  " + UCase(rsUser.Fields("First Name").Value)
Unload login
railway.Show
Dim s4 As String
s4 = "DELETE from ticket where day<'" & DateValue(Now) & "'"
CnnStr.Execute (s4)
Else
Exit Sub
End If
End Sub

Public Function isValidUser(sUserName As String, sPassword As String) As Boolean
Dim sSQL As String
Dim bReturnValue As Boolean
sSQL = "SELECT * FROM tbluserdetails WHERE Username ='" & sUserName & "' AND Password = '" & sPassword & "'"
OpenRecordSet rsUser, sSQL
If rsUser.RecordCount > 0 Then
bReturnValue = True
Else
bReturnValue = False
End If
isValidUser = bReturnValue
End Function

Private Sub Image3_Click()
Text1.Text = "": Text2.Text = "": Text3.Text = ""
Label5.Visible = False: Label6.Visible = False
Call random
End Sub

Private Sub Image4_Click()
Dim sSQL As String
sSQL = "SELECT * FROM tbluserdetails"
OpenRecordSet rsUser, sSQL
If Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text7.Text <> Text8.Text Then
MsgBox "FILL ALL FIELDS", vbInformation, "Railway Reservation System"
Exit Sub
End If
While Not rsUser.EOF
If Text6.Text = rsUser.Fields("Username").Value Then
Label15.Visible = True
Exit Sub
Else
rsUser.MoveNext
Label15.Visible = False
End If
Wend
un = Text6.Text
Call insert(Text4.Text, Text5.Text, Text6.Text, Text7.Text, Text9.Text)
railway.Label1.Caption = "WELCOME  " + UCase(Text4.Text)
Unload login
railway.Show
End Sub

Public Function insert(fname As String, lname As String, uname As String, pass As String, email As String)
Dim sSQL As String
sSQL = "INSERT INTO tbluserdetails VALUES ('" & fname & "','" & lname & "','" & email & "','" & uname & "','" & pass & "')"
CnnStr.Execute (sSQL)
End Function

Private Sub Image5_Click()
Text4.Text = "": Text5.Text = "": Text6.Text = "": Text7.Text = "": Text8.Text = "": Text9.Text = ""
Label13.Visible = False: Label14.Visible = False: Label15.Visible = False
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32 Then
Label13.Visible = False
Else
KeyAscii = 0
Label13.Visible = True
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 8 Or KeyAscii = 32 Then
Label13.Visible = False
Else
KeyAscii = 0
Label13.Visible = True
End If
End Sub

Private Sub Text8_LostFocus()
If Text7.Text <> Text8.Text Then
Label14.Visible = True
Text8.Text = ""
Else
Label14.Visible = False
End If
End Sub
