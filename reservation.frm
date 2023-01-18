VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCT2.Ocx"
Begin VB.Form reservation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reservation"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14640
   BeginProperty Font 
      Name            =   "Calibri"
      Size            =   11.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "reservation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   14640
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Proceed"
      Height          =   375
      Left            =   11280
      TabIndex        =   34
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox Text10 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10080
      MaxLength       =   1
      TabIndex        =   33
      Top             =   3960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Yes"
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
      Left            =   10080
      TabIndex        =   26
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Caption         =   "Passenger Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   2880
      TabIndex        =   25
      Top             =   4680
      Visible         =   0   'False
      Width           =   8415
      Begin VB.CommandButton Command5 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   6120
         TabIndex        =   36
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Book Ticket"
         Height          =   375
         Left            =   6120
         TabIndex        =   35
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   3840
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2880
         TabIndex        =   30
         Text            =   "Text9"
         Top             =   960
         Width           =   615
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   600
         TabIndex        =   28
         Text            =   "Text8"
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Note: Tickets once reserved cannot be exchanged or edited"
         Height          =   855
         Left            =   5520
         TabIndex        =   37
         Top             =   1560
         Width           =   2655
         WordWrap        =   -1  'True
      End
      Begin VB.Line Line1 
         X1              =   5400
         X2              =   5400
         Y1              =   240
         Y2              =   2520
      End
      Begin VB.Label Label15 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label14 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3000
         TabIndex        =   29
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "Name Of Passenger"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   27
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "E-Ticket"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   480
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
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
         Left            =   2280
         TabIndex        =   7
         Top             =   1920
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Submit"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   2
         Text            =   "(Select)"
         Top             =   600
         Width           =   2175
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1200
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         Format          =   155451393
         CurrentDate     =   43183
      End
      Begin VB.Label Label4 
         Caption         =   "Select Date"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Train Name -"
         Height          =   375
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Train Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5400
      TabIndex        =   8
      Top             =   1080
      Visible         =   0   'False
      Width           =   8775
      Begin VB.TextBox Text7 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Text            =   "Text7"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7560
         TabIndex        =   14
         Text            =   "Text6"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   13
         Text            =   "Text5"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4680
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Fare"
         Height          =   255
         Left            =   3480
         TabIndex        =   22
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Departure"
         Height          =   255
         Left            =   7560
         TabIndex        =   21
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label10 
         Caption         =   "Arrival"
         Height          =   255
         Left            =   6120
         TabIndex        =   20
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Distance"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "From"
         Height          =   255
         Left            =   1680
         TabIndex        =   16
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Train Number"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "How many seats you want to reserve -"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Want To Reserve -"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "RESERVATION FOR LUXURIOUS TRAINS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   0
      Picture         =   "reservation.frx":FD31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "reservation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser1 As New ADODB.Recordset
Dim rsUser2 As New ADODB.Recordset
Dim rsUser3 As New ADODB.Recordset
Dim c As Integer
Private Sub Check1_Click()
Label6.Visible = True: Text10.Visible = True: Command3.Visible = True
Check1.Enabled = False
End Sub

Private Sub Command1_Click()
If Combo1.Text = "(Select)" Then
MsgBox "SELECT ANY TRAIN", vbInformation, "Railway Reservation System"
Exit Sub
End If
If DTPicker1.Value <= Date Then
MsgBox "PICK ANOTHER DATE", vbInformation, "Railway Reservation System"
Exit Sub
End If
s1 = Combo1.Text
Frame2.Caption = s1
Dim sSQL As String
sSQL = "SELECT * from trains where TrainName='" & s1 & "'"
OpenRecordSet rsUser2, sSQL
Text1 = rsUser2.Fields(0): Text2 = rsUser2.Fields(2): Text3 = rsUser2.Fields(3): Text4 = rsUser2.Fields(4)
Text5 = rsUser2.Fields(5): Text6 = rsUser2.Fields(6): Text7 = rsUser2.Fields(7)
Frame2.Visible = True: Label3.Visible = True: Check1.Visible = True
End Sub

Private Sub Command2_Click()
Combo1.Text = "(Select)"
DTPicker1.Value = Date
End Sub

Private Sub Command3_Click()
If Asc(Text10.Text) = 49 Or Asc(Text10.Text) = 50 Or Asc(Text10.Text) = 51 Then
c = Text10.Text
Command3.Enabled = False
Text10.Enabled = False
Call book
Else
MsgBox "Maximum 3 seats can be Booked at Once", vbInformation, "Railway Reservation System"
Exit Sub
End If
End Sub

Private Sub Command4_Click()
For i = 0 To c - 1
If Text8(i) = "" Or Text9(i) = "" Or Combo2(i).Text = "" Then
MsgBox "FILL ALL FIELDS", vbInformation, "Railway Reservation System"
Exit Sub
End If
Next
Max = 200000: Min = 100000
Randomize
pnr = Int((Max - Min + 1) * Rnd) + Min
Dim sSQL As String
sSQL = "SELECT * FROM ticket"
OpenRecordSet rsUser3, sSQL
While Not rsUser3.EOF
If pnr = rsUser3.Fields(0).Value Then
rsUser3.MoveFirst
Randomize
pnr = Int((Max - Min + 1) * Rnd) + Min
End If
rsUser3.MoveNext
Wend
Dim num As Long
If rsUser2.Fields(0) = 12267 Then
num = 182000 * c
ElseIf rsUser2.Fields(0) = 13108 Then
num = 575000 * c
ElseIf rsUser2.Fields(0) = 14154 Then
num = 318000 * c
ElseIf rsUser2.Fields(0) = 15789 Then
num = 6804 * c
End If
For i = 0 To c - 1
sSQL = "INSERT INTO ticket VALUES ('" & pnr & "','" & Text8(i) & "','" & Text9(i) & "','" & Combo2(i).Text & "','" & Text1 & "','" & Frame2.Caption & "','" & Text2 & "','" & Text3 & "','" & Text4 & "','" & Text5 & "','" & Text6 & "','" & num & "','" & un & "','" & DTPicker1.Value & "')"
CnnStr.Execute (sSQL)
Next
ticket.pnr = pnr
Unload Me
pay.Show
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim sSQL As String
sSQL = "SELECT * FROM trains"
OpenRecordSet rsUser1, sSQL
While Not rsUser1.EOF
Combo1.AddItem rsUser1.Fields("TrainName").Value
rsUser1.MoveNext
Wend
Image1.Height = reservation.Height
Image1.Width = reservation.Width
Combo2(0).AddItem "MALE"
Combo2(0).AddItem "FEMALE"
DTPicker1.Value = Date
End Sub

Private Sub form_unload(cancel As Integer)
For i = 1 To c - 1
Unload Text8(i)
Unload Text9(i)
Unload Combo2(i)
Next
c = 0
End Sub

Private Sub book()
Frame3.Visible = True
Dim i As Integer
  Text8(0).Text = "": Text9(0).Text = ""
  For i = 1 To c - 1
    Load Text8(i)
    With Text8(i)
      .Text = ""
      .Visible = True
      .Top = Text8(i - 1).Top + 550
    End With
  Next i
  For i = 1 To c - 1
    Load Text9(i)
    With Text9(i)
      .Text = ""
      .Visible = True
      .Top = Text9(i - 1).Top + 550
    End With
  Next i
  For i = 1 To c - 1
    Load Combo2(i)
    With Combo2(i)
      .AddItem "MALE"
      .AddItem "FEMALE"
      .Text = ""
      .Visible = True
      .Top = Text8(i - 1).Top + 550
    End With
  Next i
End Sub
