VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ticket 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ticket"
   ClientHeight    =   10440
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11925
   FillStyle       =   0  'Solid
   Icon            =   "ticket.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10440
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   5400
      TabIndex        =   25
      Top             =   9960
      Width           =   1095
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   120
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "frame"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   840
      TabIndex        =   12
      Top             =   5160
      Width           =   6975
      Begin VB.TextBox Text10 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   24
         Text            =   "Text10"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   22
         Text            =   "Text9"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   20
         Text            =   "Text8"
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   16
         Text            =   "Text6"
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Text            =   "Text5"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label13 
         Caption         =   "Departure"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Arrival"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "Fare"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Date Of Journey"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Passenger Details"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   6975
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4920
         TabIndex        =   11
         Text            =   "Text4"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3120
         TabIndex        =   10
         Text            =   "Text3"
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   480
         TabIndex        =   9
         Text            =   "Text2"
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   8
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Name Of Passenger"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Enabled         =   0   'False
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
      Left            =   4200
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Shape Shape4 
      BorderWidth     =   3
      Height          =   1935
      Left            =   720
      Top             =   7560
      Width           =   10575
   End
   Begin VB.Image Image3 
      Height          =   1815
      Left            =   720
      Picture         =   "ticket.frx":FD31
      Stretch         =   -1  'True
      Top             =   7680
      Width           =   10455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirmed"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   375
      Left            =   6240
      TabIndex        =   5
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Status -"
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
      Left            =   5040
      TabIndex        =   3
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Shape Shape3 
      Height          =   9135
      Left            =   480
      Top             =   480
      Width           =   10935
   End
   Begin VB.Shape Shape2 
      BorderWidth     =   3
      Height          =   5055
      Left            =   8040
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "PNR  Number -"
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   8040
      Picture         =   "ticket.frx":211B7
      Stretch         =   -1  'True
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   4
      Height          =   9375
      Left            =   360
      Top             =   360
      Width           =   11175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1200
      Picture         =   "ticket.frx":56978
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RAILWAY RESERVATION SYSTEM (E-TICKET)"
      DataMember      =   "&H00FFFFFF&"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   8055
   End
End
Attribute VB_Name = "ticket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser As New ADODB.Recordset
Dim c As Integer
Public pnr As Variant

Private Sub Command1_Click()
Command1.Visible = False
On Error GoTo cancel
cd1.ShowPrinter
Me.PrintForm
Command1.Visible = True
cancel:
End Sub

Private Sub Form_Load()
Dim sSQL As String
sSQL = "select * from ticket where PnrNumber='" & pnr & "'"
OpenRecordSet rsUser, sSQL
c = rsUser.RecordCount
Dim i As Integer
  Text2(0).Text = "": Text3(0).Text = "": Text4(0).Text = ""
  For i = 1 To c - 1
    Load Text2(i)
    With Text2(i)
      .Text = ""
      .Visible = True
      .Top = Text2(i - 1).Top + 550
    End With
    Frame1.Height = Frame1.Height + 550
    Load Text3(i)
    With Text3(i)
      .Text = ""
      .Visible = True
      .Top = Text3(i - 1).Top + 550
    End With
    Load Text4(i)
    With Text4(i)
      .Text = ""
      .Visible = True
      .Top = Text4(i - 1).Top + 550
    End With
  Next i
 n = 0
 Text1.Text = pnr
  Text5.Text = rsUser.Fields(13).Value
 Text6.Text = rsUser.Fields(11).Value
 Text7.Text = rsUser.Fields(6).Value
 Text8.Text = rsUser.Fields(7).Value
 Text9.Text = rsUser.Fields(9).Value
 Text10.Text = rsUser.Fields(10).Value
 Frame2.Caption = rsUser.Fields(4).Value & " - " & rsUser.Fields(5).Value
 While Not rsUser.EOF
 Text2(n).Text = rsUser.Fields(1).Value
 Text3(n).Text = rsUser.Fields(2).Value
 Text4(n).Text = rsUser.Fields(3).Value
 n = n + 1
 rsUser.MoveNext
 Wend
End Sub


