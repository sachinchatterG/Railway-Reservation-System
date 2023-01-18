VERSION 5.00
Begin VB.Form plan 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "My Plans"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   15735
   Icon            =   "plan.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   15735
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   15015
      Begin VB.TextBox Text7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   9720
         TabIndex        =   18
         Text            =   "Text7"
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox Text6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   8040
         TabIndex        =   17
         Text            =   "Text6"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   6600
         TabIndex        =   16
         Text            =   "Text5"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   3960
         TabIndex        =   15
         Text            =   "Text4"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   2760
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   1560
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         Enabled         =   0   'False
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
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "CANCEL TICKET"
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
         Index           =   0
         Left            =   13200
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "GENERATE TICKET"
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
         Index           =   0
         Left            =   11280
         TabIndex        =   10
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label9 
         Caption         =   "Date Of Journey"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9720
         TabIndex        =   9
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8160
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Train Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   4080
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Train Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   2760
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "PNR Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Serial Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   4
      Height          =   615
      Left            =   6360
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "NO PLANS YET"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5880
      TabIndex        =   5
      Top             =   1680
      Width           =   3135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "My Plans"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   2325
      Left            =   0
      Picture         =   "plan.frx":FD31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3480
   End
End
Attribute VB_Name = "plan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsUser1 As New ADODB.Recordset
Dim rsUser2 As New ADODB.Recordset

Private Sub Command1_Click(Index As Integer)
ticket.pnr = Text2(Index)
Unload Me
ticket.Show
End Sub

Private Sub Command2_Click(Index As Integer)
answer = MsgBox("ARE YOU SURE TO CANCEL THE TICKETS ?", vbQuestion + vbYesNo, "Railway Reservation System")
Dim s As String
s = "delete from ticket where PnrNumber='" & Text2(Index) & "'"
If answer = vbYes Then
CnnStr.Execute (s)
Unload Me
Load Me
Me.Show
End If
End Sub

Private Sub Form_Load()
Frame1.Caption = un & "'s Journeys Planned"
Dim s1 As String
s1 = "select distinct PnrNumber from ticket where uname='" & un & "'"
OpenRecordSet rsUser1, s1
c = rsUser1.RecordCount
If c = 0 Then
Frame1.Visible = False
Image1.Height = Me.Height: Image1.Width = Me.Width
Exit Sub
Else
Frame1.Visible = True
End If
 For i = 1 To c - 1
    Load Text1(i)
    With Text1(i)
      .Text = "": .Visible = True: .Top = Text1(i - 1).Top + 550
    End With
    Frame1.Height = Frame1.Height + 550
  Me.Height = Me.Height + 550
    Load Text2(i)
    With Text2(i)
      .Text = "": .Visible = True: .Top = Text2(i - 1).Top + 550
    End With
 
    Load Text3(i)
    With Text3(i)
      .Text = "": .Visible = True: .Top = Text3(i - 1).Top + 550
    End With
 
    Load Text4(i)
    With Text4(i)
      .Text = "": .Visible = True: .Top = Text4(i - 1).Top + 550
    End With
 
    Load Text5(i)
    With Text5(i)
      .Text = "": .Visible = True: .Top = Text5(i - 1).Top + 550
    End With
 
    Load Text6(i)
    With Text6(i)
      .Text = "": .Visible = True: .Top = Text6(i - 1).Top + 550
    End With

    Load Text7(i)
    With Text7(i)
      .Text = "": .Visible = True: .Top = Text7(i - 1).Top + 550
    End With
     Load Command1(i)
    With Command1(i)
      .Caption = "GENERATE TICKET": .Visible = True: .Top = Command1(i - 1).Top + 550
    End With
     Load Command2(i)
    With Command2(i)
      .Caption = "CANCEL TICKET": .Visible = True: .Top = Command2(i - 1).Top + 550
    End With
  Next i
  For i = 0 To c - 1
  Dim s As String
  s = "select * from ticket where PnrNumber='" & rsUser1.Fields(0).Value & "'"
  OpenRecordSet rsUser2, s
  Text1(i) = i + 1
  Text2(i) = rsUser2.Fields(0).Value
  Text3(i) = rsUser2.Fields(4).Value
  Text4(i) = rsUser2.Fields(5).Value
  Text5(i) = rsUser2.Fields(6).Value
  Text6(i) = rsUser2.Fields(7).Value
  Text7(i) = rsUser2.Fields(13).Value
  rsUser1.MoveNext
  Next
  Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub
