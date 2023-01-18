VERSION 5.00
Begin VB.Form about 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About Railway Reservation System"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7155
   Icon            =   "about.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Back"
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
      Left            =   5640
      TabIndex        =   3
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   240
      X2              =   6960
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      X1              =   1440
      X2              =   6960
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   1560
      TabIndex        =   2
      Top             =   2160
      Width           =   5295
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1560
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Railway Reservation System"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   0
      Picture         =   "about.frx":FD31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
End
Attribute VB_Name = "about"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label2.Caption = "1.0.0" & vbCrLf & "This is a small scale Railway Reservation System focusing on LUXURIOUS TRAINS ticketing."
Label3.Caption = "Front End: Visual Basic 6.0"
Label3.Caption = Label3.Caption & vbCrLf & "Back End: MS-ACCESS"
Label3.Caption = Label3.Caption & vbCrLf & "Created By: Sachin Chatterjee And Mohd. Laraib Ishtiyaq"
End Sub
