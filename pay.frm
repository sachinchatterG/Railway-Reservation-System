VERSION 5.00
Begin VB.Form pay 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pay"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5955
   Icon            =   "pay.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Generate Ticket"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Debit Card"
      Height          =   375
      Left            =   4080
      TabIndex        =   3
      Top             =   1200
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Credit Card"
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Net Bankimg"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Mode Of Payment"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   9645
      Left            =   0
      Picture         =   "pay.frx":FD31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "pay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
ticket.Show
End Sub

Private Sub Form_Load()
Image1.Width = pay.Width
Image1.Height = pay.Height
End Sub
