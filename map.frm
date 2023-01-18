VERSION 5.00
Begin VB.Form map 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "North India Railway Map"
   ClientHeight    =   9045
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12480
   Icon            =   "map.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   12480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9840
      TabIndex        =   0
      Top             =   7680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   9015
      Left            =   0
      Picture         =   "map.frx":FD31
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12495
   End
End
Attribute VB_Name = "Map"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub
