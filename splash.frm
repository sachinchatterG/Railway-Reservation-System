VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form splash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Welcome"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10335
   ControlBox      =   0   'False
   Icon            =   "splash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10335
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   1920
      TabIndex        =   3
      Top             =   840
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   855
      Left            =   600
      ScaleHeight     =   795
      ScaleWidth      =   8835
      TabIndex        =   1
      Top             =   1560
      Width           =   8895
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome To Railway Reservation"
         BeginProperty Font 
            Name            =   "Algerian"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   8535
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   7815
      Left            =   -240
      TabIndex        =   0
      Top             =   -120
      Width           =   10815
      ExtentX         =   19076
      ExtentY         =   13785
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate "F:\vb project\rail\t.gif"
End Sub

Private Sub Timer1_Timer()
If ProgressBar1.Value = 0 Then
Label1.Caption = "Loading..."
ElseIf ProgressBar1.Value = 30 Then
Label1.Caption = "Opening database..."
ElseIf ProgressBar1.Value = 50 Then
Label1.Caption = "checking connectivity..."
ElseIf ProgressBar1.Value = 70 Then
Label1.Caption = "welcome to railway reservation"
Picture1.BackColor = &HFFFFC0
End If
If ProgressBar1.Value < 100 Then
ProgressBar1.Value = ProgressBar1.Value + 2
Else
Unload splash
Load login
login.Visible = True
End If
End Sub


