VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm railway 
   BackColor       =   &H8000000C&
   Caption         =   "Railway Reservation System"
   ClientHeight    =   3015
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   4560
   Icon            =   "railway.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   9495
      Left            =   0
      ScaleHeight     =   9435
      ScaleWidth      =   4500
      TabIndex        =   4
      Top             =   1155
      Width           =   4560
      Begin VB.PictureBox Picture4 
         Height          =   9375
         Left            =   13560
         Picture         =   "railway.frx":FD31
         ScaleHeight     =   9315
         ScaleWidth      =   6555
         TabIndex        =   7
         Top             =   0
         Width           =   6615
         Begin VB.Timer Timer2 
            Enabled         =   0   'False
            Interval        =   4000
            Left            =   480
            Top             =   960
         End
         Begin VB.Shape Shape2 
            BorderWidth     =   3
            Height          =   735
            Left            =   6120
            Shape           =   4  'Rounded Rectangle
            Top             =   8160
            Width           =   3255
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "THE GOLDEN CHARIOT"
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
            Left            =   6240
            TabIndex        =   8
            Top             =   8400
            Width           =   3015
         End
      End
      Begin VB.PictureBox Picture3 
         Height          =   9495
         Left            =   0
         Picture         =   "railway.frx":A2533
         ScaleHeight     =   9435
         ScaleWidth      =   12675
         TabIndex        =   5
         Top             =   0
         Width           =   12735
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   8040
            Top             =   1080
         End
         Begin VB.Shape Shape1 
            BorderWidth     =   3
            Height          =   855
            Left            =   360
            Shape           =   4  'Rounded Rectangle
            Top             =   360
            Width           =   4215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "THE MAHARAJAS EXPRESS"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   600
            TabIndex        =   6
            Top             =   600
            Width           =   3735
         End
         Begin ComctlLib.ImageList ImageList2 
            Left            =   6240
            Top             =   960
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   600
            ImageHeight     =   400
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   4
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":B456B
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":16423D
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":20068F
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":292EA1
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   4500
      TabIndex        =   0
      Top             =   0
      Width           =   4560
      Begin ComctlLib.Toolbar Toolbar1 
         Height          =   1110
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   1958
         ButtonWidth     =   1799
         ButtonHeight    =   1852
         Appearance      =   1
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   3
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Reservation"
               Object.ToolTipText     =   "To make Reservations"
               Object.Tag             =   ""
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "My Plans"
               Object.ToolTipText     =   "To view Reservations"
               Object.Tag             =   ""
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Fare"
               Object.Tag             =   ""
               ImageIndex      =   4
            EndProperty
         EndProperty
         Begin ComctlLib.ImageList ImageList1 
            Left            =   600
            Top             =   120
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   48
            ImageHeight     =   48
            MaskColor       =   12632256
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   4
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":3818CB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":3D80CD
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":3FD343
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "railway.frx":412795
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Indian Railway Map"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         TabIndex        =   2
         Top             =   360
         Width           =   3855
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   18000
         Picture         =   "railway.frx":4427E7
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   14280
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   14280
         TabIndex        =   11
         Top             =   120
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "TIME  -"
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
         Left            =   13320
         TabIndex        =   10
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "DATE  -"
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
         Left            =   13320
         TabIndex        =   9
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   17400
         TabIndex        =   3
         Top             =   0
         Width           =   3165
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu mnufile 
      Caption         =   "File"
      Begin VB.Menu mnureservation 
         Caption         =   "Reservation"
      End
      Begin VB.Menu mnufare 
         Caption         =   "Fare Details"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuwindow 
      Caption         =   "Window"
      Begin VB.Menu mnuhorizontal 
         Caption         =   "Tile Horizontal"
      End
      Begin VB.Menu mnuvertical 
         Caption         =   "Tile Vertical"
      End
      Begin VB.Menu mnucascade 
         Caption         =   "Cascade"
      End
   End
   Begin VB.Menu mnuabout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "railway"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim c, interval, d As Variant
Private Sub Command1_Click()
Map.Show
End Sub

Private Sub Image1_Click()
Me.MousePointer = vbHourglass
Unload Me
login.Show
End Sub

Private Sub MDIForm_Load()
Label1.Caption = "welcome"
Picture3.Width = Screen.Width / 2
Picture4.Width = Screen.Width / 2
Picture4.Left = Screen.Width / 2
Call picStretch1
Call picStretch2
c = 2: interval = 1000: d = 4
End Sub

Sub picStretch1()
    Picture3.ScaleMode = 3
    Picture3.AutoRedraw = True
    Picture3.PaintPicture Picture3.Picture, _
        0, 0, Picture3.ScaleWidth, Picture3.ScaleHeight, _
        0, 0, _
        Picture3.Picture.Width / 26.46, _
        Picture3.Picture.Height / 26.46
    Picture3.Picture = Picture3.Image
End Sub
Sub picStretch2()
    Picture4.ScaleMode = 3
    Picture4.AutoRedraw = True
    Picture4.PaintPicture Picture4.Picture, _
        0, 0, Picture4.ScaleWidth, Picture4.ScaleHeight, _
        0, 0, _
        Picture4.Picture.Width / 26.46, _
        Picture4.Picture.Height / 26.46
    Picture4.Picture = Picture4.Image
End Sub

Private Sub mnuabout_Click()
about.Show
End Sub

Private Sub mnucascade_Click()
railway.Arrange vbTilecascade
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnufare_Click()
fare.Show
End Sub

Private Sub mnuhorizontal_Click()
railway.Arrange vbTileHorizontal
End Sub

Private Sub mnureservation_Click()
reservation.Show
End Sub

Private Sub mnuvertical_Click()
railway.Arrange vbTileVertical
End Sub

Private Sub Timer1_Timer()
interval = interval + Timer1.interval
If c = 1 And (interval Mod 4000 = 0) Then
Picture3.Picture = ImageList2.ListImages(c).Picture
Label2.Caption = "THE MAHARAJAS EXPRESS"
c = 2
Call picStretch1
ElseIf c = 2 And (interval Mod 4000 = 0) Then
Picture3.Picture = ImageList2.ListImages(c).Picture
Label2.Caption = "THE DECCAN ODYSSEY"
c = 1
Call picStretch1
End If
If interval = 2000 Then
Timer2.Enabled = True
End If

Label7.Caption = Time$
Label6.Caption = DateValue(Now)
End Sub

Private Sub Timer2_Timer()
If d = 3 Then
Picture4.Picture = ImageList2.ListImages(d).Picture
Label3.Caption = "THE ROYAL CHARIOT"
d = 4
Label3.Top = Label3.Top + 520
Shape2.Top = Shape2.Top + 520
Call picStretch2
ElseIf d = 4 Then
Picture4.Picture = ImageList2.ListImages(d).Picture
Label3.Caption = "FAIRY QUEEN"
d = 3
Label3.Top = Label3.Top - 520
Shape2.Top = Shape2.Top - 520
Call picStretch2
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.Caption
Case "Reservation"
reservation.Show
Case "My Plans"
plan.Show
Case "Fare"
fare.Show
End Select
End Sub
