VERSION 5.00
Begin VB.Form fare 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fare Details"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16980
   Icon            =   "fare.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "fare.frx":FD31
   ScaleHeight     =   6195
   ScaleWidth      =   16980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Fairy Queen"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8640
      TabIndex        =   12
      Top             =   3000
      Width           =   8055
      Begin VB.TextBox Text4 
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   13
         Text            =   "Rs. 6,804"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image Image4 
         Height          =   1695
         Left            =   240
         Picture         =   "fare.frx":670FA
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "A short, but luxurious journey to Alwar and Sariska"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2340
         TabIndex        =   15
         Top             =   840
         Width           =   5145
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label7 
         Caption         =   "Fare -"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   14
         Top             =   2160
         Width           =   855
      End
      Begin VB.Line Line4 
         X1              =   2160
         X2              =   7800
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "The Golden Chariot"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   8640
      TabIndex        =   8
      Top             =   120
      Width           =   8055
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   9
         Text            =   "Rs. 1,82,000"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Line Line3 
         X1              =   2160
         X2              =   7800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label6 
         Caption         =   "Fare -"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   11
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Discover South India in a royal way"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2880
         TabIndex        =   10
         Top             =   600
         Width           =   3945
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image3 
         Height          =   1575
         Left            =   240
         Picture         =   "fare.frx":83564
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "The Deccan Odyssey"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   240
      TabIndex        =   4
      Top             =   3000
      Width           =   8055
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   5
         Text            =   "Rs. 3,18,000"
         Top             =   2160
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   1815
         Left            =   240
         Picture         =   "fare.frx":90623
         Stretch         =   -1  'True
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   " Experience the opulence while exploring Maharashtra and Gujarat"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   2115
         TabIndex        =   7
         Top             =   600
         Width           =   5220
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Fare -"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   2160
         Width           =   855
      End
      Begin VB.Line Line2 
         X1              =   2160
         X2              =   7800
         Y1              =   2040
         Y2              =   2040
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "The Maharajas Express"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
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
         Left            =   3960
         TabIndex        =   3
         Text            =   "Rs. 5,75,000"
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   2160
         X2              =   7800
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label Label2 
         Caption         =   "Fare -"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   15
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   2
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Take a royal journey to India’s top tourist destinations"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   4815
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   240
         Picture         =   "fare.frx":9D293
         Stretch         =   -1  'True
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Image Image5 
      Height          =   9645
      Left            =   0
      Picture         =   "fare.frx":14CF55
      Stretch         =   -1  'True
      Top             =   0
      Width           =   15360
   End
End
Attribute VB_Name = "fare"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Image5.Width = fare.Width
Image5.Height = fare.Height
End Sub
