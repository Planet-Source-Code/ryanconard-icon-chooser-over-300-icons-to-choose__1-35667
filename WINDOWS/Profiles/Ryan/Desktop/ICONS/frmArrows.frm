VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmArrows 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category: Arrows"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7110
   Icon            =   "frmArrows.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8190
   ScaleWidth      =   7110
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cd1 
      Left            =   2880
      Top             =   5520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Advanced Options"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   59
      Top             =   6240
      Width           =   6855
      Begin VB.CommandButton Command3 
         Caption         =   "Choose another category of ICOCs to choose from.. Click here!"
         Height          =   375
         Left            =   120
         TabIndex        =   62
         Top             =   1320
         Width           =   6615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save this ICON somewhere else on your computer.. Click here!"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   840
         Width           =   6615
      End
      Begin VB.CommandButton Command1 
         Caption         =   "See how this ICON would look as your form ICON.. Click here!"
         Height          =   375
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Width           =   6615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Icons"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.OptionButton Option23 
         Caption         =   "Option23"
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option59 
         Caption         =   "Option59"
         Height          =   195
         Left            =   2040
         TabIndex        =   58
         Top             =   5640
         Width           =   255
      End
      Begin VB.OptionButton Option58 
         Caption         =   "Option58"
         Height          =   195
         Left            =   1440
         TabIndex        =   57
         Top             =   5640
         Width           =   255
      End
      Begin VB.OptionButton Option57 
         Caption         =   "Option57"
         Height          =   255
         Left            =   840
         TabIndex        =   56
         Top             =   5640
         Width           =   255
      End
      Begin VB.OptionButton Option56 
         Caption         =   "Option56"
         Height          =   255
         Left            =   240
         TabIndex        =   55
         Top             =   5640
         Width           =   255
      End
      Begin VB.OptionButton Option55 
         Caption         =   "Option55"
         Height          =   255
         Left            =   6360
         TabIndex        =   54
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option54 
         Caption         =   "Option54"
         Height          =   255
         Left            =   5760
         TabIndex        =   53
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option53 
         Caption         =   "Option53"
         Height          =   255
         Left            =   5160
         TabIndex        =   52
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option52 
         Caption         =   "Option52"
         Height          =   255
         Left            =   4440
         TabIndex        =   51
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option51 
         Caption         =   "Option51"
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option50 
         Caption         =   "Option50"
         Height          =   255
         Left            =   3240
         TabIndex        =   49
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option49 
         Caption         =   "Option49"
         Height          =   255
         Left            =   2640
         TabIndex        =   48
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option48 
         Caption         =   "Option48"
         Height          =   195
         Left            =   2040
         TabIndex        =   47
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option47 
         Caption         =   "Option47"
         Height          =   255
         Left            =   1440
         TabIndex        =   46
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option46 
         Caption         =   "Option46"
         Height          =   255
         Left            =   840
         TabIndex        =   45
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option44 
         Caption         =   "Option45"
         Height          =   255
         Left            =   6360
         TabIndex        =   44
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option43 
         Caption         =   "Option44"
         Height          =   255
         Left            =   5760
         TabIndex        =   43
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option42 
         Caption         =   "Option43"
         Height          =   255
         Left            =   5160
         TabIndex        =   42
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option41 
         Caption         =   "Option42"
         Height          =   255
         Left            =   4440
         TabIndex        =   41
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option40 
         Caption         =   "Option41"
         Height          =   255
         Left            =   3840
         TabIndex        =   40
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option39 
         Caption         =   "Option40"
         Height          =   255
         Left            =   3240
         TabIndex        =   39
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option38 
         Caption         =   "Option39"
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option37 
         Caption         =   "Option38"
         Height          =   255
         Left            =   2040
         TabIndex        =   37
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option36 
         Caption         =   "Option37"
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option35 
         Caption         =   "Option36"
         Height          =   255
         Left            =   840
         TabIndex        =   35
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option33 
         Caption         =   "Option35"
         Height          =   255
         Left            =   6360
         TabIndex        =   34
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option32 
         Caption         =   "Option34"
         Height          =   255
         Left            =   5760
         TabIndex        =   33
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option31 
         Caption         =   "Option33"
         Height          =   255
         Left            =   5160
         TabIndex        =   32
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option30 
         Caption         =   "Option32"
         Height          =   255
         Left            =   4440
         TabIndex        =   31
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option29 
         Caption         =   "Option31"
         Height          =   255
         Left            =   3840
         TabIndex        =   30
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option28 
         Caption         =   "Option30"
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option27 
         Caption         =   "Option29"
         Height          =   255
         Left            =   2640
         TabIndex        =   28
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option26 
         Caption         =   "Option28"
         Height          =   255
         Left            =   2040
         TabIndex        =   27
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option25 
         Caption         =   "Option27"
         Height          =   255
         Left            =   1440
         TabIndex        =   26
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option24 
         Caption         =   "Option26"
         Height          =   255
         Left            =   840
         TabIndex        =   25
         Top             =   2880
         Width           =   255
      End
      Begin VB.OptionButton Option22 
         Caption         =   "Option25"
         Height          =   255
         Left            =   6240
         TabIndex        =   24
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option21 
         Caption         =   "Option24"
         Height          =   255
         Left            =   5640
         TabIndex        =   23
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option20 
         Caption         =   "Option23"
         Height          =   255
         Left            =   5160
         TabIndex        =   22
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Option22"
         Height          =   255
         Left            =   4560
         TabIndex        =   21
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Option21"
         Height          =   255
         Left            =   3840
         TabIndex        =   20
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Option20"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option16 
         Caption         =   "Option19"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option15 
         Caption         =   "Option18"
         Height          =   255
         Left            =   2040
         TabIndex        =   17
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Option17"
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Option16"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   15
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Option15"
         Height          =   255
         Left            =   6240
         TabIndex        =   14
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option10 
         Caption         =   "Option14"
         Height          =   255
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Option13"
         Height          =   255
         Left            =   5040
         TabIndex        =   12
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Option12"
         Height          =   255
         Left            =   4440
         TabIndex        =   11
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Option11"
         Height          =   255
         Left            =   3840
         TabIndex        =   10
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Option10"
         Height          =   255
         Left            =   3240
         TabIndex        =   9
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Option9"
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Option8"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   7
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option7"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   6
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option6"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   5
         Top             =   960
         Width           =   255
      End
      Begin VB.OptionButton Option45 
         Caption         =   "Option5"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   4800
         Width           =   255
      End
      Begin VB.OptionButton Option34 
         Caption         =   "Option4"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   3840
         Width           =   255
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Option2"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   1920
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   960
         Width           =   255
      End
      Begin VB.Image Image59 
         Height          =   480
         Left            =   2040
         Picture         =   "frmArrows.frx":0742
         Top             =   5160
         Width           =   480
      End
      Begin VB.Image Image58 
         Height          =   480
         Left            =   1440
         Picture         =   "frmArrows.frx":0B84
         Top             =   5160
         Width           =   480
      End
      Begin VB.Image Image57 
         Height          =   480
         Left            =   840
         Picture         =   "frmArrows.frx":0FC6
         Top             =   5160
         Width           =   480
      End
      Begin VB.Image Image56 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":1408
         Top             =   5160
         Width           =   480
      End
      Begin VB.Image Image55 
         Height          =   480
         Left            =   6240
         Picture         =   "frmArrows.frx":184A
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image54 
         Height          =   480
         Left            =   5640
         Picture         =   "frmArrows.frx":1C8C
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image53 
         Height          =   480
         Left            =   5040
         Picture         =   "frmArrows.frx":20CE
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image52 
         Height          =   480
         Left            =   4320
         Picture         =   "frmArrows.frx":2510
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image51 
         Height          =   480
         Left            =   3720
         Picture         =   "frmArrows.frx":2952
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image50 
         Height          =   480
         Left            =   3120
         Picture         =   "frmArrows.frx":2D94
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image49 
         Height          =   480
         Left            =   2520
         Picture         =   "frmArrows.frx":31D6
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image48 
         Height          =   480
         Left            =   1920
         Picture         =   "frmArrows.frx":3618
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image47 
         Height          =   480
         Left            =   1320
         Picture         =   "frmArrows.frx":3A5A
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image46 
         Height          =   480
         Left            =   720
         Picture         =   "frmArrows.frx":3E9C
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image45 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":42DE
         Top             =   4200
         Width           =   480
      End
      Begin VB.Image Image44 
         Height          =   480
         Left            =   6240
         Picture         =   "frmArrows.frx":4720
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image43 
         Height          =   480
         Left            =   5640
         Picture         =   "frmArrows.frx":4B62
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image42 
         Height          =   480
         Left            =   5040
         Picture         =   "frmArrows.frx":4FA4
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image41 
         Height          =   480
         Left            =   4320
         Picture         =   "frmArrows.frx":53E6
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image40 
         Height          =   480
         Left            =   3720
         Picture         =   "frmArrows.frx":5828
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image39 
         Height          =   480
         Left            =   3120
         Picture         =   "frmArrows.frx":5C6A
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image38 
         Height          =   480
         Left            =   2520
         Picture         =   "frmArrows.frx":60AC
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image37 
         Height          =   480
         Left            =   1920
         Picture         =   "frmArrows.frx":64EE
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image36 
         Height          =   480
         Left            =   1320
         Picture         =   "frmArrows.frx":6930
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image35 
         Height          =   480
         Left            =   720
         Picture         =   "frmArrows.frx":6D72
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image34 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":71B4
         Top             =   3240
         Width           =   480
      End
      Begin VB.Image Image33 
         Height          =   480
         Left            =   6240
         Picture         =   "frmArrows.frx":75F6
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image32 
         Height          =   480
         Left            =   5640
         Picture         =   "frmArrows.frx":7A38
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image31 
         Height          =   480
         Left            =   5040
         Picture         =   "frmArrows.frx":7E7A
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image30 
         Height          =   480
         Left            =   4320
         Picture         =   "frmArrows.frx":82BC
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image29 
         Height          =   480
         Left            =   3720
         Picture         =   "frmArrows.frx":86FE
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image28 
         Height          =   480
         Left            =   3120
         Picture         =   "frmArrows.frx":8B40
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image27 
         Height          =   480
         Left            =   2520
         Picture         =   "frmArrows.frx":8F82
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image26 
         Height          =   480
         Left            =   1920
         Picture         =   "frmArrows.frx":93C4
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image25 
         Height          =   480
         Left            =   1320
         Picture         =   "frmArrows.frx":9806
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image24 
         Height          =   480
         Left            =   720
         Picture         =   "frmArrows.frx":9C48
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image23 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":A08A
         Top             =   2280
         Width           =   480
      End
      Begin VB.Image Image22 
         Height          =   480
         Left            =   6120
         Picture         =   "frmArrows.frx":A4CC
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image21 
         Height          =   480
         Left            =   5520
         Picture         =   "frmArrows.frx":A90E
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image20 
         Height          =   480
         Left            =   5040
         Picture         =   "frmArrows.frx":AD50
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image19 
         Height          =   480
         Left            =   4440
         Picture         =   "frmArrows.frx":B192
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image18 
         Height          =   480
         Left            =   3720
         Picture         =   "frmArrows.frx":B5D4
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image17 
         Height          =   480
         Left            =   3120
         Picture         =   "frmArrows.frx":BA16
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image16 
         Height          =   480
         Left            =   2520
         Picture         =   "frmArrows.frx":BE58
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image15 
         Height          =   480
         Left            =   1920
         Picture         =   "frmArrows.frx":C29A
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image14 
         Height          =   480
         Left            =   1320
         Picture         =   "frmArrows.frx":C6DC
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image13 
         Height          =   480
         Left            =   720
         Picture         =   "frmArrows.frx":CB1E
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image12 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":CF60
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image11 
         Height          =   480
         Left            =   6120
         Picture         =   "frmArrows.frx":D3A2
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image10 
         Height          =   480
         Left            =   5520
         Picture         =   "frmArrows.frx":D7E4
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image9 
         Height          =   480
         Left            =   4920
         Picture         =   "frmArrows.frx":DC26
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image8 
         Height          =   480
         Left            =   4320
         Picture         =   "frmArrows.frx":E068
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image7 
         Height          =   480
         Left            =   3720
         Picture         =   "frmArrows.frx":E4AA
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image6 
         Height          =   480
         Left            =   3120
         Picture         =   "frmArrows.frx":E8EC
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image5 
         Height          =   480
         Left            =   2520
         Picture         =   "frmArrows.frx":ED2E
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image4 
         Height          =   480
         Left            =   1920
         Picture         =   "frmArrows.frx":F170
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   1320
         Picture         =   "frmArrows.frx":F5B2
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   720
         Picture         =   "frmArrows.frx":F9F4
         Top             =   360
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmArrows.frx":FE36
         Top             =   360
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Sorry about all the nested if..
'It's easy to follow though, all the options go with the images..
If Option1.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image1
Exit Sub
End If
If Option2(1).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image2
Exit Sub
End If
If Option3(1).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image3
Exit Sub
End If
If Option4(1).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image4
Exit Sub
End If
If Option5(1).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image5
Exit Sub
End If
If Option6.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image6
Exit Sub
End If
If Option7.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image7
Exit Sub
End If
If Option8.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image8
Exit Sub
End If
If Option9.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image9
Exit Sub
End If
If Option10.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image10
Exit Sub
End If
If Option11.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image11
Exit Sub
End If
If Option12(0).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image12
Exit Sub
End If
If Option13(1).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image13
Exit Sub
End If
If Option14.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image14
Exit Sub
End If
If Option15.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image15
Exit Sub
End If
If Option16.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image16
Exit Sub
End If
If Option17.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image17
Exit Sub
End If
If Option18.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image18
Exit Sub
End If
If Option19.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image19
Exit Sub
End If
If Option20.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image20
Exit Sub
End If
If Option21.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image21
Exit Sub
End If
If Option22.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image22
Exit Sub
End If
If Option23.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image23
Exit Sub
End If
If Option24.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image24
Exit Sub
End If
If Option25.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image25
Exit Sub
End If
If Option26.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image26
Exit Sub
End If
If Option27.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image27
Exit Sub
End If
If Option28.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image28
Exit Sub
End If
If Option29.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image29
Exit Sub
End If
If Option30.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image30
Exit Sub
End If
If Option31.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image31
Exit Sub
End If
If Option32.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image32
Exit Sub
End If
If Option33.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image33
Exit Sub
End If
If Option34(0).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image34
Exit Sub
End If
If Option35.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image35
Exit Sub
End If
If Option36.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image36
Exit Sub
End If
If Option37.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image37
Exit Sub
End If
If Option38.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image38
Exit Sub
End If
If Option39.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image39
Exit Sub
End If
If Option40.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image40
Exit Sub
End If
If Option41.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image41
Exit Sub
End If
If Option42.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image42
Exit Sub
End If
If Option43.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image43
Exit Sub
End If
If Option44.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image44
Exit Sub
End If
If Option45(0).Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image45
Exit Sub
End If
If Option46.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image46
Exit Sub
End If
If Option47.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image47
Exit Sub
End If
If Option48.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image48
Exit Sub
End If
If Option49.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image49
Exit Sub
End If
If Option50.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image50
Exit Sub
End If
If Option51.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image51
Exit Sub
End If
If Option52.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image52
Exit Sub
End If
If Option53.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image53
Exit Sub
End If
If Option54.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image54
Exit Sub
End If
If Option55.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image55
Exit Sub
End If
If Option56.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image56
Exit Sub
End If
If Option57.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image57
Exit Sub
End If
If Option58.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image58
Exit Sub
End If
If Option59.Value = True Then
frmTest.Show
frmTest.ICON = frmArrows.Image59
Exit Sub
End If
End Sub

Private Sub Command2_Click()
MsgBox "Yeah Right! I havn't rigured this one out yet.. I'll get to it soon enough.", vbCritical, "You wish you could save this Icon anywhere..."
End Sub

Private Sub Command3_Click()
Me.Hide
frmSelect.Show
End Sub
