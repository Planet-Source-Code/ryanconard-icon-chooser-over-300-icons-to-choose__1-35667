VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ICON - Version 1.2.7 [ALPHA]"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1215
      Left            =   3120
      TabIndex        =   10
      Top             =   3720
      Width           =   1575
      Begin VB.Line Line4 
         X1              =   240
         X2              =   240
         Y1              =   1080
         Y2              =   480
      End
      Begin VB.Line Line3 
         X1              =   1080
         X2              =   1080
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Line Line2 
         X1              =   1080
         X2              =   1440
         Y1              =   1080
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   480
         Y1              =   240
         Y2              =   720
      End
      Begin VB.Label Label5 
         Caption         =   "ICON"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   11
         Top             =   840
         Width           =   615
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   480
         Picture         =   "frmMain.frx":0742
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Updates"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   2895
      Begin VB.Label Label4 
         Caption         =   "23 Days Until Finish Date"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label3 
         Caption         =   "17 % Done..."
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Alpha Version"
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Categories"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
      Begin VB.ListBox lstIcons 
         Height          =   840
         ItemData        =   "frmMain.frx":0E84
         Left            =   120
         List            =   "frmMain.frx":0EAF
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
      Begin VB.Line Line8 
         X1              =   1080
         X2              =   2880
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Line Line7 
         X1              =   720
         X2              =   2880
         Y1              =   1560
         Y2              =   1560
      End
      Begin VB.Line Line6 
         X1              =   480
         X2              =   2880
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   2880
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         Caption         =   $"frmMain.frx":0F31
         BeginProperty Font 
            Name            =   "Myriad Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Select Case lstIcons
Case "Arrows"
frmArrows.Show
Me.Hide
Case "Communications"
frmComm.Show
Me.Hide
Case "Computers"
frmComp.Show
Me.Hide
Case "Drag & Drop"
frmDragDrop.Show
Me.Hide
Case "Elements"
frmElements.Show
Me.Hide
Case "Flags"
frmFlags.Show
Me.Hide
Case "Industry"
frmIndustry.Show
Me.Hide
Case "Mail"
frmMail.Show
Me.Hide
Case "Misc."
frmMisc.Show
Me.Hide
Case "Office"
frmOffice.Show
Me.Hide
Case "Traffic"
frmTraffic.Show
Me.Hide
Case "Windows 98"
frmWin98.Show
Me.Hide
Case "Writing"
frmWriting.Show
Me.Hide
End Select
End Sub

Private Sub Command2_Click()
Unload Me
End
End Sub
