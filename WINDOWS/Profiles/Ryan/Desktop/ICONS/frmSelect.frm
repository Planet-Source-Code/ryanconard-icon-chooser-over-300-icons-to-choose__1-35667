VERSION 5.00
Begin VB.Form frmSelect 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Selection"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4815
   Icon            =   "frmSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close The Program"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2040
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close This"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   1335
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
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ListBox lstIcons 
         Height          =   840
         ItemData        =   "frmSelect.frx":0742
         Left            =   120
         List            =   "frmSelect.frx":076D
         TabIndex        =   2
         Top             =   360
         Width           =   4335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Show Me"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   1320
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmSelect"
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
frmMain.Show
Me.Hide
End Sub

Private Sub Command3_Click()
Unload Me
End
End Sub
