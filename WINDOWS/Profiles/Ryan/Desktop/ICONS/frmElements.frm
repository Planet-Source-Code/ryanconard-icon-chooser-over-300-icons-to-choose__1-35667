VERSION 5.00
Begin VB.Form frmElements 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Category: Elements"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmElements.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Crash your computer.. Click here!"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   2760
      Width           =   4335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Go back to the main form.. Click here!"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit this program.. Click here!"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "See a working example.. Click here!"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please vote on PSC if you want more...."
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   3855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "NOT DONE YET!!!"
      BeginProperty Font 
         Name            =   "Myriad Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "frmElements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Me.Hide
frmArrows.Show
End Sub

Private Sub Command2_Click()
MsgBox "Please vote on PSC if you want me to continue this program.. 5 excellent votes will get another version!", vbInformation, "Please Vote"
Unload Me
End
End Sub

Private Sub Command3_Click()
Me.Hide
frmMain.Show
End Sub

Private Sub Command4_Click()
Dim YesNo As Integer
YesNo = MsgBox("This will really crash your computer [blue screen of death].. Are you sure you wanna do this?", vbYesNo, "Really?")
If YesNo = vbYes Then
Call Shell("RUNDLL32.DLL, disableoemlayering")
Exit Sub
End If
If YesNo = vbNo Then
Exit Sub
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "Please vote on PSC if you want me to continue this program.. 5 excellent votes will get another version!", vbInformation, "Please Vote"
Unload Me
End
End Sub
