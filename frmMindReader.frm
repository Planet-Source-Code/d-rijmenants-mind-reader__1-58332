VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMindReader 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mind Reader"
   ClientHeight    =   3345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6480
      Top             =   2760
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6000
      Top             =   2760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Try again"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2760
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   5280
      Top             =   2760
   End
   Begin VB.Image img1 
      Height          =   1545
      Left            =   120
      Picture         =   "frmMindReader.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   7110
   End
   Begin VB.Label lblComment 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This program will read your mind!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7215
   End
   Begin VB.Image img2 
      Height          =   1545
      Left            =   840
      Picture         =   "frmMindReader.frx":23D32
      Stretch         =   -1  'True
      Top             =   240
      Visible         =   0   'False
      Width           =   5535
   End
End
Attribute VB_Name = "frmMindReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private TC As Integer

Private Sub Command1_Click()
TC = 0
Me.Timer1.Enabled = True
Me.Command1.Visible = False
Me.lblComment.Caption = "This program will read your mind!"
End Sub

Private Sub Form_Activate()
TC = 0
End Sub

Private Sub Timer1_Timer()
TC = TC + 1
Select Case TC
Case 1
Me.lblComment.Caption = "You cannot believe it?"
Case 2
Me.lblComment.Caption = "See and be surprised!"
Case 3
Me.lblComment.Caption = "You can see some cards here..."
Me.img1.Visible = True
Case 4
Me.lblComment.Caption = "Do not touch anything..."
Case 5
Me.lblComment.Caption = "Choose one card in your mind..."
Case 6
Me.lblComment.Caption = "Focus on the card..."
Case 7
Me.lblComment.Caption = "Keep the card in your mind"
Case 8
Me.img1.Visible = False
Me.Timer2.Enabled = True
Me.lblComment.Caption = "I am reading your mind now..."
Case 9
Me.lblComment.Caption = "Do not touch anything..."
Case 10
Me.Timer2.Enabled = False
Me.BackColor = &H0&       '&HC0C0C0
Me.lblComment.Caption = "Calculating the image in your mind..."
Me.Timer1.Enabled = False
Me.ProgressBar1.Visible = True
Me.Timer3.Enabled = True
Case 11
Me.Timer3.Enabled = False
Me.lblComment.Caption = "Taking you card out..."
Case 12
Me.lblComment.Caption = "Look ! Your card is no longer in the row..."
Me.img2.Visible = True
Case 13
Me.lblComment.Caption = "Spooky no?"
Case 14
Me.img2.Visible = False
Me.lblComment.Caption = "Hit the button and try again with another card!"
Me.Command1.Visible = True
Me.Timer1.Enabled = False
End Select
End Sub

Private Sub Timer2_Timer()
c1 = Int(256 * Rnd)
c2 = Int(256 * Rnd)
c3 = Int(256 * Rnd)
Me.BackColor = RGB(c1, c2, c3)
End Sub

Private Sub Timer3_Timer()
Static Pval As Integer
Pval = Pval + 1
If Pval = 70 Then Me.lblComment.Caption = "Calculating the colors in your mind"
If Pval = 101 Then
    Pval = 0
    Me.Timer3.Enabled = False
    Me.Timer1.Enabled = True
    Me.ProgressBar1.Visible = False
    Me.ProgressBar1.Value = 0
    Me.lblComment.Caption = "The card is recognized..."
    Else
    Me.ProgressBar1.Value = Pval
    End If
End Sub
