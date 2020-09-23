VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2520
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2835
   LinkTopic       =   "Form1"
   ScaleHeight     =   168
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   189
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   2880
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   2520
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   2295
      Left            =   120
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Picture2.BackColor = 0
Timer1.Interval = 1000
End Sub

Private Sub Timer1_Timer()
Dim mymean As Double, myvariance As Double

Timer1.Enabled = False

GetScreen Picture1
CalculateDescriptives Picture1, mymean, myvariance
DrawZ Picture1, Picture2, mymean, myvariance, 2.5
Stats Picture2, Picture1
Timer1.Enabled = True
End Sub
