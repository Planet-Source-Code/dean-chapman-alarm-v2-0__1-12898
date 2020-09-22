VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "about"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1740
      Top             =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "*********************"
      Height          =   3090
      Left            =   30
      TabIndex        =   0
      Top             =   2985
      Width           =   4620
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
Label1.Move
Label1.Left
Label1.Top -10
End Sub
