VERSION 5.00
Begin VB.Form frmsplash 
   BackColor       =   &H00000000&
   Caption         =   "Timer"
   ClientHeight    =   3825
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmsplash.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3825
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3720
      Top             =   360
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Copyright Merlin Computers 1999-2000"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   1200
      Picture         =   "frmsplash.frx":030A
      Top             =   30
      Width           =   1650
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "This Program is Registered to :"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   1200
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   4455
   End
   Begin VB.Image Image2 
      Height          =   1290
      Left            =   2925
      Picture         =   "frmsplash.frx":0D19
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Timer1.Enabled = True
Dim X As Object
Dim s As String
Dim sz As String
Set X = CreateObject("Scripting.FileSystemObject")
If X.FileExists("c:\windows\reginfo.dll") Then
    Open "c:\windows\reginfo.dll" For Input As #1
        While Not EOF(1)
            Line Input #1, s
            sz = sz + vbCrLf + s
        Wend
    Close #1
    frmsplash.Label1.Caption = sz
Else
    MsgBox "No file!"
End If
End Sub

Private Sub Timer1_Timer()
Unload Me
frmmain.Visible = True
End Sub
