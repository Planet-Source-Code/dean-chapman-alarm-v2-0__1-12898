VERSION 5.00
Begin VB.Form frmregkey 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Enter Registration Code....."
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4575
   Icon            =   "frmregkey.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Register"
      Default         =   -1  'True
      Height          =   735
      Left            =   1560
      TabIndex        =   5
      Top             =   2880
      Width           =   1455
   End
   Begin VB.TextBox txtpass 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtname 
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "REGISTRATION KEY :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "NAME :"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Carefully enter your name and registration key"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   285
      TabIndex        =   0
      Top             =   870
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   510
      Picture         =   "frmregkey.frx":030A
      Top             =   105
      Width           =   1650
   End
   Begin VB.Image Image2 
      Height          =   1290
      Left            =   2805
      Picture         =   "frmregkey.frx":0D19
      Top             =   105
      Width           =   1725
   End
End
Attribute VB_Name = "frmregkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtpass = "12345" Then
Dim TheValue As Long
TheValue = 20
Call SaveSetting("Merlin Computers", "Timer", "12345", TheValue)
Set fs = CreateObject("Scripting.FileSystemObject")
Set a = fs.CreateTextFile("C:\windows\reginfo.dll", True)
a.writeline txtname.Text
a.Close
Unload Me
MsgBox ("Thankyou For Registering This Product")
Unload frmnag
frmsplash.Visible = True
Else
MsgBox ("Wrong Please try again")
txtpass.Text = ""
txtpass.SetFocus
End If
End Sub
