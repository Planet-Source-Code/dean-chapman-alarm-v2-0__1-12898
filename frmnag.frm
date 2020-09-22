VERSION 5.00
Begin VB.Form frmnag 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Timer"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5010
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmnag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.CommandButton cmdcon 
      Caption         =   "Continue Using Unregistered Version"
      Height          =   735
      Left            =   3240
      TabIndex        =   2
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton cmdentreg 
      Caption         =   "Enter Reg Code"
      Default         =   -1  'True
      Height          =   735
      Left            =   240
      TabIndex        =   1
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Image Image3 
      Height          =   1290
      Left            =   3270
      Picture         =   "frmnag.frx":030A
      Top             =   45
      Width           =   1725
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Copyright Merlin Computers 1999-2000"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   5040
      Width           =   4455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   $"frmnag.frx":0ADE
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   360
      TabIndex        =   0
      Top             =   2640
      Width           =   4215
   End
   Begin VB.Image Image2 
      Height          =   1125
      Left            =   30
      Picture         =   "frmnag.frx":0BE0
      Top             =   1320
      Width           =   4935
   End
   Begin VB.Image Image1 
      Height          =   705
      Left            =   645
      Picture         =   "frmnag.frx":30E9
      Top             =   285
      Width           =   1650
   End
End
Attribute VB_Name = "frmnag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcon_Click()
frmmain.Visible = True
Unload frmnag
End Sub

Private Sub cmdentreg_Click()
frmregkey.Visible = True
End Sub

Private Sub Form_Load()
Dim Value As Long
Value = GetSetting("Merlin Computers", "Timer", "12345", 0)
If Value <> 0 Then
        Unload frmnag
    On Error Resume Next
Set fs = CreateObject("Scripting.FileSystemObject")
Set File = fs.GetFile("C:\windows\reginfo.dll")
If Err.Number = 53 Then 'File not found
MsgBox "File Not Found"
End If
If Err.Number = 0 Then 'File found
frmsplash.Visible = True
Open "reginfo.dll" For Input As #1
frmsplash.Label1 = Input(LOF(1), 1)
Close #1
Unload Me
'MsgBox "File Found"
End If
Else:
    frmnag.Visible = True
End If
End Sub
