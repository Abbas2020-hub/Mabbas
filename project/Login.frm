VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Apollo Login"
   ClientHeight    =   2400
   ClientLeft      =   5190
   ClientTop       =   4080
   ClientWidth     =   7635
   FillColor       =   &H00FFFFFF&
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Password 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox UserText 
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton Cancel 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton OK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Ok"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   2415
      Left            =   0
      Picture         =   "Login.frx":628A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1035
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   1575
      Left            =   6240
      Picture         =   "Login.frx":F29C
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1305
   End
   Begin VB.Label Pass 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   5
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label User 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Username:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   480
      Width           =   1335
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Login
End Sub

Private Sub OK_Click()
If UserText.Text = "apollo" And Password.Text = "apollo" Then
Main.Show
Unload Login
Else: MsgBox "Invalid Username or Password....", vbRetryCancel + vbCritical, "Error!!"
UserText = ""
Password = ""
UserText.SetFocus
End If
End Sub

