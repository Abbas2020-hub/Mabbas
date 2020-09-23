VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Splash 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3210
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7410
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   3090
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   240
         Top             =   2160
      End
      Begin MSComctlLib.ProgressBar p 
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   2520
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         Max             =   5000
         Scrolling       =   1
      End
      Begin VB.Timer Timer1 
         Interval        =   4500
         Left            =   0
         Top             =   3600
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "WinSys Software Solutions"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   1920
         TabIndex        =   5
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Loading...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         Caption         =   "Smallville Medical Center"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Hospital Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   6675
      End
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyPress(KeyAscii As Integer)
 Call Login.Show
Call Login.SetFocus
Unload Me
End Sub

Private Sub Frame1_Click()
Call Login.Show
Call Login.SetFocus
Unload Me
End Sub


Private Sub Timer1_Timer()
Call Login.Show
Call Login.SetFocus
Unload Me
End Sub

Private Sub Timer2_Timer()
Static s
s = s + 10
If s = 500 Then
For i = 1 To 5000
p.Value = i
Next i
End If
End Sub

