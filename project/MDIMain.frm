VERSION 5.00
Begin VB.MDIForm MDIMain 
   BackColor       =   &H00808000&
   Caption         =   "Smallville Medical Center"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11010
   Icon            =   "MDIMain.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   14520
      Top             =   9480
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1035
      ScaleWidth      =   10950
      TabIndex        =   0
      Top             =   3600
      Width           =   11010
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         Caption         =   "&Main Page"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         Caption         =   "Disclaimer"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         Caption         =   "Help && &About"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   480
         Width           =   1695
      End
      Begin VB.Image Image1 
         Height          =   825
         Left            =   13440
         Picture         =   "MDIMain.frx":628A
         Stretch         =   -1  'True
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   14280
         TabIndex        =   6
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblTime 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   11160
         TabIndex        =   5
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label lbldate 
         Alignment       =   1  'Right Justify
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   6840
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label lblday 
         Alignment       =   1  'Right Justify
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   8400
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.Label lblMonth 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   8880
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblyear 
         Alignment       =   1  'Right Justify
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   10320
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "MDIMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Call About.Show
About.SetFocus
End Sub

Private Sub Command2_Click()
Call Disclaimer.Show
Disclaimer.SetFocus
End Sub

Private Sub Command3_Click()
Main.Show
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()
Call Main.Show
End Sub

Private Sub Timer1_Timer()
Dim Today As Variant
Today = Now
lbldate.Caption = Format(Today, "dddd")
lblMonth.Caption = Format(Today, "mmmm")
lblyear.Caption = Format(Today, "yyyy")
lblday.Caption = Format(Today, "d")
lblTime.Caption = Format(Today, "h:mm:ss ampm")
End Sub
