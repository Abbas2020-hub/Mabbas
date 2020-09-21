VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H80000007&
   Caption         =   "Apollo Hospital"
   ClientHeight    =   8730
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   12135
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8730
   ScaleWidth      =   12135
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "  Smallville Medical Center "
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   4455
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   11415
      Begin VB.CommandButton Department 
         Caption         =   "Department"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         TabIndex        =   0
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Staff 
         Caption         =   "Staff"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton PInfo 
         Caption         =   "Patient Info"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6120
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton billbut 
         Caption         =   "Billing"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7320
         TabIndex        =   5
         Top             =   2880
         Width           =   1455
      End
      Begin VB.CommandButton Medicines 
         Caption         =   "Medicines"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8520
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.CommandButton Stores 
         Caption         =   "Stores"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2520
         TabIndex        =   4
         Top             =   2880
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
      BackStyle       =   0  'Transparent
      Caption         =   "caring for 62 years..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   9240
      TabIndex        =   9
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000007&
      Caption         =   "Smallville, Kansas"
      BeginProperty Font 
         Name            =   "Book Antiqua"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   6720
      TabIndex        =   7
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000007&
      Caption         =   "Smallville Medical Center Hospital"
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   4320
      TabIndex        =   6
      Top             =   600
      Width           =   6615
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Billing_Click()
End Sub

Private Sub billbut_Click()
Call Billing.Show
Billing.SetFocus

End Sub

Private Sub Department_Click()
Dept.Show
Dept.SetFocus

End Sub

Private Sub Image1_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Main.WindowState = 2
End Sub

Private Sub Medicines_Click()
medicine.Show
medicine.SetFocus

End Sub

Private Sub PInfo_Click()
patient.Show
patient.SetFocus
End Sub

Private Sub Stores_Click()
store.Show
store.SetFocus
End Sub

Private Sub Label3_Click()
Unload Me
End Sub

Private Sub Staff_Click()
Staffinfo.Show
Staffinfo.SetFocus
End Sub
