VERSION 5.00
Begin VB.Form StaffInfo 
   BackColor       =   &H00404080&
   Caption         =   "General Staff Info"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7230
   DrawMode        =   1  'Blackness
   Icon            =   "staff.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00404080&
      Caption         =   " General Staff "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   6135
      Left            =   3840
      TabIndex        =   12
      Top             =   1440
      Width           =   8055
      Begin VB.CommandButton Command5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Edit Entry"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5280
         Width           =   1575
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&Clear Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5280
         Width           =   1575
      End
      Begin VB.ComboBox Staffid 
         DataField       =   "Staff_id"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3720
         TabIndex        =   1
         Top             =   960
         Width           =   3015
      End
      Begin VB.TextBox depmt 
         DataField       =   "Department"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3720
         TabIndex        =   2
         Top             =   1800
         Width           =   3015
      End
      Begin VB.TextBox post 
         DataSource      =   "Adodc1"
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
         Left            =   3720
         TabIndex        =   3
         Top             =   2640
         Width           =   3015
      End
      Begin VB.TextBox Staffname 
         DataSource      =   "Adodc1"
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
         Left            =   3720
         TabIndex        =   4
         Top             =   3480
         Width           =   3015
      End
      Begin VB.TextBox Salary 
         DataSource      =   "Adodc1"
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
         Left            =   3720
         TabIndex        =   5
         Top             =   4320
         Width           =   3015
      End
      Begin VB.Label Staffinfo 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff id:"
         DataField       =   "Staff_r"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Staff Name:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   16
         Top             =   3480
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   15
         Top             =   4320
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Post:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   14
         Top             =   2640
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "Department:"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   495
         Left            =   1200
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Erase Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Modify Record"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   1575
   End
   Begin VB.PictureBox Adodc3 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   435
      ScaleWidth      =   555
      TabIndex        =   11
      Top             =   480
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      BackStyle       =   0  'Transparent
      Caption         =   "General Staff Info"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   855
      Left            =   5160
      TabIndex        =   0
      Top             =   480
      Width           =   5775
   End
End
Attribute VB_Name = "StaffInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public myrs As New ADODB.Recordset


Private Sub Command4_Click()
Call clear_focus
End Sub


Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Staffid.Text = "" Then
MsgBox "Enter the Staff ID", vbExclamation + vbOKCancel, "Staff ID?"
Else
If Staffid.Text = "" Or depmt.Text = "" Or post.Text = "" Or Staffname.Text = "" Or Salary.Text = "" Then
MsgBox "Some data missing. Enter the full details", vbCritical + vbOKCancel, "Missing Data."
Else
myrs.Open "select * from Staff_info", conn, adOpenKeyset, adLockOptimistic, -1
myrs.AddNew
myrs!staff_id = Trim(Staffid.Text)
myrs!Department = depmt.Text
myrs!post = post.Text
myrs!staff_name = Staffname.Text
myrs!Salary = Salary.Text
myrs.Update
myrs.Close
myrs.Open "select * from staff_info where staff_id = '" & Staffid.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
Call clear_focus
While Not myrs.EOF
Staffid.AddItem (myrs!staff_id)
myrs.MoveNext
Wend
myrs.Close
End If
End If
End Sub

Private Sub Command3_Click()
If Staffid.Text = "" Then
MsgBox "There is no current record to delete. Please select a Medicine Name.", vbExclamation + vbOKOnly, "Missing..."
Staffid.SetFocus
Exit Sub
End If
myrs.Open "select * from staff_info where Staff_id = '" & Staffid.Text & "' ", conn, adOpenKeyset, adLockOptimistic, -1
myrs.Delete
MsgBox "Entry Deleted!", vbOKOnly + vbExclamation, "Success..."
myrs.Close
Call clear_focus
Unload Me
Main.Show
Main.WindowState = 2
End Sub

Private Sub Command5_Click()
myrs.Open "select * from staff_info where staff_id = '" & Staffid & "'", conn, adOpenStatic, adLockOptimistic, -1
myrs!staff_id = Trim(Staffid.Text)
myrs!staff_name = Trim(Staffname.Text)
myrs!Department = Trim(depmt.Text)
myrs!post = Trim(post.Text)
myrs!Salary = Trim(Salary.Text)
myrs.Update
MsgBox "Record Modified Successfully", vbInformation, "Success"
myrs.Close
myrs.Open "select * from staff_info where staff_id = '" & Staffid & "'", conn, adOpenStatic, adLockOptimistic, -1
Staffname.Text = myrs!staff_name
depmt.Text = myrs!Department
post.Text = myrs!post
Salary.Text = myrs!Salary
myrs.Close
Call clear_focus
End Sub

Private Sub Form_Load()
Call connect
myrs.Open "select * from Staff_info", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
Staffid.AddItem (myrs!staff_id)
myrs.MoveNext
Wend
myrs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
End Sub


Private Sub Staffid_click()
myrs.Open "select * from staff_info where staff_id = '" & Staffid.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
depmt.Text = myrs!Department
post.Text = myrs!post
Staffname.Text = myrs!staff_name
Salary.Text = myrs!Salary
myrs.Close
End Sub


Private Function clear_focus()
Staffid.Text = ""
depmt.Text = ""
post.Text = ""
Staffname.Text = ""
Salary.Text = ""
Staffid.SetFocus
End Function

