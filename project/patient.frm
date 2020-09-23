VERSION 5.00
Begin VB.Form patient 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Patient Details"
   ClientHeight    =   6780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6615
   Icon            =   "patient.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6780
   ScaleWidth      =   6615
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   " Patient Details "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   5775
      Left            =   4080
      TabIndex        =   10
      Top             =   1560
      Width           =   7815
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Edit Entry"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5040
         Width           =   1455
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Clear Page"
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
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5040
         Width           =   1455
      End
      Begin VB.ComboBox pat_id 
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
         Left            =   4200
         TabIndex        =   1
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox Pat_name 
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
         Left            =   4200
         TabIndex        =   2
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox ward_no 
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
         Left            =   4200
         TabIndex        =   4
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox details 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   4200
         TabIndex        =   3
         Top             =   2520
         Width           =   2895
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient ID:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Ward No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   12
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00404000&
         BackStyle       =   0  'Transparent
         Caption         =   "Address && Other Info:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   11
         Top             =   2640
         Width           =   3135
      End
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Delete Entry"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Save Details"
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
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7920
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "E&xit"
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
      Left            =   10080
      MaskColor       =   &H00404000&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Pat_info 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient Details"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   615
      Left            =   6240
      TabIndex        =   0
      Top             =   600
      Width           =   3495
   End
End
Attribute VB_Name = "patient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myrs As New ADODB.Recordset


Private Sub Command1_Click()
Unload Me
Main.Show
Main.SetFocus
Main.WindowState = 2
End Sub

Private Sub Command2_Click()
If pat_id.Text = "" Then
MsgBox "Enter the Patient Id", vbExclamation + vbApplicationModal, "Patient Id?"
Else
If Pat_name.Text = "" Or details.Text = "" Or ward_no.Text = "" Then
MsgBox "Some data missing. Enter the full details", vbCritical + vbOKCancel, "Missing Data."
Else
myrs.Open "select * from patient", conn, adOpenKeyset, adLockOptimistic, -1
myrs.AddNew
myrs!patient_id = pat_id.Text
myrs!patient_name = Pat_name.Text
myrs!patient_details = details.Text
myrs!ward_no = ward_no.Text
myrs.Update
myrs.Close
myrs.Open "select * from patient where patient_id = '" & pat_id.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
Call clear_focus
While Not myrs.EOF
pat_id.AddItem (myrs!patient_id)
myrs.MoveNext
Wend
myrs.Close
End If
End If

End Sub

Private Sub Command3_Click()
If pat_id.Text = "" Then
MsgBox "There is no current record to delete. Please select a Medicine Name.", vbExclamation + vbOKOnly, "Missing..."
pat_id.SetFocus
Exit Sub
End If
myrs.Open "select * from patient where patient_id =" & pat_id & " ", conn, adOpenKeyset, adLockOptimistic, -1
myrs.Delete
MsgBox "Entry Deleted!", vbOKOnly + vbExclamation, "Success..."
myrs.Close
Call clear_focus
Unload Me
Main.Show
Main.WindowState = 2
End Sub

Private Sub Command4_Click()
Call clear_focus
End Sub

Private Sub Command5_Click()
myrs.Open "select * from patient where patient_id = '" & pat_id & "'", conn, adOpenStatic, adLockOptimistic, -1
myrs!patient_id = Trim(pat_id.Text)
myrs!patient_name = Trim(Pat_name.Text)
myrs!patient_details = Trim(details.Text)
myrs!ward_no = Trim(ward_no.Text)
myrs.Update
MsgBox "Record Modified Successfully", vbInformation, "Success"
myrs.Close
myrs.Open "select * from patient where patient_id = '" & pat_id & "'", conn, adOpenStatic, adLockOptimistic, -1
Pat_name.Text = myrs!patient_name
details.Text = myrs!patient_details
ward_no.Text = myrs!ward_no
myrs.Close
Call clear_focus
End Sub

Private Sub Form_Load()
Call connect
myrs.Open "select * from patient", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
pat_id.AddItem (myrs!patient_id)
myrs.MoveNext
Wend
myrs.Close
End Sub


Private Sub pat_id_click()
myrs.Open "select * from patient where patient_id = '" & pat_id.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
Pat_name.Text = myrs!patient_name
details.Text = myrs!patient_details
ward_no.Text = myrs!ward_no
myrs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
End Sub

Private Function clear_focus()
pat_id.Text = ""
Pat_name.Text = ""
details.Text = ""
ward_no.Text = ""
pat_id.SetFocus
End Function
