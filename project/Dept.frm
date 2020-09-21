VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Dept 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Department Info"
   ClientHeight    =   6240
   ClientLeft      =   5640
   ClientTop       =   4815
   ClientWidth     =   7155
   Icon            =   "Dept.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "Close && &Return"
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
      Left            =   9600
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Delete Present"
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8040
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Insert Values"
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
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8040
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   " Department Info "
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
      Width           =   7695
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0C0FF&
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
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   4920
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
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
         Left            =   5280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         Width           =   1575
      End
      Begin VB.ComboBox Dpno 
         DataField       =   "Department_no"
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
         Left            =   3600
         TabIndex        =   1
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox Deptname 
         DataField       =   "Department_name"
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
         Left            =   3600
         TabIndex        =   2
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox Doctorname 
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
         Left            =   3600
         TabIndex        =   3
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox Contact 
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
         Left            =   3600
         TabIndex        =   4
         Top             =   3960
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Department No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   14
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Department Name:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Name of Doctor:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   12
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Contact No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   3960
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   -360
      Top             =   -240
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=MATRIX"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=MATRIX"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Department Information"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   5775
   End
End
Attribute VB_Name = "Dept"
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
Call clear_focus
End Sub

Private Sub Command3_Click()
If Dpno.Text = "" Then
MsgBox "Enter the Department Number", vbExclamation + vbOKCancel, "Missing..."
Else
If Deptname.Text = "" Or Doctorname.Text = "" Or Contact.Text = "" Then
MsgBox "Some data missing. Enter the full details", vbCritical + vbApplicationModal, "Missing Data!"
Else
myrs.Open "select * from Department_information", conn, adOpenKeyset, adLockOptimistic, -1
myrs.AddNew
myrs!Department_no = Dpno.Text
myrs!Department_name = Deptname.Text
myrs!Name_of_doctor = Doctorname.Text
myrs!contact_no = Contact.Text
myrs.Update
myrs.Close
myrs.Open "select * from department_information where Department_no = '" & Dpno.Text & "' ", conn, adOpenKeyset, adLockOptimistic, -1
Call clear_focus
While Not myrs.EOF
Dpno.AddItem (myrs!Department_no)
myrs.MoveNext
Wend
myrs.Close
End If
End If
End Sub

Private Sub Command4_Click()
If Dpno.Text = "" Then
MsgBox "There is no current record to delete. Please select a Department No.", vbExclamation + vbOKOnly, "Missing..."
Dpno.SetFocus
Exit Sub
End If
myrs.Open "select * from department_information where Department_no = '" & Dpno & "' ", conn, adOpenKeyset, adLockOptimistic, -1
myrs.Delete
MsgBox "Entry Deleted!", vbOKOnly + vbExclamation, "Success..."
myrs.Close
Call clear_focus
Unload Me
Main.Show
Main.WindowState = 2
End Sub


Private Sub Command5_Click()
Dim n As String
n = Trim(Dpno.Text)
myrs.Open "select * from department_information where Department_no = '" & n & "' ", conn, adOpenStatic, adLockOptimistic, -1
myrs!Department_no = Trim(Dpno.Text)
myrs!Department_name = Trim(Deptname.Text)
myrs!Name_of_doctor = Trim(Doctorname.Text)
myrs!contact_no = Trim(Contact.Text)
myrs.Update
MsgBox "Record Modified Successfully", vbInformation, "Success"
myrs.Close
myrs.Open "select * from department_information where Department_no = '" & Dpno & "'", conn, adOpenStatic, adLockOptimistic, -1
Deptname.Text = myrs!Department_name
Doctorname.Text = myrs!Name_of_doctor
Contact.Text = myrs!contact_no
myrs.Close
Call clear_focus
End Sub

Private Sub Form_Load()
Call connect
myrs.Open "select * from department_information", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
Dpno.AddItem (myrs!Department_no)
myrs.MoveNext
Wend
myrs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
End Sub

Private Sub Dpno_click()
myrs.Open "select * from department_information where Department_no = '" & Dpno.Text & "' ", conn, adOpenKeyset, adLockOptimistic, -1
Deptname.Text = myrs!Department_name
Doctorname.Text = myrs!Name_of_doctor
Contact.Text = myrs!contact_no
myrs.Close
End Sub

Private Function clear_focus()
Dpno.Text = ""
Deptname.Text = ""
Doctorname.Text = ""
Contact.Text = ""
Dpno.SetFocus
End Function
