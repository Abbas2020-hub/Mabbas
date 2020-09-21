VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form StaffInfo 
   BackColor       =   &H80000006&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Staff"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7230
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   720
      Top             =   240
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
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
      Connect         =   ""
      OLEDBString     =   ""
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
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   720
      ScaleHeight     =   555
      ScaleWidth      =   2355
      TabIndex        =   12
      Top             =   6480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   6480
      Width           =   1095
   End
   Begin VB.TextBox Salary 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2760
      TabIndex        =   10
      Top             =   5400
      Width           =   3015
   End
   Begin VB.TextBox Staffname 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox post 
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2760
      TabIndex        =   8
      Top             =   3720
      Width           =   3015
   End
   Begin VB.TextBox depmt 
      DataField       =   "Department"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   2760
      TabIndex        =   7
      Top             =   2880
      Width           =   3015
   End
   Begin VB.ComboBox sid 
      DataField       =   "Staff_id"
      DataSource      =   "Adodc1"
      Height          =   315
      Left            =   2760
      TabIndex        =   2
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   5
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000012&
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000012&
      Caption         =   "Staff_Name:"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Staffinfo 
      BackColor       =   &H80000012&
      Caption         =   "Staffid:"
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
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Staff Info"
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   3015
   End
End
Attribute VB_Name = "StaffInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public conn As New ADODB.Connection
Public myrs As New ADODB.Recordset
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
conn.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=C5-15"
myrs.Open "select * from Staff_info", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
sid.AddItem (myrs!Staff_id)
myrs.MoveNext
Wend
myrs.Close
End Sub

Private Sub Sid_click()
myrs.Open "select * from Staff_info where Staff_id ='" & sid.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
depmt.Text = myrs!Department
post.Text = myrs!post
Staffname.Text = myrs!Staff_Name
Salary.Text = myrs!Salary
myrs.Close
End Sub


