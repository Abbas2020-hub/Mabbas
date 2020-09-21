VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form medicine 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Medicine Details"
   ClientHeight    =   7200
   ClientLeft      =   5865
   ClientTop       =   6945
   ClientWidth     =   7095
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "medicine.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   73611.35
   ScaleMode       =   0  'User
   ScaleWidth      =   3.24426e5
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   " Medicine Details "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   5895
      Left            =   4200
      TabIndex        =   10
      Top             =   1680
      Width           =   7815
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFC0&
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
         Left            =   840
         MaskColor       =   &H00FFFFC0&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   5160
         Width           =   1455
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
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
         MaskColor       =   &H00FFFFC0&
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5160
         Width           =   1455
      End
      Begin VB.ComboBox medcn_nm 
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
         Left            =   3960
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox manf 
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
         Left            =   3960
         TabIndex        =   2
         Top             =   2160
         Width           =   2775
      End
      Begin VB.TextBox exp 
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
         Left            =   3960
         TabIndex        =   3
         Top             =   3240
         Width           =   2775
      End
      Begin VB.TextBox pri 
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
         Left            =   3960
         TabIndex        =   4
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label Med_name 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of the Medicine:"
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
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label manf_date 
         BackStyle       =   0  'Transparent
         Caption         =   "Manufacture Date:"
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
         Left            =   840
         TabIndex        =   13
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label Exp_date 
         BackStyle       =   0  'Transparent
         Caption         =   "Expiry Date:"
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
         Left            =   840
         TabIndex        =   12
         Top             =   3240
         Width           =   2295
      End
      Begin VB.Label med_price 
         BackStyle       =   0  'Transparent
         Caption         =   "Price:"
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
         Left            =   840
         TabIndex        =   11
         Top             =   4320
         Width           =   1215
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   -240
      Top             =   -360
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=C5-15"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=Hospital;Data Source=C5-15"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Trash Entry"
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
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Add Entry"
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
      Left            =   5040
      MaskColor       =   &H00FFFFC0&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8160
      Width           =   1455
   End
   Begin VB.CommandButton End 
      BackColor       =   &H00C0FFC0&
      Caption         =   "E&xit"
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
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8160
      Width           =   1455
   End
   Begin VB.Label Medicines 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Medicine Details && Info"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   5040
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "medicine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public myrs As New ADODB.Recordset


Private Sub Command3_Click()
Call clear_focus
End Sub

Private Sub Command4_Click()
myrs.Open "select * from Medicinal_stores where name_of_the_medicine = '" & medcn_nm & "'", conn, adOpenStatic, adLockOptimistic, -1
myrs!name_of_the_medicine = Trim(medcn_nm.Text)
myrs!manufacture_date = Trim(manf.Text)
myrs!expiry_date = Trim(exp.Text)
myrs!price = Trim(pri.Text)
myrs.Update
MsgBox "Record Modified Successfully", vbInformation, "Success"
myrs.Close
myrs.Open "select * from Medicinal_stores where name_of_the_medicine = '" & medcn_nm & "'", conn, adOpenStatic, adLockOptimistic, -1
manf.Text = myrs!manufacture_date
exp.Text = myrs!expiry_date
pri.Text = myrs!price
myrs.Close
Call clear_focus
End Sub

Private Sub end_Click()
Unload Me
End Sub



Private Sub Command1_Click()
If medcn_nm.Text = "" Then
MsgBox "Enter the Name of the medicine.", vbExclamation + vbApplicationModal, "Medicine name missing..."
Else
If manf.Text = "" Or exp.Text = "" Or pri.Text = "" Then
MsgBox "Some data missing. Enter the full details", vbCritical + vbOKCancel, "Missing Data!"
Else
myrs.Open "select * from medicinal_stores", conn, adOpenKeyset, adLockOptimistic, -1
myrs.AddNew
myrs!name_of_the_medicine = medcn_nm.Text
myrs!manufacture_date = manf.Text
myrs!expiry_date = exp.Text
myrs!price = pri.Text
myrs.Update
myrs.Close
myrs.Open "select * from medicinal_stores where name_of_the_medicine = '" & medcn_nm & " ", conn, adOpenKeyset, adLockOptimistic, -1
Call clear_focus
While Not myrs.EOF
medcn_nm.AddItem (myrs!name_of_the_medicine)
myrs.MoveNext
Wend
myrs.Close
End If
End If
End Sub

Private Sub Command2_Click()
If medcn_nm.Text = "" Then
MsgBox "There is no current record to delete. Please select a Medicine Name.", vbExclamation + vbOKOnly, "Missing..."
medcn_nm.SetFocus
Exit Sub
End If
myrs.Open "select * from medicinal_stores where name_of_the_medicine = '" & medcn_nm & "' ", conn, adOpenKeyset, adLockOptimistic, -1
myrs.Delete
MsgBox "Entry Deleted!", vbOKOnly + vbExclamation, "Success..."
myrs.Close
Call clear_focus
Unload Me
Main.Show
Main.WindowState = 2
End Sub

Private Sub Form_Load()
'conn.Close
Call connect
myrs.Open "select * from medicinal_stores", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
medcn_nm.AddItem (myrs!name_of_the_medicine)
myrs.MoveNext
Wend
myrs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
End Sub

Private Sub medcn_nm_click()
myrs.Open "select * from medicinal_stores where name_of_the_medicine = '" & medcn_nm.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
medcn_nm.Text = myrs!name_of_the_medicine
manf.Text = myrs!manufacture_date
exp.Text = myrs!expiry_date
pri.Text = myrs!price
myrs.Close
End Sub


Private Function clear_focus()
medcn_nm.Text = ""
manf.Text = ""
exp.Text = ""
pri.Text = ""
medcn_nm.SetFocus
End Function
