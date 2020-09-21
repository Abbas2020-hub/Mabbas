VERSION 5.00
Begin VB.Form Billing 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Apollo Acounts & Billing"
   ClientHeight    =   8850
   ClientLeft      =   6735
   ClientTop       =   6945
   ClientWidth     =   13590
   Icon            =   "Billing.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8850
   ScaleWidth      =   13590
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Add Bill"
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
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " Billing "
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
      Height          =   5535
      Left            =   3720
      TabIndex        =   12
      Top             =   1800
      Width           =   7695
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "&Edit Bill"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C&lear Page"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4680
         Width           =   1215
      End
      Begin VB.ComboBox bno 
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
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox text1 
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
         Top             =   1680
         Width           =   2655
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   3
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text3 
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
         Top             =   3240
         Width           =   2655
      End
      Begin VB.TextBox Text4 
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
         TabIndex        =   5
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bill no:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   17
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name of the Patient:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   16
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount Paid:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   14
         Top             =   3240
         Width           =   1935
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Balance:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   840
         TabIndex        =   13
         Top             =   3960
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Bill &Print"
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
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "&Clear Bill"
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
      Left            =   5520
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7920
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
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
      Left            =   10080
      MaskColor       =   &H0080FFFF&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Accounts && Billing"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   4560
      TabIndex        =   0
      Top             =   720
      Width           =   6255
   End
End
Attribute VB_Name = "Billing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public myrs As New ADODB.Recordset

Private Sub Command1_Click()
On Error Resume Next
Unload Bill
Unload Me
Main.Show
Main.SetFocus
Main.WindowState = 2
End Sub


Private Sub Command2_Click()
If bno.Text = "" Then
MsgBox "Enter the Bill No.", vbExclamation + vbOKCancel, "Missing Bill No."
Else
If text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
MsgBox "Missing data. Enter the full details", vbCritical + vbOKCancel, "Missing Data"
Else
myrs.Open "select * from billing", conn, adOpenKeyset, adLockOptimistic, -1
myrs.AddNew
myrs!bill_no = bno.Text
myrs!patient_name = text1.Text
myrs!total_amount = Text2.Text
myrs!amount_paid = Text3.Text
myrs!balance = Text4.Text
myrs.Update
myrs.Close
myrs.Open "select * from billing where bill_no =" & bno.Text & " ", conn, adOpenKeyset, adLockOptimistic, -1
Call clear_focus
While Not myrs.EOF
bno.AddItem (myrs!bill_no)
myrs.MoveNext
Wend
myrs.Close
End If
End If
End Sub

Private Sub Command3_Click()
If bno.Text = "" Then
MsgBox "There is no current record to delete. Please select a Bill No.", vbExclamation + vbOKOnly, "Missing..."
bno.SetFocus
Exit Sub
End If
myrs.Open "select * from billing where bill_no = '" & bno.Text & "' ", conn, adOpenKeyset, adLockOptimistic, -1
myrs.Delete
MsgBox "Entry Deleted!", vbOKOnly + vbExclamation, "Success..."
myrs.Close
Call clear_focus
Unload Me
Main.Show
Main.WindowState = 2
End Sub

Private Sub Command4_Click()
Call Bill.Show
End Sub

Private Sub Command5_Click()
Call clear_focus
End Sub

Private Sub Command6_Click()
myrs.Open "select * from billing where bill_no = '" & bno & "'", conn, adOpenStatic, adLockOptimistic, -1
myrs!bill_no = Trim(bno.Text)
myrs!patient_name = Trim(text1.Text)
myrs!total_amount = Trim(Text2.Text)
myrs!amount_paid = Trim(Text3.Text)
myrs!balance = Trim(Text4.Text)
myrs.Update
MsgBox "Record Modified Successfully", vbInformation, "Success"
myrs.Close
myrs.Open "select * from billing where bill_no = '" & bno & "'", conn, adOpenStatic, adLockOptimistic, -1
text1.Text = myrs!patient_name
Text2.Text = myrs!total_amount
Text3.Text = myrs!amount_paid
Text4.Text = myrs!balance
myrs.Close
Call clear_focus
End Sub

Private Sub Form_Load()
Call connect
myrs.Open "select * from billing", conn, adOpenKeyset, adLockOptimistic, -1
While Not myrs.EOF
bno.AddItem (myrs!bill_no)
myrs.MoveNext
Wend
myrs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
conn.Close
End Sub

Private Sub bno_click()
myrs.Open "select * from billing where bill_no = '" & bno.Text & "'", conn, adOpenKeyset, adLockOptimistic, -1
text1.Text = myrs!patient_name
Text2.Text = myrs!total_amount
Text3.Text = myrs!amount_paid
Text4.Text = myrs!balance
myrs.Close
End Sub

Private Sub Text3_Change()
Text4.Text = Val(Text2.Text) - Val(Text3.Text)
End Sub

Private Function clear_focus()
bno.Text = ""
text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
bno.SetFocus
End Function

