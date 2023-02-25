VERSION 5.00
Begin VB.Form frmclsthr 
   Caption         =   "Form1"
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11220
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   11220
   Begin VB.TextBox Txtcrsid 
      Height          =   495
      Left            =   7680
      TabIndex        =   13
      Top             =   2520
      Width           =   735
   End
   Begin VB.TextBox Txtstaff 
      Height          =   495
      Left            =   7680
      TabIndex        =   12
      Top             =   4440
      Width           =   735
   End
   Begin VB.ComboBox Cmbcrs 
      Height          =   315
      Left            =   5040
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtctid 
      Height          =   495
      Left            =   5040
      TabIndex        =   9
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7080
      TabIndex        =   7
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmdnew 
      Appearance      =   0  'Flat
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   5
      Top             =   5760
      Width           =   1455
   End
   Begin VB.ComboBox cmbstaff 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Top             =   4560
      Width           =   2535
   End
   Begin VB.ComboBox cmbsem 
      Height          =   315
      ItemData        =   "frmclsthr.frx":0000
      Left            =   5040
      List            =   "frmclsthr.frx":0016
      TabIndex        =   3
      Top             =   3600
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Course"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.Label Label5 
      Caption         =   "CLASS TEACHER"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3120
      TabIndex        =   8
      Top             =   120
      Width           =   5655
   End
   Begin VB.Label St 
      Caption         =   "Staff Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   2
      Top             =   4440
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Semester"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   3480
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "CTid"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   960
      TabIndex        =   0
      Top             =   1440
      Width           =   3255
   End
End
Attribute VB_Name = "frmclsthr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub cmbstaff_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from staff where name='" & cmbstaff.Text & "'", con, 3, 3
If Not rs.EOF Then
txtStaff.Text = rs(0)
End If
End Sub

Private Sub cmdcancel_Click()
cmdsave.Enabled = False
cmdnew.Enabled = True
End Sub

Private Sub cmdnew_Click()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(ctid),7000 )+1 from classteacher", con, 3, 3
If Not rs.EOF Then
txtctid.Text = rs(0)
End If
getId
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
End Sub

Private Sub cmdsave_Click()
If txtctid.Text = "" Or Cmbcrs.Text = "" Or cmbsem.Text = "" Or cmbstaff.Text = "" Then
MsgBox " fields cannot left blank "
Else
cmd.CommandText = "insert into classteacher values(" & Val(txtctid.Text) & ",'" & Txtcrsid.Text & "','" & cmbsem.Text & "','" & Val(txtStaff.Text) & "')"
cmd.Execute
MsgBox "record inserted successfully"
End If
End Sub

Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(ctid),7000)+1 from classteacher", con, 3, 3
If Not rs.EOF Then
txtctid = rs(0)
End If
End Sub

Private Sub Form_Load()
Module1.getcon
cmdsave.Enabled = False
cmdcancel.Enabled = False
addcrs
addstaff
End Sub

Private Sub cmbcrs_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & Cmbcrs.Text & "'", con, 3, 3
If Not rs.EOF Then
Txtcrsid.Text = rs(0)
End If
End Sub

Public Sub addcrs()
Cmbcrs.Clear
If rs.State = 1 Then rs.Close
rs.Open "course", con, 3, 3
While Not rs.EOF
Cmbcrs.AddItem rs(1)
rs.MoveNext
Wend
End Sub

Public Sub addstaff()
cmbstaff.Clear
If rs.State = 1 Then rs.Close
rs.Open "staff", con, 3, 3
While Not rs.EOF
cmbstaff.AddItem rs(1)
rs.MoveNext
Wend
End Sub
