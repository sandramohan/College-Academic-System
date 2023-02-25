VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSTUD 
   Caption         =   "Form1"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8655
   ScaleWidth      =   15255
   Begin VB.TextBox txtyr 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12480
      TabIndex        =   30
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtCid 
      Height          =   375
      Left            =   7680
      TabIndex        =   29
      Top             =   4680
      Width           =   735
   End
   Begin VB.ComboBox cmbCourse 
      Height          =   315
      Left            =   4800
      TabIndex        =   28
      Top             =   4680
      Width           =   2775
   End
   Begin VB.ComboBox cmbsearch 
      Height          =   315
      Left            =   9960
      TabIndex        =   27
      Text            =   "Student ID"
      Top             =   5160
      Width           =   3855
   End
   Begin MSComCtl2.DTPicker DTdob 
      Height          =   495
      Left            =   4800
      TabIndex        =   26
      Top             =   3000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   873
      _Version        =   393216
      Format          =   41287681
      CurrentDate     =   43371
   End
   Begin VB.OptionButton opf 
      Caption         =   "Female"
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      Top             =   2160
      Width           =   1335
   End
   Begin VB.OptionButton opm 
      Caption         =   "Male"
      Height          =   255
      Left            =   4800
      TabIndex        =   24
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdsearch 
      Caption         =   "SEARCH"
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
      Left            =   11160
      TabIndex        =   22
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "DELETE"
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
      Left            =   11040
      TabIndex        =   21
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
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
      Left            =   8760
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
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
      Left            =   13080
      TabIndex        =   19
      Top             =   2400
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
      Left            =   11040
      TabIndex        =   18
      Top             =   2400
      Width           =   1455
   End
   Begin VB.CommandButton cmdnew 
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
      Left            =   8760
      TabIndex        =   17
      Top             =   2400
      Width           =   1575
   End
   Begin VB.TextBox txtcourse 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   16
      Top             =   7920
      Width           =   2775
   End
   Begin VB.TextBox txtadd 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   15
      Top             =   7080
      Width           =   2775
   End
   Begin VB.TextBox txtguardian 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   14
      Top             =   6240
      Width           =   2775
   End
   Begin VB.TextBox txtph 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   13
      Top             =   5400
      Width           =   2775
   End
   Begin VB.TextBox txtemail 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   12
      Top             =   3840
      Width           =   2775
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   11
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   4800
      TabIndex        =   10
      Top             =   600
      Width           =   2775
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Year Of Admission"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   9120
      TabIndex        =   31
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "STUDENT DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6360
      TabIndex        =   23
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label10 
      Caption         =   "Course Duration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   9
      Top             =   8040
      Width           =   2535
   End
   Begin VB.Label Label9 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   8
      Top             =   7200
      Width           =   2535
   End
   Begin VB.Label Label8 
      Caption         =   "Guardian  Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label Label7 
      Caption         =   "Phone Number"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   6
      Top             =   5520
      Width           =   2535
   End
   Begin VB.Label Label6 
      Caption         =   "Department"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   5
      Top             =   4680
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "DOB"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   3000
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "Stud id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
End
Attribute VB_Name = "frmSTUD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim gen As String

Private Sub cmbcourse_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & cmbcourse.Text & "'", con, 3, 3
If Not rs.EOF Then
txtCid.Text = rs(0)
End If

End Sub

Private Sub cmdcancel_Click()
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdupdate.Enabled = False
cmddel.Enabled = False
cmdsearch.Enabled = True
cmdcancel.Enabled = False
End Sub

Private Sub cmddel_Click()
Dim r
r = MsgBox("Do you want to delete this record?", vbYesNo)
If r = vbYes Then
cmd.CommandText = "delete from student where studid=" & Val(studid.Text)
cmd.Execute
clr
MsgBox "record deleted successfully'"
End If
End Sub

Private Sub cmdnew_Click()
clr
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(studid),5000 )+1 from student", con, 3, 3
If Not rs.EOF Then
txtid.Text = rs(0)
End If
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
End Sub

Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(studid),5000)+1 from student", con, 3, 3
If Not rs.EOF Then
id = rs(0)
End If
End Sub

Private Sub cmdsave_Click()
If (txtid.Text = "" Or txtname.Text = "" Or txtemail.Text = "" Or cmbcourse.Text = "" Or txtph.Text = "" Or txtguardian.Text = "" Or txtadd.Text = "" Or txtcourse.Text = "") Then
MsgBox "fields cannot left blank"
Else
cmd.CommandText = "insert into student values(" & Val(txtid.Text) & ",'" & txtname.Text & "','" & gen & "','" & DTdob.Value & "','" & txtemail.Text & "','" & txtCid.Text & "'," & txtph.Text & ",'" & txtguardian.Text & "','" & txtadd.Text & "'," & Val(txtcourse.Text) & "," & Val(txtyr.Text) & ",0)"
MsgBox cmd.CommandText
cmd.Execute

MsgBox "Record Inserted Successfully"
addStud
End If
End Sub

Private Sub cmdsearch_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from student where studid=" & cmbsearch, con, 3, 3
If Not rs.EOF Then
txtid.Text = rs(0)
txtname.Text = rs(1)
DTdob.Value = rs(3)
txtemail.Text = rs(4)
txtph.Text = rs(6)
txtguardian.Text = rs(7)
txtadd.Text = rs(8)
txtcourse.Text = rs(9)
txtyr.Text = rs(10)
txtCid.Text = rs(5)
gen = rs(2)
If gen = "Male" Then
opm.Value = True
ElseIf gen = "Female" Then
opf.Value = True
End If
End If
End Sub

Private Sub cmdupdate_Click()
If txtid.Text = "" Or txtname.Text = "" Or txtemail.Text = "" Or cmbcourse.Text = "" Or txtph.Text = "" Or txtguardian.Text = "" Or txtadd.Text = "" Or txtcourse.Text = "" Then
MsgBox "fields cannot left blank"
Else
cmdcommandtext = "insert into student values(" & Val(txtid.Text) & "'" & txtname.Text & "','" & gen & "','" & DTdob.Value & "','" & txtemail.txt & "','" & txtdept.Text & "','" & txtph.Text & "','" & txtguardian.Text & "','" & txtadd.Text & "','" & Val(txtcourse.Text) & ")"
cmd.Execute
MsgBox "record inserted successfully"
End If
End Sub

Public Sub addStud()
cmbsearch.Clear
If rs.State = 1 Then rs.Close
rs.Open "student", con, 3, 3
While Not rs.EOF
cmbsearch.AddItem rs(0)
rs.MoveNext
Wend
End Sub

Private Sub Form_Load()
Module1.getcon
addStud
addcourse
DTdob.Value = Format(Now, "dd/mm/yyyy")

cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdupdate.Enabled = False
cmddel.Enabled = False
End Sub

Private Sub opf_Click()
If opf.Value = True Then
gen = "Female"
End If
End Sub

Private Sub Opm_Click()
If opm.Value = True Then
gen = "Male"
End If
End Sub
Private Sub clr()
txtname.Text = ""
txtemail.Text = ""
txtph.Text = ""
txtguardian.Text = ""
txtadd.Text = ""
txtcourse.Text = ""

End Sub
Public Sub addcourse()
cmbcourse.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "course", con, 3, 3
While Not rs1.EOF
cmbcourse.AddItem rs1(1)
rs1.MoveNext
Wend
End Sub



Private Sub txtCid_Change()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where courseid=" & txtCid.Text, con, 3, 3
If Not rs.EOF Then
cmbcourse.Text = rs(1)
End If
End Sub

Private Sub txtyr_LostFocus()
If Val(txtyr.Text) > Year(Now) Or (Year(Now) - Val(txtyr.Text)) > 2 Then
MsgBox "Cannot exceed current year or not more than 3 years"
txtyr.SetFocus
End If
End Sub

