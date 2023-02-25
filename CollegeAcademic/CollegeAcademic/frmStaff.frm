VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmStaff 
   Caption         =   "Form1"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   19035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9000
   ScaleWidth      =   19035
   Begin VB.TextBox txtPwd 
      Appearance      =   0  'Flat
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   13920
      PasswordChar    =   "*"
      TabIndex        =   28
      Top             =   6960
      Width           =   2415
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   13920
      TabIndex        =   27
      Top             =   6120
      Width           =   2415
   End
   Begin VB.CommandButton cmdsearch 
      Appearance      =   0  'Flat
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
      Left            =   12120
      TabIndex        =   26
      Top             =   2040
      Width           =   1455
   End
   Begin VB.ComboBox cmbsearch 
      Height          =   315
      Left            =   11520
      TabIndex        =   25
      Text            =   "Staff ID"
      Top             =   1560
      Width           =   2895
   End
   Begin MSComCtl2.DTPicker DTPdob 
      Height          =   375
      Left            =   7560
      TabIndex        =   23
      Top             =   3120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   43357
   End
   Begin VB.TextBox txtph 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   22
      Top             =   4560
      Width           =   2415
   End
   Begin VB.TextBox txtmail 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   21
      Top             =   3720
      Width           =   2415
   End
   Begin VB.OptionButton opF 
      Caption         =   "Female"
      Height          =   495
      Left            =   8880
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.OptionButton OpM 
      Caption         =   "Male"
      Height          =   495
      Left            =   7680
      TabIndex        =   17
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.TextBox txtexp 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   5280
      Width           =   2415
   End
   Begin VB.TextBox txtdes 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   6840
      Width           =   2415
   End
   Begin VB.TextBox txtqua 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   7560
      TabIndex        =   5
      Top             =   7560
      Width           =   2415
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
      Left            =   11040
      TabIndex        =   4
      Top             =   3360
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
      Left            =   12600
      TabIndex        =   3
      Top             =   3360
      Width           =   1455
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
      Left            =   10320
      TabIndex        =   2
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton cmdupd 
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
      Left            =   11880
      TabIndex        =   1
      Top             =   4200
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
      Left            =   13560
      TabIndex        =   0
      Top             =   4200
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPdoj 
      Height          =   375
      Left            =   7560
      TabIndex        =   24
      Top             =   6120
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Format          =   118554625
      CurrentDate     =   43357
   End
   Begin VB.Label Label12 
      Caption         =   "Password"
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
      Left            =   10680
      TabIndex        =   30
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label Label11 
      Caption         =   "UserName"
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
      Left            =   10680
      TabIndex        =   29
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label Label10 
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
      Left            =   2520
      TabIndex        =   20
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Label Label9 
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
      Left            =   2520
      TabIndex        =   19
      Top             =   3840
      Width           =   2775
   End
   Begin VB.Label Label8 
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
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Width           =   2775
   End
   Begin VB.Label Label7 
      Caption         =   "Date of birth"
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
      Left            =   2520
      TabIndex        =   15
      Top             =   3120
      Width           =   2775
   End
   Begin VB.Label Label6 
      Caption         =   "STAFF REGISTRATION"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   14
      Top             =   360
      Width           =   7695
   End
   Begin VB.Label Label1 
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
      Left            =   2520
      TabIndex        =   13
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "Experience"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label3 
      Caption         =   "Date of joining"
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
      Left            =   2520
      TabIndex        =   11
      Top             =   6000
      Width           =   2775
   End
   Begin VB.Label Label4 
      Caption         =   "Designation"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   6720
      Width           =   2775
   End
   Begin VB.Label Label5 
      Caption         =   "Qualification"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   7560
      Width           =   2775
   End
End
Attribute VB_Name = "frmStaff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gen As String
Dim id As Integer
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub cmdcancel_Click()
clr
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdupd.Enabled = False
cmddel.Enabled = False
End Sub

Private Sub cmddel_Click()
Dim r
r = MsgBox("do you want to delete this record ?", vbYesNo)
If r = vbYes Then
cmd.CommandText = "delete from staff where name=" & Val(txtname.Text)
cmd.Execute
clr
MsgBox "record deleted successfully"
End If
End Sub

Private Sub cmdnew_Click()
clr
getId
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
End Sub

Private Sub cmdsave_Click()
If txtname.Text = "" Or txtmail.Text = "" Or txtph.Text = "" Or txtexp.Text = "" Or txtdes.Text = "" Or txtqua.Text = "" Or txtUser.Text = "" Or txtPwd.Text = "" Then
MsgBox "fields cannot left blank"
Else
cmd.CommandText = "insert into staff values(" & id & ",'" & txtname.Text & "',' " & gen & "','" & DTPdob.Value & "','" & txtmail.Text & "'," & txtph.Text & "," & txtexp.Text & ",'" & DTPdoj.Value & "','" & txtdes.Text & "','" & txtqua.Text & "',0)"
MsgBox cmd.CommandText
cmd.Execute
cmd.CommandText = "insert into login values('" & txtUser.Text & "','" & txtPwd.Text & "','Staff'," & id & ")"
cmd.Execute
MsgBox "Record inserted successfully"
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdnew.Enabled = True
addstaff
End If
End Sub


Private Sub cmdsearch_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from staff where StaffId =" & cmbsearch.Text, con, 3, 3
If Not rs.EOF Then
txtname.Text = rs(1)
txtmail.Text = rs(4)
txtph.Text = rs(5)
txtexp.Text = rs(6)
DTPdoj.Value = rs(7)
txtdes.Text = rs(8)
txtqua.Text = rs(9)
gen = rs(2)
If gen = "M" Then
opm.Value = True
ElseIf gen = "F" Then
opf.Value = True
End If
cmddel.Enabled = True
cmdupd.Enabled = True
End If
End Sub

Private Sub cmdupd_Click()
If txtname.Text = "" Or txtmail.Text = "" Or txtph.Text = "" Or txtexp.Text = "" Or DTPdoj = "" Or txtdes.Text = "" Or txtqua.Text = "" Then
MsgBox "fields cannot be left blank"
Else
cmd.CommandText = "update staff set Name='" & (txtname.Text) & "',DOB='" & (DTPdob.Value) & "',Email = '" & (txtmail.Text) & "',Phno='" & (txtph.Text) & "',Experience=" & (txtexp.Text) & ",DOJ='" & (DTPdoj.Value) & "',Designation='" & (txtdes.Text) & "',Qualification='" & (txtqua.Text) & "' where staffid=" & cmbsearch.Text
MsgBox cmd.CommandText
cmd.Execute
MsgBox "record updated successfully"
End If
End Sub

Private Sub Form_Load()
Module1.getcon
addstaff
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdupd.Enabled = False
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
txtmail.Text = ""
txtph.Text = ""
txtexp.Text = ""
txtdes.Text = ""
txtqua.Text = ""
txtUser.Text = ""
txtPwd.Text = ""
End Sub

Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(staffid),2000)+1 from staff", con, 3, 3
If Not rs.EOF Then
id = rs(0)
End If
End Sub

Public Sub addstaff()
cmbsearch.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "staff", con, 3, 3
While Not rs1.EOF
cmbsearch.AddItem rs1(0)
rs1.MoveNext
Wend
End Sub

