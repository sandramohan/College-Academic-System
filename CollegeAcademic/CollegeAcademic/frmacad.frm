VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Frmacad 
   Caption         =   "Form1"
   ClientHeight    =   8670
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8670
   ScaleWidth      =   11715
   Begin MSComctlLib.ListView lstview 
      Height          =   1575
      Left            =   2280
      TabIndex        =   23
      Top             =   5520
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2778
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Sub Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Sub Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Mark"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Internal"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox txtadno 
      Height          =   285
      Left            =   3720
      TabIndex        =   13
      Top             =   1920
      Width           =   495
   End
   Begin VB.TextBox txtCid 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Top             =   1200
      Width           =   495
   End
   Begin VB.ComboBox cmbadno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   11
      Top             =   1920
      Width           =   2055
   End
   Begin VB.TextBox txtsub 
      Height          =   375
      Left            =   4680
      TabIndex        =   10
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton cmdinternal 
      Appearance      =   0  'Flat
      Caption         =   "ADD MORE INTERNAL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   4920
      TabIndex        =   9
      Top             =   3480
      Width           =   1695
   End
   Begin VB.TextBox txtmark 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      TabIndex        =   8
      Top             =   2640
      Width           =   2055
   End
   Begin VB.CommandButton cmdcancel 
      Appearance      =   0  'Flat
      Caption         =   "CANCEL"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      TabIndex        =   7
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      Caption         =   "SAVE"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   6
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CommandButton cmdnw 
      Appearance      =   0  'Flat
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2520
      TabIndex        =   5
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ComboBox cmbsubject 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   2055
   End
   Begin VB.ComboBox cmbsemester 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmacad.frx":0000
      Left            =   8640
      List            =   "frmacad.frx":0016
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.ComboBox cmbcourse 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   2400
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.ComboBox cmbInt 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmacad.frx":002C
      Left            =   2400
      List            =   "frmacad.frx":0036
      TabIndex        =   0
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Mark"
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
      Height          =   375
      Left            =   5760
      TabIndex        =   22
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject"
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
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   2760
      Width           =   2175
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      TabIndex        =   20
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Student"
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
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   5760
      TabIndex        =   18
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " Course"
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
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ACADEMIC DETAILS"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   3360
      TabIndex        =   16
      Top             =   120
      Width           =   7215
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "/10"
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
      Height          =   375
      Left            =   10800
      TabIndex        =   15
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label frmacad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internal"
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
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   2175
   End
End
Attribute VB_Name = "Frmacad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim lst As ListItem
Dim id As Integer


Private Sub cmbadno_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from student where studid=" & cmbadno.Text, con, 3, 3
If Not rs.EOF Then
txtname.Text = rs(1)
End If
End Sub

Private Sub cmbcourse_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & cmbcourse.Text & "'", con, 3, 3
If Not rs.EOF Then
txtCid.Text = rs(0)
End If
End Sub

Private Sub cmbsemester_Click()
txtsub.Text = ""
adno
addsubject
End Sub

Private Sub cmbsubject_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from subject where subname='" & cmbsubject.Text & "'", con, 3, 3
If Not rs.EOF Then
txtsub.Text = rs(0)
End If
txtmark.Text = ""
cmbInt.Text = ""
End Sub

Private Sub cmdcancel_Click()
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdnw.Enabled = True
End Sub



Private Sub cmdinternal_Click()
Set lst = lstview.ListItems.Add(, , txtsub.Text)
lst.SubItems(1) = cmbsubject.Text
lst.SubItems(2) = txtmark.Text
lst.SubItems(3) = cmbInt.Text
cmbcourse.Enabled = False
cmbsemester.Enabled = False
cmbadno.Enabled = False
cmdsave.Enabled = True
End Sub

Private Sub cmdnw_Click()
lstview.ListItems.Clear
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(subid),9000)+1 from academic", con, 3, 3
If Not rs.EOF Then
txtCid.Text = rs(0)
cmdinternal.Enabled = True
cmdcancel.Enabled = True
cmdnw.Enabled = False
End If
End Sub

Private Sub cmdsave_Click()
For i = 1 To lstview.ListItems.Count
Set lst = lstview.ListItems(i)
cmd.CommandText = "insert into academic values(" & cmbadno.Text & "," & lstview.ListItems(i) & "," & lst.SubItems(2) & "," & lst.SubItems(3) & ")"
MsgBox cmd.CommandText
cmd.Execute
Next
cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdnw.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
Module1.getcon
addcourse
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

Public Sub adno()
cmbadno.Clear
If rs2.State = 1 Then rs2.Close
rs2.Open "select student.studid from student,ssemallocation where student.studid=ssemallocation.studid and courseid=" & txtCid.Text & " and SEM=" & cmbsemester.Text & "", con, 3, 3
While Not rs2.EOF
cmbadno.AddItem rs2(0)
rs2.MoveNext
Wend
End Sub

Public Sub addsubject()
cmbsubject.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "SELECT SUBNAME FROM SUBject WHERE SUBID IN( SELECT SubId FROM coursecrs WHERE CourseID=" & txtCid.Text & " AND SEM=" & cmbsemester.Text & ")", con, 3, 3
While Not rs1.EOF
cmbsubject.AddItem rs1(0)
rs1.MoveNext
Wend
End Sub

