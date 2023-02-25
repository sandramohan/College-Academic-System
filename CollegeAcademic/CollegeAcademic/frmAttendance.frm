VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAttendance 
   Caption         =   $"frmAttendance.frx":0000
   ClientHeight    =   6525
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10965
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6525
   ScaleWidth      =   10965
   Begin MSComctlLib.ListView lstView 
      Height          =   3735
      Left            =   7080
      TabIndex        =   14
      Top             =   1200
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   6588
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Student Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.ComboBox cmbco 
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
      Left            =   3840
      TabIndex        =   7
      Top             =   1200
      Width           =   1815
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
      ItemData        =   "frmAttendance.frx":010E
      Left            =   3840
      List            =   "frmAttendance.frx":0124
      TabIndex        =   6
      Top             =   2040
      Width           =   1815
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
      ItemData        =   "frmAttendance.frx":013A
      Left            =   3840
      List            =   "frmAttendance.frx":013C
      TabIndex        =   5
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cmbsession 
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
      ItemData        =   "frmAttendance.frx":013E
      Left            =   3840
      List            =   "frmAttendance.frx":0148
      TabIndex        =   4
      Top             =   3600
      Width           =   1815
   End
   Begin VB.TextBox txtdate 
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
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdmark 
      Appearance      =   0  'Flat
      Caption         =   "MARK ATTENDENCE"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.TextBox txtcourse 
      Height          =   405
      Left            =   5760
      TabIndex        =   1
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtsub 
      Height          =   405
      Left            =   5760
      TabIndex        =   0
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "ATTENDENCE"
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
      Left            =   3840
      TabIndex        =   13
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label cmbcourse 
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
      Height          =   495
      Left            =   2040
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label4 
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
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Session"
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
      Height          =   495
      Left            =   2040
      TabIndex        =   9
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   4440
      Width           =   2055
   End
End
Attribute VB_Name = "frmAttendance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim lst As ListItem
Dim id As Integer

Private Sub cmbco_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & cmbco.Text & "'", con, 3, 3
If Not rs.EOF Then
txtcourse.Text = rs(0)
End If
End Sub

Private Sub cmbsemester_Click()
addsubject
txtsub.Text = ""
AddStudents
End Sub

Private Sub cmbsubject_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUBJECT where SUBNAME='" & cmbsubject.Text & "'", con, 3, 3
If Not rs.EOF Then
txtsub.Text = rs(0)
End If

End Sub

Private Sub cmdmark_Click()
For i = 1 To lstview.ListItems.Count
getId
If lstview.ListItems(i).Checked Then
Set lst = lstview.ListItems(i)
cmd.CommandText = "insert into attendance values(" & id & "," & txtcourse.Text & "," & cmbsemester.Text & "," & txtsub.Text & ",'" & cmbsession.Text & "','" & txtdate.Text & "'," & lstview.ListItems(i) & "," & Module1.id & ")"
cmd.Execute
End If
Next
MsgBox "Attendance marked successfully"
End Sub

Private Sub Form_Load()
Module1.getcon
addcourse
getId
txtdate.Text = Format(Now, "dd/mm/yyyy")
End Sub


Public Sub addcourse()
cmbco.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "course", con, 3, 3
While Not rs1.EOF
cmbco.AddItem rs1(1)
rs1.MoveNext
Wend
End Sub

Public Sub addsubject()
cmbsubject.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "SELECT SUBNAME FROM SUBJECT WHERE SUBID IN( SELECT SUBID FROM staff WHERE staffid=" & Module1.id & " and SubId IN(SELECT SubId FROM coursecrs WHERE courseid=" & txtcourse.Text & " and sem= " & cmbsemester.Text & "))", con, 3, 3
While Not rs1.EOF
cmbsubject.AddItem rs1(0)
rs1.MoveNext
Wend
End Sub


Public Sub AddStudents()
lstview.ListItems.Clear
If rs2.State = 1 Then rs2.Close
rs2.Open "select student.studid,name from student,ssemallocation where student.studid=ssemallocation.studid and courseid=" & txtcourse.Text & " and sem=" & cmbsemester.Text, con, 3, 3
While Not rs2.EOF
Set lst = lstview.ListItems.Add(, , rs2(0))
lst.SubItems(1) = rs2(1)
rs2.MoveNext
Wend
End Sub

Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(AID),8000)+1 from attendance", con, 3, 3
If Not rs.EOF Then
id = rs(0)
cmdmark.Enabled = True
End If
End Sub

