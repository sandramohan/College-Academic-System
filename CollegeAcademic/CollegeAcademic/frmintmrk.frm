VERSION 5.00
Begin VB.Form frmintmrk 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15285
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8415
   ScaleWidth      =   15285
   Begin VB.TextBox txtadno 
      Height          =   285
      Left            =   3840
      TabIndex        =   14
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtCid 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Top             =   1320
      Width           =   495
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
      TabIndex        =   12
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ComboBox cmbsemester 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frmintmrk.frx":0000
      Left            =   8280
      List            =   "frmintmrk.frx":0016
      TabIndex        =   11
      Top             =   1200
      Width           =   2535
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
      Left            =   8280
      TabIndex        =   10
      Top             =   1920
      Width           =   2535
   End
   Begin VB.TextBox txtattend 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtassign 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton cmdnw 
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
      Height          =   720
      Left            =   2400
      TabIndex        =   7
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdsave 
      Appearance      =   0  'Flat
      Caption         =   "SAVE"
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
      Height          =   735
      Left            =   5520
      TabIndex        =   6
      Top             =   5520
      Width           =   1455
   End
   Begin VB.CommandButton cmdcancel 
      Appearance      =   0  'Flat
      Caption         =   "CANCEL"
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
      Height          =   735
      Left            =   8520
      TabIndex        =   5
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtInt2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   4
      Top             =   4320
      Width           =   2535
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
      TabIndex        =   3
      Top             =   2040
      Width           =   2055
   End
   Begin VB.TextBox txtInt1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   2
      Top             =   4320
      Width           =   2055
   End
   Begin VB.TextBox txtStaff 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2640
      Width           =   2535
   End
   Begin VB.TextBox txtWDays 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Text            =   "1"
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " INTERNAL  MARK"
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
      Left            =   2400
      TabIndex        =   29
      Top             =   240
      Width           =   6855
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
      TabIndex        =   28
      Top             =   1320
      Width           =   2055
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
      Left            =   5400
      TabIndex        =   27
      Top             =   1200
      Width           =   1935
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
      TabIndex        =   26
      Top             =   2040
      Width           =   1935
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
      Left            =   5400
      TabIndex        =   25
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attendance"
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
      TabIndex        =   24
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Assignment/Seminar"
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
      Left            =   5400
      TabIndex        =   23
      Top             =   3600
      Width           =   2895
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internal2"
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
      Left            =   5400
      TabIndex        =   22
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Internal1"
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
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Label Label9 
      Caption         =   "/5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   20
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label10 
      Caption         =   "/5"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11040
      TabIndex        =   19
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Label12 
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
      Height          =   375
      Left            =   4560
      TabIndex        =   18
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label13 
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
      Height          =   375
      Left            =   10920
      TabIndex        =   17
      Top             =   4440
      Width           =   615
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Faculty"
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
      Left            =   5400
      TabIndex        =   16
      Top             =   2760
      Width           =   2895
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Working Days"
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
      TabIndex        =   15
      Top             =   2760
      Width           =   2055
   End
End
Attribute VB_Name = "frmintmrk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim rs4 As New ADODB.Recordset
Dim rs5 As New ADODB.Recordset
Dim lst As ListItem
Dim id As Integer
Dim no, attnd As Integer
Dim nosub, totmks1, totmks2 As Integer

Private Sub cmbadno_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from student where studid=" & cmbadno.Text, con, 3, 3
If Not rs.EOF Then
txtname.Text = rs(1)
End If
GetDetails
End Sub

Private Sub cmbcourse_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & cmbcourse.Text & "'", con, 3, 3
If Not rs.EOF Then
txtCid.Text = rs(0)
End If
End Sub

Private Sub cmbsemester_Click()
adno
End Sub

Private Sub cmbsubject_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from SUBDETAILS1 where subname='" & cmbsubject.Text & "'", con, 3, 3
If Not rs.EOF Then
txtsub.Text = rs(0)
End If
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
End Sub

Private Sub cmdnw_Click()
clr
cmdsave.Enabled = True
cmdcancel.Enabled = True
cmdnw.Enabled = False

End Sub



Private Sub cmdsave_Click()
If txtCid.Text = "" Or cmbsemester.Text = "" Or cmbadno.Text = "" Or txtattend.Text = "" Or txtassign.Text = "" Or txtInt1.Text = "" Or txtInt2.Text = "" Then
MsgBox "Fields cannot left blank"
Else
If rs5.State = 1 Then rs5.Close
rs5.Open "select * from internalmark where CID=" & txtCid.Text & " and sem=" & cmbsemester.Text & " and adno=" & cmbadno.Text, con, 3, 3
If Not rs5.EOF Then
cmd.CommandText = "update internalmark set attendence=" & txtattend.Text & ",seminar=" & txtassign.Text & ",internal1=" & txtInt1.Text & ",internal2=" & txtInt2.Text & " where cid=" & txtCid.Text & " and sem=" & cmbsemester.Text & " and adno=" & cmbadno.Text
Else
cmd.CommandText = "insert into internalmark values(" & txtCid.Text & "," & cmbsemester.Text & "," & cmbadno.Text & "," & txtattend.Text & "," & txtassign.Text & "," & txtInt1.Text & "," & txtInt2.Text & ")"
'MsgBox cmd.CommandText
End If
cmd.Execute
MsgBox "Marks added successfully"

cmdsave.Enabled = False
cmdcancel.Enabled = False
cmdnw.Enabled = True
Unload Me
End If
End Sub


Private Sub Form_Load()
Module1.getcon
addcourse
getId
txtStaff.Text = Module1.id
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
rs2.Open "select student.studid from student,ssemallocation where student.studid=ssemallocation.studid and courseid=" & txtCid.Text & " and SEM=" & cmbsemester.Text, con, 3, 3
While Not rs2.EOF
cmbadno.AddItem rs2(0)

rs2.MoveNext

Wend
End Sub


Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(cid),11000)+1 from internalmark", con, 3, 3
If Not rs.EOF Then
id = rs(0)

End If
End Sub


Public Sub GetDetails()

If rs2.State = 1 Then rs2.Close
rs2.Open "select SUM(mark) from academic where studid=" & cmbadno.Text & " and internal=1", con, 3, 3
If Not IsNull(rs2(0)) Then
totmks1 = rs2(0)
Else
totmks1 = 0
End If

If rs4.State = 1 Then rs4.Close
rs4.Open "select SUM(mark) from academic where studid=" & cmbadno.Text & " and internal=2", con, 3, 3
If Not IsNull(rs4(0)) Then
totmks2 = rs4(0)
Else
totmks2 = 0
End If


If rs2.State = 1 Then rs2.Close
rs2.Open "select COUNT(*) from coursecrs where CourseID=" & txtCid.Text & " and SEM=" & cmbsemester.Text, con, 3, 3
On Error GoTo err1
If Not rs2.EOF Then
nosub = rs2(0)
End If
err1:
txtInt1.Text = 0
txtInt2.Text = 0

txtInt1.Text = totmks1 / (nosub * 10) * 10
txtInt2.Text = totmks2 / (nosub * 10) * 10
End Sub

Public Sub GetAttendance()
txtattend.Text = ""
If rs3.State = 1 Then rs3.Close
rs3.Open "select COUNT(*) from attendance where COURSEID=" & txtCid.Text & " and SEM=" & cmbsemester.Text & " and studid=" & cmbadno.Text, con, 3, 3
'On Error GoTo err
If Not IsNull(rs3(0)) Then
no = rs3(0)
Else
no = 0
End If

attnd = no / Val(txtWDays.Text) * 5
txtattend.Text = attnd
End Sub

Public Sub clr()
cmbcourse.Text = ""
txtCid.Text = ""
cmbsemester.Text = ""
cmbadno.Text = ""
txtname.Text = ""
txtattend.Text = ""
txtInt1.Text = ""
txtInt2.Text = ""
txtassign.Text = ""
'txtWDays.Text = "0"
End Sub

Private Sub txtattend_GotFocus()
GetAttendance
End Sub

