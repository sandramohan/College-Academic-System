VERSION 5.00
Begin VB.Form frmStudSem 
   Caption         =   "Form1"
   ClientHeight    =   8355
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8355
   ScaleWidth      =   12240
   Begin VB.TextBox txtyr 
      Height          =   285
      Left            =   7080
      TabIndex        =   5
      Top             =   2040
      Width           =   495
   End
   Begin VB.TextBox txtname 
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.TextBox txtcourse 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   2775
   End
   Begin VB.TextBox txtsemester 
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton cmdalot 
      Caption         =   "ALOT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3720
      TabIndex        =   1
      Top             =   5280
      Width           =   2055
   End
   Begin VB.ComboBox cmbadno 
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
      Left            =   5040
      TabIndex        =   0
      Top             =   2040
      Width           =   2775
   End
   Begin VB.Label Label1 
      Caption         =   "STUDENT  SEMESTER  ALLOCATION"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   10
      Top             =   720
      Width           =   11655
   End
   Begin VB.Label Label2 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   9
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label5 
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
      Height          =   375
      Left            =   3240
      TabIndex        =   6
      Top             =   4440
      Width           =   1335
   End
End
Attribute VB_Name = "frmStudSem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub cmbadno_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from student where studid=" & cmbadno.Text, con, 3, 3
If Not rs.EOF Then
txtname.Text = rs(1)
txtyr.Text = rs(10)
End If
If rs.State = 1 Then rs.Close
rs.Open "select name from COURSE where courseid=(select courseid from student where studid=" & cmbadno.Text & ")", con, 3, 3
If Not rs.EOF Then
txtcourse.Text = rs(0)
End If
GetSem
End Sub



Private Sub cmdalot_Click()
If cmbadno.Text = "" Or txtsemester.Text = "" Then
MsgBox "Cannot left fields blank"
Else
If rs.State = 1 Then rs.Close
rs.Open "select * from ssemallocation where studid=" & cmbadno.Text, con, 3, 3
If Not rs.EOF Then
cmd.CommandText = "update ssemallocation set sem=" & txtsemester.Text & " where adno=" & cmbadno.Text
Else
cmd.CommandText = "insert into ssemallocation values(" & cmbadno.Text & "," & txtsemester.Text & ")"
End If
cmd.Execute
MsgBox "record inserted successfully"
clr
End If

End Sub

Private Sub Form_Load()
Module1.getcon
adno
End Sub


Public Sub adno()
cmbadno.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "student", con, 3, 3
While Not rs1.EOF
cmbadno.AddItem rs1(0)
rs1.MoveNext
Wend
End Sub

Public Sub clr()
cmbadno.Text = ""
txtname.Text = ""
txtcourse.Text = ""
txtsemester.Text = ""
txtyr.Text = ""

End Sub




Private Sub GetSem()
Dim mn
mn = DateDiff("m", "06/01/" & txtyr.Text, Now)
If mn > 30 Then
txtsemester.Text = 6
ElseIf mn > 24 Then
txtsemester.Text = 5
ElseIf mn > 18 Then
txtsemester.Text = 4
ElseIf mn > 12 Then
txtsemester.Text = 3
ElseIf mn > 6 Then
txtsemester.Text = 2
ElseIf mn > 0 Then
txtsemester.Text = 1
Else
txtsemester.Text = ""
End If
End Sub


