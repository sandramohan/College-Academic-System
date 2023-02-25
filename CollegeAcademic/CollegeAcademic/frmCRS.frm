VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRS 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11535
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11535
   Begin VB.TextBox txtSubId 
      Height          =   375
      Left            =   6600
      TabIndex        =   13
      Text            =   " "
      Top             =   3960
      Width           =   735
   End
   Begin VB.TextBox txtid 
      Height          =   375
      Left            =   6480
      TabIndex        =   12
      Top             =   2280
      Width           =   735
   End
   Begin MSComctlLib.ListView lstview 
      Height          =   2415
      Left            =   2160
      TabIndex        =   10
      Top             =   4800
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4260
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
         Text            =   "Course Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Semester"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sub Id"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Id"
         Object.Width           =   2540
      EndProperty
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
      Left            =   9240
      TabIndex        =   9
      Top             =   4200
      Width           =   1215
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
      Left            =   9240
      TabIndex        =   8
      Top             =   3480
      Width           =   1215
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
      Left            =   9240
      TabIndex        =   7
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton cmdadd2 
      Caption         =   "ADD"
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
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.ComboBox cmbsub 
      Height          =   315
      Left            =   4440
      TabIndex        =   5
      Top             =   3960
      Width           =   2055
   End
   Begin VB.ComboBox cmbsem 
      Height          =   315
      ItemData        =   "frmCRS.frx":0000
      Left            =   4440
      List            =   "frmCRS.frx":0016
      TabIndex        =   4
      Top             =   3240
      Width           =   2055
   End
   Begin VB.ComboBox cmbcrse 
      Height          =   315
      Left            =   4440
      TabIndex        =   3
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "COURSE-SUBJECT ALLOCATION"
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
      Left            =   960
      TabIndex        =   11
      Top             =   480
      Width           =   10095
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   1920
      TabIndex        =   2
      Top             =   3840
      Width           =   1935
   End
   Begin VB.Label Label2 
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
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Select Course"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   2280
      Width           =   1935
   End
End
Attribute VB_Name = "frmCRS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim id As Integer

Public Sub addsub()
cmbsub.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "subject", con, 3, 3
While Not rs1.EOF
cmbsub.AddItem rs1(1)
rs1.MoveNext
Wend
End Sub
Public Sub addcrs()
cmbcrse.Clear
If rs1.State = 1 Then rs1.Close
rs1.Open "course", con, 3, 3
While Not rs1.EOF
cmbcrse.AddItem rs1(1)
rs1.MoveNext
Wend
End Sub
Private Sub cmbcrse_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from course where name='" & cmbcrse.Text & "'", con, 3, 3
If Not rs.EOF Then
txtid.Text = rs(0)
End If
End Sub

Private Sub cmbsub_Click()
If rs.State = 1 Then rs.Close
rs.Open "select * from subject where subname='" & cmbsub.Text & "'", con, 3, 3
If Not rs.EOF Then
txtSubId.Text = rs(0)
End If

End Sub

Private Sub cmdadd2_Click()

Set lst = lstview.ListItems.Add(, , txtid.Text)
lst.SubItems(1) = cmbsem.Text
lst.SubItems(2) = txtSubId.Text
lst.SubItems(3) = id

id = id + 1
cmbcrse.Locked = False
cmbsem.Locked = False
cmbsub.Locked = False
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = False
End Sub

Private Sub cmdcancel_Click()
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
End Sub

Private Sub cmdnew_Click()
cmdnew.Enabled = False
cmdadd2.Enabled = True
cmdsave.Enabled = True
cmdcancel.Enabled = True
getId
End Sub

Private Sub cmdsave_Click()
For i = 1 To lstview.ListItems.Count
Set lst = lstview.ListItems(i)
cmd.CommandText = "insert into coursecrs values(" & lst.SubItems(3) & "," & lstview.ListItems(i) & "," & lst.SubItems(2) & "," & lst.SubItems(1) & ")"
MsgBox cmd.CommandText
cmd.Execute
Next
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False

End Sub

Private Sub Form_Load()
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
Module1.getcon
addcrs
addsub
End Sub

Public Sub getId()
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(csid),5000)+1 from coursecrs", con, 3, 3
If Not rs.EOF Then
id = rs(0)
End If
End Sub
