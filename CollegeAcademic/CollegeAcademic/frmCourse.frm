VERSION 5.00
Begin VB.Form frmCourse 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12075
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   12075
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   6120
      TabIndex        =   5
      Top             =   2040
      Width           =   3375
   End
   Begin VB.TextBox txtname 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   6120
      TabIndex        =   4
      Top             =   3360
      Width           =   3375
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
      Left            =   3240
      TabIndex        =   3
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
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
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
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
      Height          =   495
      Left            =   7560
      TabIndex        =   1
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox txtdur 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   6120
      TabIndex        =   0
      Top             =   4560
      Width           =   3375
   End
   Begin VB.Label Label4 
      Caption         =   "COURSE REGISTRATION"
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
      Left            =   2160
      TabIndex        =   9
      Top             =   240
      Width           =   7815
   End
   Begin VB.Label Label1 
      Caption         =   "Course ID"
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
      Left            =   600
      TabIndex        =   8
      Top             =   2040
      Width           =   3255
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
      Height          =   735
      Left            =   600
      TabIndex        =   7
      Top             =   3360
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "Duration"
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
      Left            =   600
      TabIndex        =   6
      Top             =   4560
      Width           =   3255
   End
End
Attribute VB_Name = "frmCourse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Public Sub clr()
txtid.Text = ""
txtname.Text = ""
txtdur.Text = ""
End Sub

Private Sub cmdcancel_Click()
clr
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
End Sub

Private Sub cmdnew_Click()
clr
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(courseid),1000)+1 from course", con, 3, 3
If Not rs.EOF Then
txtid.Text = rs(0)
End If
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
End Sub

Private Sub cmdsave_Click()
If txtid.Text = "" Or txtname.Text = "" Or txtdur.Text = "" Then
MsgBox "fields cannot left blank"
Else
cmd.CommandText = "insert into course values(" & Val(txtid.Text) & ",'" & txtname.Text & "'," & txtdur.Text & ")"
cmd.Execute
MsgBox "Record inserted successfully"
End If
End Sub

Private Sub Form_Load()
getcon
End Sub
