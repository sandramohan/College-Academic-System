VERSION 5.00
Begin VB.Form frmSubject 
   Caption         =   "Form1"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   12525
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
      Left            =   1320
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
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
      Left            =   4200
      TabIndex        =   6
      Top             =   3960
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
      Left            =   2760
      TabIndex        =   4
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txtsname 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3720
      TabIndex        =   3
      Top             =   2640
      Width           =   2055
   End
   Begin VB.TextBox txtid 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   3720
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "SUBJECT DETAILS"
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
      Left            =   960
      TabIndex        =   5
      Top             =   720
      Width           =   6135
   End
   Begin VB.Label Label2 
      Caption         =   "Subject Name"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Subject ID"
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
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   1815
   End
End
Attribute VB_Name = "frmSubject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim rs As New ADODB.Recordset
Private Sub cmdcancel_Click()
clr
cmdnew.Enabled = True
cmdsave.Enabled = False
cmdcancel.Enabled = False
End Sub

Private Sub cmdnew_Click()
clr
cmdnew.Enabled = False
cmdsave.Enabled = True
cmdcancel.Enabled = True
If rs.State = 1 Then rs.Close
rs.Open "select isnull(max(subid),1000)+1 from subject", con, 3, 3
If Not rs.EOF Then
txtid.Text = rs(0)
End If
End Sub

Private Sub cmdsave_Click()
If txtid.Text = "" Or txtsname.Text = "" Then
MsgBox "fields cannot left blank"
Else
cmd.CommandText = "insert into subject values(" & Val(txtid.Text) & ",'" & txtsname.Text & "')"
cmd.Execute
MsgBox "Record inserted successfully"
End If
End Sub

Public Sub clr()
txtid.Text = ""
txtsname.Text = ""

End Sub

Private Sub Form_Load()
getcon
cmdsave.Enabled = False
cmdcancel.Enabled = False
End Sub
