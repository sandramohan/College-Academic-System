VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Form1"
   ClientHeight    =   7605
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   ScaleHeight     =   7605
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox TXTPWD 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3840
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   3120
      Width           =   2535
   End
   Begin VB.TextBox TXTUNAME 
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Cmdok 
      Appearance      =   0  'Flat
      Caption         =   "LOGIN"
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
      Left            =   3000
      TabIndex        =   0
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "LOGIN PAGE"
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
      Left            =   1680
      TabIndex        =   5
      Top             =   720
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
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
      Left            =   1080
      TabIndex        =   4
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
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
      Left            =   1080
      TabIndex        =   3
      Top             =   3000
      Width           =   2175
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Cmdok_Click()
If rs.State = 1 Then rs.Close
rs.Open "SELECT * FROM LOGIN where username='" & TXTUNAME.Text & "' and password='" & txtPwd.Text & "'", con, 3, 3
If Not rs.EOF Then
Module1.usr = rs(2)
Module1.id = rs(3)
Unload Me
MDIForm1.Show
Else
MsgBox "Invalid User!..."
TXTUNAME.Text = ""
txtPwd.Text = ""
TXTUNAME.SetFocus
End If
End Sub

Private Sub Form_Load()
Module1.getcon
End Sub
