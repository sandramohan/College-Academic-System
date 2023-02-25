VERSION 5.00
Begin VB.Form frmStaffSub 
   Caption         =   "Form1"
   ClientHeight    =   7875
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   7875
   ScaleWidth      =   12855
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   8
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton cmdnew 
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7560
      TabIndex        =   6
      Top             =   3240
      Width           =   1215
   End
   Begin VB.ComboBox cmbsub 
      Height          =   315
      Left            =   5400
      TabIndex        =   5
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox cmbsem 
      Height          =   315
      Left            =   5400
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.ComboBox cmbstaff 
      Height          =   315
      Left            =   5400
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "STAFF-SUBJECT REGISTRATION"
      BeginProperty Font 
         Name            =   "Lucida Handwriting"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1800
      TabIndex        =   9
      Top             =   720
      Width           =   9615
   End
   Begin VB.Label Label3 
      Caption         =   "Select Semester"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   2
      Top             =   3960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Select Subject"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   1
      Top             =   4920
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Select Staff"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
End
Attribute VB_Name = "frmStaffSub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdnew_Click()
clr
End Sub

Private Sub cmdsave_Click()
If cmbstaff.Text = "" Or cmbsem.Text = "" Or cmbsub.Text = "" Then
MsgBox "Fields cannot left blank"
Else
cmd.CommandText= insert into
End Sub

