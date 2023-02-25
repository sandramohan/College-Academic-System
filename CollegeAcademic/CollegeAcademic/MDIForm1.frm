VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7065
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   12345
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuCourseDet 
      Caption         =   "Course Details"
      Begin VB.Menu mnuNewCou 
         Caption         =   "New Course"
      End
      Begin VB.Menu mnuSub 
         Caption         =   "Subject Entry"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStaff 
         Caption         =   "Staff Registration"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnustudent1 
         Caption         =   "Student Registration"
      End
   End
   Begin VB.Menu mnuAlloc 
      Caption         =   "Allocation"
      Begin VB.Menu mnucourse 
         Caption         =   "Course Subject"
      End
      Begin VB.Menu mnuclass 
         Caption         =   "Class Teacher"
      End
   End
   Begin VB.Menu mnustudent 
      Caption         =   "Student information"
      Begin VB.Menu mnuattendence 
         Caption         =   "Attendence"
      End
      Begin VB.Menu mnuinternal 
         Caption         =   "Internal"
      End
      Begin VB.Menu mnusemalloc 
         Caption         =   "Semester Allocation"
      End
      Begin VB.Menu mnumark 
         Caption         =   "Mark Entry"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "View"
      Begin VB.Menu mnuAttnd 
         Caption         =   "Attendance Sheet"
      End
      Begin VB.Menu mnuInt 
         Caption         =   "Internal Marks"
      End
   End
   Begin VB.Menu mnuRep 
      Caption         =   "Reports"
      Begin VB.Menu mnuRptStudList 
         Caption         =   "Student Llist"
      End
      Begin VB.Menu mnurptCourseSyll 
         Caption         =   "Course Syllabus"
      End
   End
   Begin VB.Menu mnuSec 
      Caption         =   "Security"
      Begin VB.Menu mnuSignout 
         Caption         =   "Signout"
      End
      Begin VB.Menu mnuclose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
If Module1.usr = "Admin" Then
mnuCourseDet.Enabled = True
mnuAlloc.Enabled = True
mnustudent.Enabled = False
mnuView.Enabled = False
ElseIf Module1.usr = "Staff" Then
mnuCourseDet.Enabled = False
mnuAlloc.Enabled = False
mnustudent.Enabled = True
mnuView.Enabled = False
ElseIf Module1.usr = "Student" Then
mnuCourseDet.Enabled = False
mnuAlloc.Enabled = False
mnustudent.Enabled = False
mnuView.Enabled = True
End If
End Sub


Private Sub mnuattendence_Click()
frmAttendance.Show
End Sub

Private Sub mnuclass_Click()
frmclsthr.Show
End Sub

Private Sub mnucourse_Click()
frmCRS.Show
End Sub

Private Sub mnuinternal_Click()
frmacad.Show
End Sub

Private Sub mnumark_Click()
frmintmrk.Show
End Sub

Private Sub mnuNewCou_Click()
frmCourse.Show
End Sub

Private Sub mnurptCourseSyll_Click()
rptSyllabus.Show
End Sub

Private Sub mnuRptStudList_Click()
rptStudList.Show
End Sub

Private Sub mnusemalloc_Click()
frmStudSem.Show
End Sub

Private Sub mnuSignout_Click()
Unload Me
frmLogin.Show
End Sub

Private Sub mnuStaff_Click()
frmStaff.Show
End Sub

Private Sub mnustudent1_Click()
frmSTUD.Show
End Sub

Private Sub mnuSub_Click()
frmSubject.Show
End Sub
