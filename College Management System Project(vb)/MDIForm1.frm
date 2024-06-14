VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00C0FFC0&
   Caption         =   "MDIForm1"
   ClientHeight    =   10080
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   22800
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIForm1.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu studentdetail 
      Caption         =   "Student Detail"
   End
   Begin VB.Menu staffdetail 
      Caption         =   "Staff Details"
   End
   Begin VB.Menu salarydetail 
      Caption         =   "Salary Detail"
   End
   Begin VB.Menu feesdetail 
      Caption         =   "Fees Detail"
   End
   Begin VB.Menu coursedetail 
      Caption         =   "Course Detail"
   End
   Begin VB.Menu report 
      Caption         =   "Reports"
      Begin VB.Menu studentreport 
         Caption         =   "Student Report"
      End
      Begin VB.Menu staffreport 
         Caption         =   "Staff Report"
      End
      Begin VB.Menu feesreport 
         Caption         =   "Fees Report"
      End
      Begin VB.Menu coursereport 
         Caption         =   "Course Report"
      End
      Begin VB.Menu salaryreport 
         Caption         =   "Salary Report"
      End
   End
   Begin VB.Menu logout 
      Caption         =   "Logout"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub coursedetail_Click()
Form6.Show
End Sub

Private Sub coursereport_Click()
Form12.Show
End Sub

Private Sub feesdetail_Click()
Form5.Show
End Sub

Private Sub feesreceipt_Click()
Form7.Show
End Sub

Private Sub feesreport_Click()
Form10.Show
End Sub

Private Sub logout_Click()
Unload Me
form1.Show
MsgBox ("logout successful")
End Sub
Private Sub salarydetail_Click()
Form4.Show
End Sub

Private Sub salaryreport_Click()
Form11.Show
End Sub

Private Sub staffdetail_Click()
Form3.Show
End Sub

Private Sub staffreport_Click()
Form9.Show
End Sub

Private Sub studentdetail_Click()
Form2.Show
End Sub

Private Sub studentreport_Click()
Form8.Show
End Sub
