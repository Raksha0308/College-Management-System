VERSION 5.00
Begin VB.Form Form6 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form6"
   ClientHeight    =   9690
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16245
   LinkTopic       =   "Form6"
   ScaleHeight     =   9690
   ScaleWidth      =   16245
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprevious 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   15
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      TabIndex        =   14
      Top             =   8280
      Width           =   2415
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10440
      TabIndex        =   13
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7920
      TabIndex        =   12
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   11
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   10
      Top             =   7440
      Width           =   2175
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   9
      Top             =   7440
      Width           =   2175
   End
   Begin VB.ComboBox txtduration 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "Course Detail.frx":0000
      Left            =   5760
      List            =   "Course Detail.frx":0013
      TabIndex        =   8
      Top             =   5640
      Width           =   4695
   End
   Begin VB.ComboBox txtcoursename 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      ItemData        =   "Course Detail.frx":003F
      Left            =   5760
      List            =   "Course Detail.frx":004F
      TabIndex        =   7
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox txtcoursefees 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5760
      TabIndex        =   6
      Top             =   6360
      Width           =   4695
   End
   Begin VB.TextBox txtcourseid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   5760
      TabIndex        =   5
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Fees"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   6360
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Duration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Course ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Information"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   9600
      TabIndex        =   0
      Top             =   3000
      Width           =   6375
   End
   Begin VB.Image Image2 
      Height          =   4425
      Left            =   0
      Picture         =   "Course Detail.frx":006E
      Top             =   -120
      Width           =   23100
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   0
      OLEDropMode     =   1  'Manual
      Picture         =   "Course Detail.frx":2114F
      Top             =   0
      Width           =   53760
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New ADODB.Connection
Dim R As New ADODB.Recordset
Dim S As String

Private Sub cmdadd_Click()
        txtcourseid.Text = " "
        txtcoursename.Text = " "
        txtduration.Text = " "
        txtcoursefees.Text = " "
        
        txtcourseid.SetFocus

End Sub

Private Sub cmddelete_Click()
confirm = MsgBox("Do you want to delete staff record", vbYesNo + vbCritical, "Deletion Confirmation")
R.Close
     S = "delete from Course where(Course_ID=" & Val(txtcourseid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Course"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
      
    End If
        MsgBox "Course Deleted Successfully!", vbInformation, "Course"
End Sub



Private Sub cmdexit_Click()
Me.Hide
End Sub

Private Sub cmdnext_Click()
R.MoveNext
        If Not R.EOF Then
        txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
      Else
        MsgBox "No More Records!", vbInformation, "Course"
        End If
End Sub

Private Sub cmdprevious_Click()
R.MovePrevious
        If Not R.BOF Then
        txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
      Else
        MsgBox "No More Records!", vbInformation, "Course"
       End If
End Sub

Private Sub cmdsave_Click()
R.Close

     S = "Insert Into Course Values(" & Val(txtcourseid.Text) & " , '" & txtcoursename.Text & "','" & txtduration.Text & "'," & Val(txtcoursefees.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Course"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
      Else
        MsgBox "Course Save Successfully!", vbInformation, "Course"
  End If
End Sub

Private Sub cmdupdate_Click()
 R.Close
    S = "Update Course set Course_name='" & txtcoursename.Text & "',Course_duration='" & txtduration.Text & "',Course_fees=" & Val(txtcoursefees.Text) & " Where Course_ID=" & Val(txtcourseid.Text) & ""
    R.Open S, C, adOpenDynamic, adLockOptimistic
    S = "select * from Course"
    R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
       
    End If
        MsgBox "Student Updated Successfully!", vbInformation, "Course"

End Sub

Private Sub Form_Load()
S = "select * from Course"
C.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
R.Open S, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtcourseid.Text = R.Fields(0).Value
        txtcoursename.Text = R.Fields(1).Value
        txtduration.Text = R.Fields(2).Value
        txtcoursefees.Text = R.Fields(3).Value
       
        End If
End Sub

