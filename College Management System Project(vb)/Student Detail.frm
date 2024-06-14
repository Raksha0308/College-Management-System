VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Form2"
   ClientHeight    =   12375
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19350
   LinkTopic       =   "Form2"
   ScaleHeight     =   12375
   ScaleWidth      =   19350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprevious 
      Caption         =   "PREVIOUS"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14640
      TabIndex        =   25
      Top             =   6120
      Width           =   2535
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   14640
      TabIndex        =   24
      Top             =   7320
      Width           =   2535
   End
   Begin VB.ComboBox txtstudcourse 
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
      ItemData        =   "Student Detail.frx":0000
      Left            =   6240
      List            =   "Student Detail.frx":0010
      TabIndex        =   23
      Top             =   9240
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker studDOB 
      Height          =   495
      Left            =   6240
      TabIndex        =   22
      Top             =   7440
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   114622465
      CurrentDate     =   44285
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   21
      Top             =   8760
      Width           =   2175
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   20
      Top             =   7680
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   19
      Top             =   6720
      Width           =   2175
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   18
      Top             =   5760
      Width           =   2175
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   17
      Top             =   4680
      Width           =   2175
   End
   Begin VB.ComboBox txtstudgender 
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
      ItemData        =   "Student Detail.frx":002F
      Left            =   6240
      List            =   "Student Detail.frx":0039
      TabIndex        =   16
      Top             =   8040
      Width           =   4695
   End
   Begin VB.TextBox txtstudcontact 
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
      Left            =   6240
      TabIndex        =   15
      Top             =   8640
      Width           =   4695
   End
   Begin VB.TextBox txtaddr 
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
      Left            =   6240
      TabIndex        =   14
      Top             =   6840
      Width           =   4695
   End
   Begin VB.TextBox txtmname 
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
      Left            =   6240
      TabIndex        =   13
      Top             =   6240
      Width           =   4695
   End
   Begin VB.TextBox txtfname 
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
      Left            =   6240
      TabIndex        =   12
      Top             =   5640
      Width           =   4695
   End
   Begin VB.TextBox txtstudname 
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
      Left            =   6240
      TabIndex        =   11
      Top             =   5040
      Width           =   4695
   End
   Begin VB.TextBox txtstudid 
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
      Left            =   6240
      TabIndex        =   10
      Top             =   4440
      Width           =   4695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Applied"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   9240
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   8
      Top             =   8640
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   8040
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   5
      Top             =   6840
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   4440
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Information"
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
      Height          =   855
      Left            =   10080
      TabIndex        =   0
      Top             =   3240
      Width           =   6615
   End
   Begin VB.Image Image2 
      Height          =   4425
      Left            =   0
      Picture         =   "Student Detail.frx":004B
      Top             =   0
      Width           =   23100
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   -720
      Picture         =   "Student Detail.frx":2112C
      Top             =   480
      Width           =   53760
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim C As New ADODB.Connection
Dim R As New ADODB.Recordset
Dim S As String

Private Sub cmdadd_Click()
        txtstudid.Text = " "
        txtstudname.Text = " "
        txtfname.Text = " "
        txtmname.Text = " "
        txtaddr.Text = " "
        txtstudgender.Text = " "
        txtstudcontact.Text = " "
        txtstudcourse.Text = " "
        
        txtstudid.SetFocus

End Sub

Private Sub cmddelete_Click()
confirm = MsgBox("Do you want to delete staff record", vbYesNo + vbCritical, "Deletion Confirmation")
R.Close
     S = "delete from Student where(S_ID=" & Val(txtstudid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Student"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
    End If
        MsgBox "Student Deleted Successfully!", vbInformation, "Student"


End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub


Private Sub cmdnext_Click()
R.MoveNext
        If Not R.EOF Then
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
       Else
        MsgBox "No More Records!", vbInformation, "Student"
        End If
End Sub

Private Sub cmdprevious_Click()
R.MovePrevious
        If Not R.BOF Then
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
        Else
        MsgBox "No More Records!", vbInformation, "Student"
       End If
End Sub
        

Private Sub cmdsave_Click()
R.Close

     S = "Insert Into Student Values(" & Val(txtstudid.Text) & " , '" & txtstudname.Text & "','" & txtfname.Text & "' , '" & txtmname.Text & "','" & txtaddr.Text & "' ,'" & studDOB & "','" & txtstudgender & "','" & txtstudcontact.Text & "','" & txtstudcourse.Text & "')"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Student"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
    End If
        MsgBox "Student Added Successfully!", vbInformation, "Student"

End Sub

Private Sub cmdupdate_Click()
    R.Close
    S = "Update Student set S_name='" & txtstudname.Text & "',S_f_name='" & txtfname.Text & "',S_m_name='" & txtmname.Text & "',S_addr='" & txtaddr.Text & "',S_DOB='" & studDOB & "',S_gender='" & txtstudgender.Text & "',S_contact='" & txtstudcontact.Text & "',S_course='" & txtstudcourse.Text & "' Where S_ID=" & Val(txtstudid.Text) & ""
    R.Open S, C, adOpenDynamic, adLockOptimistic
    S = "select * from Student"
    R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
    End If
        MsgBox "Student Updated Successfully!", vbInformation, "Student"

End Sub

Private Sub Form_Load()
  S = "select * from Student"
    C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
    R.Open S, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
    End If

End Sub
Private Sub txtstudid_KeyUp(KeyCode As Integer, Shift As Integer)
R.Close
     S = "Select * from Student where(S_ID=" & Val(txtstudid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Student"
      If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstudid.Text = R.Fields(0).Value
        txtstudname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtaddr.Text = R.Fields(4).Value
        studDOB = R.Fields(5).Value
        txtstudgender.Text = R.Fields(6).Value
        txtstudcontact.Text = R.Fields(7).Value
        txtstudcourse.Text = R.Fields(8).Value
        MsgBox "Student Find Successfully!", vbInformation, "Student"
        Else
        MsgBox "No Record Found!", vbInformation, "Student"
        End If

End Sub
