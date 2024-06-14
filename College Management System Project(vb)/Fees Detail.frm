VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form5"
   ClientHeight    =   8685
   ClientLeft      =   195
   ClientTop       =   510
   ClientWidth     =   16095
   LinkTopic       =   "Form5"
   ScaleHeight     =   8685
   ScaleWidth      =   16095
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtunpaidfee 
      Height          =   525
      Left            =   9960
      TabIndex        =   37
      Top             =   6480
      Width           =   3135
   End
   Begin VB.TextBox txtpaidfee 
      Height          =   525
      Left            =   3720
      TabIndex        =   36
      Top             =   6600
      Width           =   3135
   End
   Begin VB.TextBox txttotalfee 
      Height          =   525
      Left            =   9960
      TabIndex        =   35
      Top             =   5640
      Width           =   3135
   End
   Begin VB.TextBox txtexam 
      Height          =   525
      Left            =   3720
      TabIndex        =   34
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox txticard 
      Height          =   525
      Left            =   9960
      TabIndex        =   33
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox txtlab 
      Height          =   525
      Left            =   3720
      TabIndex        =   32
      Top             =   4920
      Width           =   3135
   End
   Begin VB.TextBox txtinternet 
      Height          =   525
      Left            =   9960
      TabIndex        =   31
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txtliabrary 
      Height          =   525
      Left            =   3720
      TabIndex        =   30
      Top             =   4200
      Width           =   3135
   End
   Begin VB.TextBox txttution 
      Height          =   525
      Left            =   9960
      TabIndex        =   29
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox txtcourse 
      Height          =   525
      Left            =   3720
      TabIndex        =   28
      Top             =   3480
      Width           =   3135
   End
   Begin VB.TextBox txtstudcourse 
      Height          =   525
      Left            =   9960
      TabIndex        =   27
      Top             =   2640
      Width           =   3135
   End
   Begin VB.TextBox txtstudname 
      Height          =   525
      Left            =   3720
      TabIndex        =   26
      Top             =   2760
      Width           =   3135
   End
   Begin VB.TextBox txtstudid 
      Height          =   525
      Left            =   7440
      TabIndex        =   25
      Top             =   2040
      Width           =   2295
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "SHOW"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   16680
      TabIndex        =   22
      Top             =   5640
      Width           =   2175
   End
   Begin VB.CommandButton cmdprevious 
      Caption         =   "PREV"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   16680
      TabIndex        =   21
      Top             =   3480
      Width           =   2175
   End
   Begin VB.CommandButton cmdnext 
      Caption         =   "NEXT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   16680
      TabIndex        =   20
      Top             =   2400
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker txtfeedate 
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   113573889
      CurrentDate     =   44285
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   14040
      TabIndex        =   16
      Top             =   6240
      Width           =   2175
   End
   Begin VB.CommandButton cmdcalculate 
      Caption         =   "CAL"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   16680
      TabIndex        =   15
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton cmddelete 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   14040
      TabIndex        =   14
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   14040
      TabIndex        =   13
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton cmdupdate 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   14040
      TabIndex        =   12
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   14040
      TabIndex        =   11
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Course applied"
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
      Left            =   7560
      TabIndex        =   24
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name"
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
      Left            =   1440
      TabIndex        =   23
      Top             =   2760
      Width           =   2295
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Unpaid Fee"
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
      Left            =   7680
      TabIndex        =   18
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Fee"
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
      Left            =   1440
      TabIndex        =   17
      Top             =   6720
      Width           =   3135
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fee"
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
      Left            =   7680
      TabIndex        =   10
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "I-Card Fee"
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
      Left            =   7680
      TabIndex        =   9
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Fee"
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
      Left            =   7560
      TabIndex        =   8
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Tution Fee"
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
      Left            =   7560
      TabIndex        =   7
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Exam Fee"
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
      Left            =   1440
      TabIndex        =   6
      Top             =   5760
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Fee"
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
      Left            =   1440
      TabIndex        =   5
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Liabrary Fee"
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
      Left            =   1440
      TabIndex        =   4
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Fee"
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
      Left            =   1440
      TabIndex        =   3
      Top             =   3480
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID"
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
      Left            =   5520
      TabIndex        =   2
      Top             =   2040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Fees Detail"
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
      Height          =   852
      Left            =   10200
      TabIndex        =   0
      Top             =   720
      Width           =   5532
   End
   Begin VB.Image Image2 
      Height          =   4425
      Left            =   0
      Picture         =   "Fees Detail.frx":0000
      Top             =   -2640
      Width           =   23100
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   0
      Picture         =   "Fees Detail.frx":210E1
      Top             =   -2400
      Width           =   53760
   End
End
Attribute VB_Name = "Form5"
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
        txtstudcourse.Text = " "
        txtcourse.Text = " "
        txtliabrary.Text = " "
        txtlab.Text = " "
        txtexam.Text = " "
        txttution.Text = " "
        txtinternet.Text = " "
        txticard.Text = " "
        txttotalfee.Text = " "
        txtpaidfee.Text = " "
        txtunpaidfee.Text = " "
        txtstudid.SetFocus
         
End Sub

Private Sub cmdcalculate_Click()
txtpaidfee.Text = Val(txtcourse.Text) + Val(txtliabrary.Text) + Val(txtlab.Text) + Val(txtexam.Text) + Val(txttution.Text) + Val(txtinternet.Text) + Val(txticard.Text)
txtunpaidfee.Text = Val(txttotalfee.Text) - Val(txtpaidfee.Text)
End Sub

Private Sub cmddelete_Click()
confirm = MsgBox("Do you want to delete staff record", vbYesNo + vbCritical, "Deletion Confirmation")
R.Close
     S = "delete from Fees where(S_ID=" & Val(txtstudid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Fees "
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
        
    End If
        MsgBox "Fees Deleted Successfully!", vbInformation, "Fees"

End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub

Private Sub cmdnext_Click()
R.MoveNext
      If Not R.EOF Then
         txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
        
    Else
        MsgBox "No More Records!", vbInformation, "Fees"
    End If

End Sub

Private Sub cmdprevious_Click()
R.MovePrevious
      If Not R.BOF Then
        txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
    Else
        MsgBox "No More Records!", vbInformation, "Fees"
    End If


End Sub

Private Sub cmdsave_Click()
R.Close

     S = "Insert Into Fees Values('" & txtfeedate & "' , " & Val(txtstudid.Text) & ", '" & txtstudname & "', '" & txtstudcourse & "'," & Val(txtcourse.Text) & " , " & Val(txtliabrary.Text) & "," & Val(txtlab.Text) & " ," & Val(txtexam) & "," & Val(txttution.Text) & "," & Val(txtinternet.Text) & "," & Val(txticard.Text) & "," & Val(txttotalfee.Text) & "," & Val(txtpaidfee.Text) & "," & Val(txtunpaidfee.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Fees"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
       txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
        
    End If
        MsgBox "Fees Added Successfully!", vbInformation, "Fees"

End Sub



Private Sub cmdupdate_Click()
R.Close
    S = "Update Fees Set Date='" & txtfeedate & "', S_ID='" & txtstudname & "', S_course='" & txtstudcourse & "',Course_fee=" & Val(txtcourse.Text) & ",Liabrary_fee=" & Val(txtliabrary.Text) & ", Lab_fee=" & Val(txtlab.Text) & ",Exam_fee=" & Val(txtexam.Text) & ",Tution_fee=" & Val(txttution.Text) & ",Internet_fee=" & Val(txtinternet.Text) & ",I_card_fee=" & Val(txticard.Text) & ",Total_fee=" & Val(txttotalfee.Text) & ",Paid_fee=" & Val(txtpaidfee.Text) & ",Unpaid_fee=" & Val(txtunpaidfee.Text) & "  Where S_ID=" & Val(txtstudid.Text) & ""
  
    R.Open S, C, adOpenDynamic, adLockOptimistic
    S = "select * from Fees"
    R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
           
            End If
        MsgBox "Fees Updated Successfully!", vbInformation, "Fees"
End Sub

Private Sub Command1_Click()
Form7.Show
End Sub

Private Sub Form_Load()
S = "select * from Fees"
C.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
R.Open S, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtfeedate = R.Fields(0).Value
        txtstudid.Text = R.Fields(1).Value
        txtstudname.Text = R.Fields(2).Value
        txtstudcourse.Text = R.Fields(3).Value
        txtcourse.Text = R.Fields(4).Value
        txtliabrary.Text = R.Fields(5).Value
        txtlab.Text = R.Fields(6).Value
        txtexam.Text = R.Fields(7).Value
        txttution.Text = R.Fields(8).Value
        txtinternet.Text = R.Fields(9).Value
        txticard.Text = R.Fields(10).Value
        txttotalfee.Text = R.Fields(11).Value
        txtpaidfee.Text = R.Fields(12).Value
        txtunpaidfee.Text = R.Fields(13).Value
        End If

End Sub

