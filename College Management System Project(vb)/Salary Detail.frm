VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form4"
   ClientHeight    =   9945
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18855
   LinkTopic       =   "Form4"
   ScaleHeight     =   9945
   ScaleWidth      =   18855
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
      Left            =   3960
      TabIndex        =   17
      Top             =   9000
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
      Left            =   7080
      TabIndex        =   16
      Top             =   9000
      Width           =   2175
   End
   Begin MSComCtl2.DTPicker txtsaldate 
      Height          =   495
      Left            =   5760
      TabIndex        =   15
      Top             =   4320
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
      Format          =   113836033
      CurrentDate     =   44288
   End
   Begin VB.ComboBox txtsalpaid 
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
      ItemData        =   "Salary Detail.frx":0000
      Left            =   5760
      List            =   "Salary Detail.frx":000A
      TabIndex        =   14
      Top             =   7200
      Width           =   4695
   End
   Begin VB.TextBox txtamount 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5760
      TabIndex        =   13
      Top             =   6480
      Width           =   4695
   End
   Begin VB.TextBox txtstaffname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5760
      TabIndex        =   12
      Top             =   5760
      Width           =   4695
   End
   Begin VB.TextBox txtstaffid 
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
      TabIndex        =   11
      Top             =   5040
      Width           =   4695
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
      Left            =   10680
      TabIndex        =   10
      Top             =   8160
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
      Left            =   8160
      TabIndex        =   9
      Top             =   8160
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
      Left            =   5640
      TabIndex        =   8
      Top             =   8160
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
      Left            =   3120
      TabIndex        =   7
      Top             =   8160
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
      Left            =   720
      TabIndex        =   6
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary is paid"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
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
      Left            =   2400
      TabIndex        =   4
      Top             =   6480
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Name"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff ID"
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
      Left            =   2400
      TabIndex        =   2
      Top             =   5040
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
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
      Left            =   2400
      TabIndex        =   1
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Salary Detail"
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
      Height          =   735
      Left            =   10440
      TabIndex        =   0
      Top             =   2880
      Width           =   4575
   End
   Begin VB.Image Image2 
      Height          =   4425
      Left            =   0
      Picture         =   "Salary Detail.frx":0017
      Top             =   -240
      Width           =   23100
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   0
      Picture         =   "Salary Detail.frx":210F8
      Top             =   0
      Width           =   53760
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim C As New ADODB.Connection
Dim R As New ADODB.Recordset
Dim S As String

Private Sub cmdadd_Click()

        txtstaffid.Text = " "
        txtstaffname.Text = " "
        txtamount.Text = " "
        txtsalpaid.Text = " "
        txtstaffid.SetFocus
End Sub

Private Sub cmddelete_Click()
confirm = MsgBox("Do you want to delete staff record", vbCritical, "Deletion Confirmation")
R.Close
     S = "delete from Salary where(Staff_ID=" & Val(txtstaffid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Salary"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
        End If
        MsgBox "Salary Deleted Successfully!", vbInformation, "Salary"
End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub


Private Sub cmdnext_Click()
R.MoveNext
        If Not R.EOF Then
        txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
        Else
        MsgBox "No More Records!", vbInformation, "Salary"
        End If
End Sub

Private Sub cmdprevious_Click()
R.MovePrevious
        If Not R.BOF Then
        txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
        Else
        MsgBox "No More Records!", vbInformation, "Salary"
        End If
End Sub

Private Sub cmdsave_Click()
R.Close

     S = "Insert Into Salary Values('" & txtsaldate & "'," & Val(txtstaffid.Text) & ",'" & txtstaffname.Text & "'," & Val(txtamount.Text) & ",'" & txtsalpaid.Text & "')"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Salary"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
       End If
        MsgBox "Salary Saved Successfully!", vbInformation, "Salary"

End Sub

Private Sub cmdshow_Click()
R.Close
     S = "Select * from Staff where(Staff_ID=" & Val(txtstaffid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Staff"
      If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        Else
         MsgBox "Record Not Found!", vbInformation, "Staff"
        End If

End Sub

Private Sub cmdupdate_Click()
    R.Close
    S = "Update Salary set Date='" & txtsaldate & "',Staff_name='" & txtstaffname.Text & "',Amount=" & Val(txtamount.Text) & ",Salary_is_paid='" & txtsalpaid.Text & "' Where Staff_ID=" & Val(txtstaffid.Text) & ""
    R.Open S, C, adOpenDynamic, adLockOptimistic
    S = "select * from Salary"
    R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
        End If
        MsgBox "Salary Updated Successfully!", vbInformation, "Salary"
End Sub

Private Sub Form_Load()
S = "select * from Salary"
C.Open "Provider = Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
R.Open S, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtsaldate = R.Fields(0).Value
        txtstaffid.Text = R.Fields(1).Value
        txtstaffname.Text = R.Fields(2).Value
        txtamount.Text = R.Fields(3).Value
        txtsalpaid.Text = R.Fields(4).Value
        End If
End Sub



Private Sub txtstaffid_KeyUp(KeyCode As Integer, Shift As Integer)
R.Close
     S = "Select * from Staff where(Staff_ID=" & Val(txtstaffid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Staff"
      If Not R.BOF And Not R.EOF Then
        R.MoveFirst
         txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        Else
         MsgBox "Record Not Found!", vbInformation, "Staff"
        End If
End Sub
