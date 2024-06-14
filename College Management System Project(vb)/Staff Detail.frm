VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   10680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17250
   LinkTopic       =   "Form3"
   ScaleHeight     =   10680
   ScaleWidth      =   17250
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
      Left            =   13920
      TabIndex        =   27
      Top             =   6120
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
      Left            =   13920
      TabIndex        =   26
      Top             =   7320
      Width           =   2535
   End
   Begin VB.ComboBox txtstaffcourse 
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
      ItemData        =   "Staff Detail.frx":0000
      Left            =   6000
      List            =   "Staff Detail.frx":0010
      TabIndex        =   25
      Top             =   9600
      Width           =   4695
   End
   Begin MSComCtl2.DTPicker staffDOB 
      Height          =   495
      Left            =   6000
      TabIndex        =   24
      Top             =   7200
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
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11640
      TabIndex        =   23
      Top             =   8880
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
      Left            =   11640
      TabIndex        =   22
      Top             =   7800
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
      Left            =   11640
      TabIndex        =   21
      Top             =   6720
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
      Left            =   11640
      TabIndex        =   20
      Top             =   5520
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
      Left            =   11640
      TabIndex        =   19
      Top             =   4440
      Width           =   2175
   End
   Begin VB.ComboBox txtstaffgender 
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
      ItemData        =   "Staff Detail.frx":002F
      Left            =   6000
      List            =   "Staff Detail.frx":0039
      TabIndex        =   18
      Top             =   7800
      Width           =   4695
   End
   Begin VB.TextBox txtqualification 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6000
      TabIndex        =   17
      Top             =   9000
      Width           =   4695
   End
   Begin VB.TextBox txtstaffcontact 
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
      Left            =   6000
      TabIndex        =   16
      Top             =   8400
      Width           =   4695
   End
   Begin VB.TextBox txtstaffaddr 
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
      Left            =   6000
      TabIndex        =   15
      Top             =   6600
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
      Left            =   6000
      TabIndex        =   14
      Top             =   6000
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
      Left            =   6000
      TabIndex        =   13
      Top             =   5400
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
      Height          =   525
      Left            =   6000
      TabIndex        =   12
      Top             =   4800
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
      Left            =   6000
      TabIndex        =   11
      Top             =   4200
      Width           =   4695
   End
   Begin VB.Label Label11 
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
      Left            =   2040
      TabIndex        =   10
      Top             =   9600
      Width           =   3135
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Qualification"
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
      Left            =   2040
      TabIndex        =   9
      Top             =   9000
      Width           =   3135
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Contact No"
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
      Left            =   2040
      TabIndex        =   8
      Top             =   8400
      Width           =   3135
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      Left            =   2040
      TabIndex        =   7
      Top             =   7800
      Width           =   3135
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      Left            =   2040
      TabIndex        =   6
      Top             =   7200
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   2040
      TabIndex        =   5
      Top             =   6600
      Width           =   3135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Mother Name"
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
      Left            =   2040
      TabIndex        =   4
      Top             =   6000
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Father Name"
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
      Left            =   2040
      TabIndex        =   3
      Top             =   5400
      Width           =   3135
   End
   Begin VB.Label Label3 
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
      Left            =   2040
      TabIndex        =   2
      Top             =   4800
      Width           =   3135
   End
   Begin VB.Label Label2 
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
      Left            =   2040
      TabIndex        =   1
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Staff Information"
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
      Height          =   1215
      Left            =   10080
      TabIndex        =   0
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Image Image2 
      Height          =   4425
      Left            =   0
      Picture         =   "Staff Detail.frx":004B
      Top             =   -240
      Width           =   23100
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   0
      Picture         =   "Staff Detail.frx":2112C
      Top             =   600
      Width           =   53760
   End
End
Attribute VB_Name = "Form3"
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
        txtfname.Text = "  "
        txtmname.Text = " "
        txtstaffaddr.Text = " "
        txtstaffgender.Text = " "
        txtstaffcontact.Text = " "
        txtqualification.Text = " "
        txtstaffcourse.Text = " "
        
        txtstaffid.SetFocus
End Sub

Private Sub cmddelete_Click()
confirm = MsgBox("Do you want to delete staff record", vbYesNo + vbCritical, "Deletion Confirmation")
R.Close
     S = "delete from Staff where(Staff_ID=" & Val(txtstaffid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Staff"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
    End If
        MsgBox "Staff Deleted Successfully!", vbInformation, "Staff"

End Sub

Private Sub cmdexit_Click()
Me.Hide
End Sub



Private Sub cmdnext_Click()
R.MoveNext
        If Not R.EOF Then
         txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
        Else
        MsgBox "No More Records!", vbInformation, "Staff"
        End If
End Sub

Private Sub cmdprevious_Click()
R.MovePrevious
        If Not R.BOF Then
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
         Else
        MsgBox "No More Records!", vbInformation, "Staff"
        End If

End Sub

Private Sub cmdsave_Click()
R.Close

     S = "Insert Into Staff Values(" & Val(txtstaffid.Text) & " , '" & txtstaffname.Text & "','" & txtfname.Text & "' , '" & txtmname.Text & "','" & txtstaffaddr.Text & "' ,'" & staffDOB & "','" & txtstaffgender & "','" & txtstaffcontact.Text & "','" & txtqualification.Text & "','" & txtstaffcourse.Text & "')"
     R.Open S, C, adOpenDynamic, adLockOptimistic
  
     S = "select * from Staff"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
    End If
        MsgBox "Staff Added Successfully!", vbInformation, "Staff"

End Sub

Private Sub cmdupdate_Click()
R.Close
    S = "Update Staff set Staff_name='" & txtstaffname.Text & "',Staff_f_name='" & txtfname.Text & "',Staff_m_name='" & txtmname.Text & "',Staff_addr='" & txtstaffaddr.Text & "',Staff_DOB='" & staffDOB & "',Staff_gender='" & txtstaffgender & "',Staff_contact='" & txtstaffcontact.Text & "',Staff_qualification='" & txtqualification.Text & "',Staff_course='" & txtstaffcourse.Text & "'  Where Staff_ID=" & Val(txtstaffid.Text) & ""
    R.Open S, C, adOpenDynamic, adLockOptimistic
    S = "select * from Staff"
    R.Open S, C, adOpenDynamic, adLockOptimistic
     If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
    End If
        MsgBox "Student Updated Successfully!", vbInformation, "Staff"

End Sub



Private Sub Form_Load()
 S = "select * from Staff"
    C.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
    R.Open S, C, adOpenDynamic, adLockOptimistic
        If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
    End If

End Sub
Private Sub txtstaffid_KeyUp(KeyCode As Integer, Shift As Integer)
R.Close
     S = "Select * from Staff where(Staff_ID=" & Val(txtstaffid.Text) & ")"  'or Staff_name='" & txtstaffname.Text & "' or Staff_addr= '" & txtstaffaddr.Text & "')"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Staff"
      If Not R.BOF And Not R.EOF Then
        R.MoveFirst
        txtstaffid.Text = R.Fields(0).Value
        txtstaffname.Text = R.Fields(1).Value
        txtfname.Text = R.Fields(2).Value
        txtmname.Text = R.Fields(3).Value
        txtstaffaddr.Text = R.Fields(4).Value
        staffDOB = R.Fields(5).Value
        txtstaffgender.Text = R.Fields(6).Value
        txtstaffcontact.Text = R.Fields(7).Value
        txtqualification.Text = R.Fields(8).Value
        txtstaffcourse.Text = R.Fields(9).Value
        MsgBox "Staff Find Successfully!", vbInformation, "Staff"
        Else
        MsgBox "No record Found!", vbInformation, "Staff"
        End If

End Sub
