VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form7 
   Caption         =   "Form7"
   ClientHeight    =   10185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16275
   LinkTopic       =   "Form7"
   ScaleHeight     =   10185
   ScaleWidth      =   16275
   StartUpPosition =   3  'Windows Default
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
      Height          =   775
      Left            =   8280
      TabIndex        =   31
      Top             =   8520
      Width           =   2175
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "PRINT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   775
      Left            =   5640
      TabIndex        =   30
      Top             =   8520
      Width           =   2175
   End
   Begin VB.TextBox txtunpaidfee 
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
      Left            =   13680
      TabIndex        =   29
      Top             =   7800
      Width           =   2295
   End
   Begin VB.TextBox txtpaidfee 
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
      Left            =   8760
      TabIndex        =   27
      Top             =   7680
      Width           =   2295
   End
   Begin VB.TextBox txttotalfee 
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
      Left            =   4080
      TabIndex        =   25
      Top             =   7800
      Width           =   2655
   End
   Begin VB.TextBox txticard 
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
      Left            =   4080
      TabIndex        =   24
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox txtinternet 
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
      Left            =   4080
      TabIndex        =   23
      Top             =   6600
      Width           =   2655
   End
   Begin VB.TextBox txttution 
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
      Left            =   4080
      TabIndex        =   22
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txtexam 
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
      Left            =   4080
      TabIndex        =   21
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtlab 
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
      Left            =   4080
      TabIndex        =   20
      Top             =   4800
      Width           =   2655
   End
   Begin VB.TextBox txtlibrary 
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
      Left            =   4080
      TabIndex        =   19
      Top             =   4200
      Width           =   2655
   End
   Begin VB.TextBox txtcourse 
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
      Left            =   4080
      TabIndex        =   18
      Top             =   3600
      Width           =   2655
   End
   Begin VB.TextBox txtstudcourse 
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
      Left            =   11640
      TabIndex        =   8
      Top             =   1920
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker txtdate 
      Height          =   495
      Left            =   14040
      TabIndex        =   5
      Top             =   120
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   113442817
      CurrentDate     =   44285
   End
   Begin VB.TextBox txtstudname 
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
      Left            =   4920
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtstudid 
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
      Left            =   8040
      TabIndex        =   1
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Unpaid Fee"
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
      Left            =   11400
      TabIndex        =   28
      Top             =   7800
      Width           =   2175
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Paid Fee"
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
      Left            =   6960
      TabIndex        =   26
      Top             =   7800
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Fee"
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
      Left            =   1080
      TabIndex        =   17
      Top             =   7800
      Width           =   2655
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "I-Card Fee"
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
      Left            =   1080
      TabIndex        =   16
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet Fee"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   6600
      Width           =   2655
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Tution Fee"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   6000
      Width           =   2655
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Exam Fee"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   5400
      Width           =   2655
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Lab Fee"
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
      Left            =   1080
      TabIndex        =   12
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Liabrary Fee"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Fee"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   3600
      Width           =   2655
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "-------------FEES DETAIL-------------"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5160
      TabIndex        =   9
      Top             =   2880
      Width           =   6135
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Course Applied:"
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
      Left            =   8280
      TabIndex        =   7
      Top             =   1920
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "                     **********WOMEN'S COLLEGE KOPARGAON**********"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1200
      TabIndex        =   6
      Top             =   480
      Width           =   14055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Date:"
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
      Left            =   11880
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Student Name:"
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
      Left            =   2040
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Student ID :"
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
      Left            =   5640
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   600
      Picture         =   "Fees Reciept.frx":0000
      Top             =   0
      Width           =   53760
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
Me.Hide
End Sub



Private Sub txtstudid_KeyUp(KeyCode As Integer, Shift As Integer)
R.Close
     S = "Select * from Fees where(S_ID=" & Val(txtstudid.Text) & ")"
     R.Open S, C, adOpenDynamic, adLockOptimistic
     S = "select * from Fees"
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
         
         
        Else
         MsgBox "Record Not Found!", vbInformation, "Fees"
End Sub
