VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Form1"
   ClientHeight    =   10500
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16245
   LinkTopic       =   "Form1"
   ScaleHeight     =   10500
   ScaleWidth      =   16245
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9360
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   7200
      MaskColor       =   &H00808080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8400
      Width           =   1935
   End
   Begin VB.CommandButton Logincmd 
      BackColor       =   &H00C0C0C0&
      Caption         =   "LOGIN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8400
      UseMaskColor    =   -1  'True
      Width           =   1935
   End
   Begin VB.TextBox txtpassword 
      Height          =   855
      IMEMode         =   3  'DISABLE
      Left            =   8400
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   7200
      Width           =   2895
   End
   Begin VB.TextBox txtusername 
      Height          =   855
      Left            =   8400
      TabIndex        =   3
      Top             =   6000
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   4920
      TabIndex        =   2
      Top             =   7320
      Width           =   2655
   End
   Begin VB.Label Label2 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   6120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Caption         =   "   LOGIN "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   735
      Left            =   6720
      TabIndex        =   0
      Top             =   5040
      Width           =   2895
   End
   Begin VB.Image Image1 
      Height          =   5040
      Left            =   0
      Picture         =   "Login Form.frx":0000
      Top             =   0
      Width           =   16815
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
txtusername = ""
txtpassword = ""
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Logincmd_Click()
Dim username, password As String
username = "Admin"
password = "sai123"
If txtusername.Text = username And txtpassword.Text = password Then
MsgBox "login successful", vbInformation
MDIForm1.Show
Me.Hide
ElseIf txtpassword.Text <> username Or txtpassword.Text <> password Then
MsgBox "Login Faield", vbCritical
End If
If txtusername.Text = "" Then
MsgBox "Username field can not be left blank", vbInformation
End If
If txtpassword.Text = "" Then
MsgBox "Password field can not be left blank", vbInformation
End If
End Sub
