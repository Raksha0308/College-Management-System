VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form12 
   Caption         =   "Form12"
   ClientHeight    =   10365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   LinkTopic       =   "Form12"
   ScaleHeight     =   10365
   ScaleWidth      =   14685
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3495
      Left            =   2520
      TabIndex        =   0
      Top             =   3120
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6165
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "COURSE REPORT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   2520
      Width           =   3855
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   0
      Picture         =   "Course Report.frx":0000
      Top             =   0
      Width           =   53760
   End
End
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Form_Load()

  con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
  rs.CursorLocation = adUseClient
  rs.Open "Select * from Course", con, adOpenKeyset, adLockPessimistic, adcmdtxt
  Set DataGrid1.DataSource = rs
  DataGrid1.Refresh
  Set rs = Nothing
End Sub
