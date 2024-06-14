VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form10 
   Caption         =   "Form10"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15585
   LinkTopic       =   "Form10"
   ScaleHeight     =   9540
   ScaleWidth      =   15585
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3855
      Left            =   1320
      TabIndex        =   0
      Top             =   2520
      Width           =   13215
      _ExtentX        =   23310
      _ExtentY        =   6800
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
      Caption         =   "FEES REPORT"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   1920
      Width           =   3375
   End
   Begin VB.Image Image1 
      Height          =   18480
      Left            =   -120
      Picture         =   "Fees Report.frx":0000
      Top             =   -120
      Width           =   53760
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Form_Load()
  con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
  rs.CursorLocation = adUseClient
  rs.Open "Select * from Fees", con, adOpenKeyset, adLockPessimistic, adcmdtxt
  Set DataGrid1.DataSource = rs
  DataGrid1.Refresh
  Set rs = Nothing
End Sub

