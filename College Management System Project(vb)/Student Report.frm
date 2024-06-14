VERSION 5.00
Begin VB.Form Form11 
   Caption         =   "Form11"
   ClientHeight    =   10395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15255
   LinkTopic       =   "Form11"
   ScaleHeight     =   10395
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset


Private Sub Form_Load()
  con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\College Management System Project\DATABASE OF CLG\databclg.mdb;Persist Security Info=False"
  rs.CursorLocation = adUseClient
  rs.Open "Select * from Salary", con, adOpenKeyset, adLockPessimistic, adcmdtxt
  Set DataGrid1.DataSource = rs
  DataGrid1.Refresh
  Set rs = Nothing
End Sub
End Sub
