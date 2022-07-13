VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Coneccao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim Con As ADODB.Connection
Dim Buscar As ADODB.Recordset

Public Sub AbrirBD()
Set Con = New ADODB.Connection
Dim ConStr As String

If Con.State = 0 Then
    ConStr = "Provider=SQLNCLI11;Server=CADU-DESK\SQL2017;Database=AgBd;Trusted_Connection=Yes"
    Con.Open = ConStr
End If
End Sub

Public Sub FecharBD()
If Con.State = 1 Then Con.Close
    Set Con = Nothing
End Sub

Public Function ListaUsu() As ADODB.Recordset
AbrirBD
Set Buscar = New ADODB.Recordset
Buscar.CursorLocation = adUseClient
Buscar.Open "Select * from User order by UserNome", Con, adOpenForwardOnly, adLockReadOnly
Buscar.MoveFirst
Set ListaUsu = Buscar
End Function
