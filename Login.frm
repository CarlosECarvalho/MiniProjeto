VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   1605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1605
   ScaleWidth      =   3915
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   6
      ToolTipText     =   "Insira a senha"
      Top             =   1080
      Width           =   2775
   End
   Begin VB.ComboBox ComboUser 
      Height          =   315
      Left            =   120
      Style           =   1  'Simple Combo
      TabIndex        =   4
      ToolTipText     =   "Informe o usuário"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton ButCad 
      Caption         =   "Cadastrar"
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton ButEntrar 
      Caption         =   "Entrar"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton ButSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LabSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   2535
   End
   Begin VB.Label LabUser 
      BackStyle       =   0  'Transparent
      Caption         =   "Usuário"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ComboUser_Change()
Dim ListarUsu As New Coneccao
Dim str As String
Dim valor As String
Dim usu As String
valor = ComboUser.Text
str = "UserName Like " & valor
usu = ListarUsu.ListaUsu.Find ("'User"

While Not ListarUsu.ListaUsu.EOF
    Set ComboUser.DataSource = ListarUsu.ListaUsu
    ComboUser.AddItem (usu)
End While
End Sub

Private Sub ButEntrar_Click()
CadEmp.Show
Unload Me
End Sub

Private Sub ButSair_Click()

Unload Me
End Sub

