VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form CadEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Empresas"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5250
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   5250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton ButDel 
      Caption         =   "Apagar"
      Height          =   495
      Left            =   4200
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton ButAlt 
      Caption         =   "Atualizar"
      Height          =   495
      Left            =   2880
      TabIndex        =   13
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton ButBuscar 
      Caption         =   "Buscar"
      Height          =   495
      Left            =   1440
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton ButNovo 
      Caption         =   "Novo"
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton ButLimpar 
      Caption         =   "Limpar"
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   1080
      Width           =   855
   End
   Begin VB.CommandButton ButLogoff 
      Caption         =   "Logoff"
      Height          =   375
      Left            =   4320
      TabIndex        =   9
      Top             =   600
      Width           =   855
   End
   Begin VB.TextBox TxtMail 
      Height          =   285
      Left            =   600
      TabIndex        =   8
      Top             =   1440
      Width           =   3495
   End
   Begin MSMask.MaskEdBox MskTel 
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   960
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   14
      Mask            =   "(##)#####-####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   600
      TabIndex        =   6
      Top             =   480
      Width           =   3495
   End
   Begin MSMask.MaskEdBox MskCNPJ 
      Height          =   255
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      MaxLength       =   15
      Mask            =   "###.###.####-##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton ButSair 
      Caption         =   "Sair"
      Height          =   375
      Left            =   4320
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Label LabEmail 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label LabTel 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel."
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label LabNome 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.Label LabCNPJ 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "CadEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButLogoff_Click()
Login.Show
Unload Me
End Sub

Private Sub ButSair_Click()
Unload Me
End Sub
