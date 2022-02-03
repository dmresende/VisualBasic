VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.MDIForm frmBiblio 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Bibliotecário"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   7185
   Icon            =   "frmBiblio.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      _ExtentX        =   12674
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cadastro de Livros"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cadastro de Usuários"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Empréstimo de Livros"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Devoluções de Livros"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Saís do Sistema"
            ImageIndex      =   6
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   45
      Top             =   870
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":0624
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":5E16
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":66F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":6A0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBiblio.frx":8B44
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuCadastro 
      Caption         =   "&Cadastro"
      Begin VB.Menu mnuCadLivros 
         Caption         =   "&Livros"
      End
      Begin VB.Menu mnuCadUsuarios 
         Caption         =   "&Usuarios"
      End
      Begin VB.Menu mnuCadCategorias 
         Caption         =   "&Categorias"
      End
      Begin VB.Menu mnuCadEditoras 
         Caption         =   "&Editoras"
      End
      Begin VB.Menu mnuSair 
         Caption         =   "&Sair do Sistema"
      End
   End
   Begin VB.Menu mnuOperacoes 
      Caption         =   "&Operações"
      Begin VB.Menu mnuEmprestimos 
         Caption         =   "&Emprestimo de Livros"
      End
      Begin VB.Menu mnuDevolucao 
         Caption         =   "&Devolução de Livros"
      End
      Begin VB.Menu mnuConsultas 
         Caption         =   "&Consulta"
      End
      Begin VB.Menu mnuLivros 
         Caption         =   "&Livros"
         Begin VB.Menu mnuConTodos 
            Caption         =   "&Todos"
         End
         Begin VB.Menu BS2 
            Caption         =   "-"
         End
         Begin VB.Menu mnuLivrosPorAutor 
            Caption         =   "Por &Autor"
         End
         Begin VB.Menu mnuLivrosPorCategoria 
            Caption         =   "Por &Categoria"
         End
         Begin VB.Menu mnuLivroPorEditora 
            Caption         =   "Por &Editora"
         End
         Begin VB.Menu BS3 
            Caption         =   "-"
         End
         Begin VB.Menu mnuConEmprestado 
            Caption         =   "&Emprestado"
         End
         Begin VB.Menu mnuConAtraso 
            Caption         =   "&em Atraso"
         End
      End
      Begin VB.Menu mnuConUsuarios 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnuConCategorias 
         Caption         =   "&Categorias"
      End
      Begin VB.Menu mnuConEditoras 
         Caption         =   "&Editoras"
      End
   End
   Begin VB.Menu mnuRelatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu mnuRelLivros 
         Caption         =   "&Livros"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRelUsuarios 
         Caption         =   "&Usuários"
      End
      Begin VB.Menu mnuRelCategorias 
         Caption         =   "&Categorias"
      End
      Begin VB.Menu mnuRelEditoras 
         Caption         =   "&Editoras"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Ajuda"
      Begin VB.Menu mnuSobre 
         Caption         =   "&Sobre o Bibliotecário"
      End
   End
End
Attribute VB_Name = "frmBiblio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Config sub menu sair
Private Sub mnuSair_Click()
Dim vOkas As Integer
    vOk = MsgBox("Confirma o encerramento do sistema?", vbYesNo + vbQuestion + vbApplicationModal, "Saída")
    If vOk = vbYes Then End
End Sub

'config sub menu sobre
Private Sub mnuSobre_Click()
    frmAbout.Show vbModal
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

      If Button.Index = 5 Then  'na apostila não o códigoe stá em o "Then" e sem o "End If" o que dava erro na execução do código, inserindo eles fuinciona normalmente.
         mnuSair_Click
     ElseIf Button.Index = 1 Then
         frmCadUsuarios.Show
     End If
     
End Sub

'Por não conseguir uma conexão com o banco de dados tive que comentara parte do exercício que pedia para codificar as instrições de conexão

'Private Sub MDIForm_Unload(Cancel As Integer)
'    Set cnnBiblio = Nothing
'End Sub
 

