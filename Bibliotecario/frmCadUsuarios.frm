VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmCadUsuarios 
   Caption         =   "Cadastro de Usu�rios"
   ClientHeight    =   4005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4005
   ScaleWidth      =   7305
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtCEP 
      Height          =   285
      Left            =   1260
      MaxLength       =   9
      TabIndex        =   7
      Top             =   3675
      Width           =   1470
   End
   Begin VB.TextBox txtTelefone 
      Height          =   285
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   6
      Top             =   3270
      Width           =   2070
   End
   Begin VB.TextBox txtEstado 
      Height          =   285
      Left            =   1230
      MaxLength       =   2
      TabIndex        =   5
      Top             =   2835
      Width           =   675
   End
   Begin VB.TextBox txtCidade 
      Height          =   285
      Left            =   1245
      MaxLength       =   25
      TabIndex        =   4
      Top             =   2355
      Width           =   3345
   End
   Begin VB.TextBox txtEndereco 
      Height          =   285
      Left            =   1230
      MaxLength       =   60
      TabIndex        =   3
      Top             =   1905
      Width           =   6030
   End
   Begin VB.TextBox txtNomeUsuario 
      Height          =   285
      Left            =   1230
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1500
      Width           =   3720
   End
   Begin VB.TextBox txtCodUsuario 
      Height          =   285
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1005
      Width           =   870
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Apagar Registro Atual"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Retorna ao menu principal"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4260
      Top             =   3090
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadUsuarios.frx":57E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadUsuarios.frx":60BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadUsuarios.frx":6F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadUsuarios.frx":8C18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCEP 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CEP:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   645
      TabIndex        =   14
      Top             =   3675
      Width           =   870
   End
   Begin VB.Label lblTelefone 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   255
      TabIndex        =   13
      Top             =   3300
      Width           =   870
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   465
      TabIndex        =   12
      Top             =   2865
      Width           =   615
   End
   Begin VB.Label lblCidade 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   480
      TabIndex        =   11
      Top             =   2370
      Width           =   615
   End
   Begin VB.Label lblEndereco 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endere�o:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   225
      TabIndex        =   10
      Top             =   1890
      Width           =   840
   End
   Begin VB.Label lblNome 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   525
      TabIndex        =   9
      Top             =   1470
      Width           =   540
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C�digo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   435
      TabIndex        =   8
      Top             =   1005
      Width           =   615
   End
End
Attribute VB_Name = "frmCadUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim vInclusao As Boolean

'centraliza o formul�rio na �rea de trabalho MDI
Private Sub Form_Load()
    Me.Left = (frmBiblio.ScaleWidth - Me.Width) / 2
    Me.Top = (frmBiblio.ScaleHeight - Me.Height) / 2
End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys CStr(text), wait
    Set WshShell = Nothing
End Sub

'tecla enter for precionada o foco vai para  proximo controle
Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        Sendkeys "{Tab}"
        KeyAscii = 0
    End If

End Sub


'--------------------------------------------------------------------------------
'FAZER CODIFICA��O DOS BOT�ES
'--------------------------------------------------------------------------------
'Por n�o conseguir uma conex�o com o banco de dados tive que comentara parte do exerc�cio que pedia para codificar as instri��es de conex�o
'Private Sub txtCodUsuario_LostFocus()
'
'    Dim cnnComando As New ADODB.Command
'
'    Dim reSelecao  As New ADODB.Recordset
'
'    On Error GoTo errSelecao
'
'    'verifica a validade
'    If Val(txtCodUsuario.text) = 0 Then
'        MsgBox "N�o foi digitado um c�digo v�lido,verifique.", vbExclamation + vbOKOnly + vbApplicationModal, "Erro"
'
'        Exit Sub
'
'    End If
'
'    Screen.MousePointer = vbHourglass
'
'    With cnnComando
'        .ActiveConnection = cnnBiblio
'        .CommandType = adCmdText
'        'Montando contato com SQL [ SELECT ] para se comunicar com a tabela
'        .CommandText = "SELECT * FROM Usuarios WHERE CodUsuario = " & txtCodUsuario.text & ";"
'        Set rsSelecao = .Execute
'    End With
'
'    With rsSelecao
'
'        If .EOF And .BOF Then
'            'no caso de de recordset vazio '
'            Limpar_Dados
'            'j� come�a ma inclus�o
'            vInclusao = True
'        Else
'            txtNomeUsuario.text = !NomeUsuario
'            txtEndereco.text = !Endereco
'            txtCidade.text = !Cidade
'            txtEstado.text = !Estado
'            txtCEP.text = !CEP
'            txtTelefone.text = Empty & !Telefone
'            'reconhece como altera��o
'            vInclusao = False
'            'habilita bot�o excluir
'            Toobar1.Button(3).Enabled = True
'        End If
'
'    End With
'
'    'desabilita o digita��o do cod
'    txtCodUsuario.Enabled = False
'Saida:
'    'elimina command e o recordset da memoria
'    Set rsSelecao = Nothing
'    Set cnnComando = Nothing
'    Screen.MousePointer = vbDefault
'
'    Exit Sub
'
'    errSelecao
'
'    With Err
'
'        If .Number <> 0 Then
'            MsgBox "Houve um erro na recupera��o do registro solicitado", vbExclamation + vbOKOnly + vbApplicationModal, "Aviso"
'            .Number = 0
'            GoTo Saida
'        End If
'
'    End With
'
'End Sub
'// ESSE PARTE DE CASE DOS BOT�ES N�O VAI DAR CERTO UMA VEZ QUE N�O ESTOU USANDO O BANCO DE DADOS E ELE VAI ENTRAR EM CONTATO COM O DQL
'// ENT�O NESTE CASO VOU FAZER UM PROVIS�RIO PAR QUE O PROGRAMA FUNCIONE SEM GRAVAR NO BANCO DE DADOS
'Private Sub LimparDados()
'        'limpa os dados
'        txtNomeUsuario.text = Empty
'        txtEndereco.text = Empty
'        txtCidade.text = Empty
'        txtEstado.text = Empty
'        txtCEP.text = Empty
'        txtTelefone.text = Empty
'End Sub
'
'Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    Select Case Button.Index
'        Case 1
'            'Bot�o gravar:
'            GravarDados
'        Case 2
'            'Bot�o Limpar
'            LimparTela
'        Case 3
'            'Bot�o Excluir
'            ExcluirRegistro
'        Case 4
'            'Bot�o Retornar
'            Unload Me
'        End Select
'End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then  'na apostila n�o o c�digoe st� em o "Then" e sem o "End If" o que dava erro na execu��o do c�digo, inserindo eles fuinciona normalmente.
        txtCodUsuario.text = Empty
        txtNomeUsuario.text = Empty
        txtEndereco.text = Empty
        txtCidade.text = Empty
        txtEstado.text = Empty
        txtCEP.text = Empty
        txtTelefone.text = Empty
        MsgBox "Grava��o conclu�da com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Gravado"
    ElseIf Button.Index = 2 Then
        txtCodUsuario.text = Empty
        txtNomeUsuario.text = Empty
        txtEndereco.text = Empty
        txtCidade.text = Empty
        txtEstado.text = Empty
        txtCEP.text = Empty
        txtTelefone.text = Empty
        MsgBox "Limpo, com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Limpo"
    ElseIf Button.Index = 3 Then
        vOk = MsgBox("Confirma a exclusao desse regitro?", vbApplicationModal + vbQuestion + vbYesNo, "Exclus�o")
        txtCodUsuario.text = Empty
        txtNomeUsuario.text = Empty
        txtEndereco.text = Empty
        txtCidade.text = Empty
        txtEstado.text = Empty
        txtCEP.text = Empty
        txtTelefone.text = Empty
        MsgBox "Excluido com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Exclu�do"
    ElseIf Button.Index = 4 Then
        Unload Me
        frmBiblio.Show
    End If
    
End Sub
