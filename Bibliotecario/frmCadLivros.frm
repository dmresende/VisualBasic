VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "Mscomctl.ocx"
Begin VB.Form frmCadLivros 
   Caption         =   "Cadastro de Livros"
   ClientHeight    =   5445
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCadLivros.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5445
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraIdioma 
      Caption         =   "Idioma"
      Height          =   1755
      Left            =   7020
      TabIndex        =   18
      Top             =   3315
      Width           =   1335
      Begin VB.OptionButton OptOutro 
         Caption         =   "Outro"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   11
         Top             =   1275
         Width           =   960
      End
      Begin VB.OptionButton OptIngl�s 
         Caption         =   "Ingl�s"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   10
         Top             =   810
         Width           =   1035
      End
      Begin VB.OptionButton OptPortugu�s 
         Caption         =   "Portugu�s"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   330
         Value           =   -1  'True
         Width           =   1050
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   8490
      _ExtentX        =   14975
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
            Object.ToolTipText     =   "Limpar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excluir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Voltar"
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin VB.Frame FraAcompanha 
      Caption         =   "Acompanha:"
      Height          =   1425
      Left            =   6990
      TabIndex        =   17
      Top             =   1275
      Width           =   1380
      Begin VB.CheckBox chkDisquete 
         Caption         =   "Disquete"
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   945
         Width           =   930
      End
      Begin VB.CheckBox chkCD 
         Caption         =   "CD"
         Height          =   210
         Left            =   165
         TabIndex        =   7
         Top             =   300
         Width           =   705
      End
   End
   Begin VB.TextBox txtCodLivro 
      Height          =   285
      Left            =   1350
      MaxLength       =   5
      TabIndex        =   0
      Top             =   945
      Width           =   1125
   End
   Begin VB.TextBox txtTitulo 
      Height          =   285
      Left            =   1335
      MaxLength       =   50
      TabIndex        =   1
      Top             =   1380
      Width           =   5190
   End
   Begin VB.TextBox txtAutor 
      Height          =   285
      Left            =   1335
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1830
      Width           =   5040
   End
   Begin VB.ComboBox cboCategoria 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2880
      Width           =   3870
   End
   Begin VB.ComboBox cboEditora 
      Height          =   315
      Left            =   1335
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2325
      Width           =   3870
   End
   Begin VB.TextBox txtObservacoes 
      Height          =   1935
      Left            =   1305
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   3330
      Width           =   5580
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   510
      Top             =   4530
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
            Picture         =   "frmCadLivros.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadLivros.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadLivros.frx":1FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCadLivros.frx":3D00
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblObserva��es 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Observa��es:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   75
      TabIndex        =   19
      Top             =   3330
      Width           =   1185
   End
   Begin VB.Label lblCodigo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C�digo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   585
      TabIndex        =   16
      Top             =   960
      Width           =   630
   End
   Begin VB.Label lblTitulo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T�tulo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   720
      TabIndex        =   15
      Top             =   1395
      Width           =   630
   End
   Begin VB.Label lblAutor 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Autor:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   750
      TabIndex        =   14
      Top             =   1830
      Width           =   630
   End
   Begin VB.Label lblEditora 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Editora:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   585
      TabIndex        =   13
      Top             =   2340
      Width           =   630
   End
   Begin VB.Label lblCategoria 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categoria:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   390
      TabIndex        =   5
      Top             =   2910
      Width           =   855
   End
End
Attribute VB_Name = "frmCadLivros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Embora configure a ComboBox, n�o tem a lista ainda, mas n�o foi tentado com Banco de dados
'Private Sub cboCategoria_Click()
'    With cboCategoria
'        If .ListIndex <> -1 Then
'            vCodCategoria = ItemData(.ListIndex)
'        Else
'            vCodEditora = 0
'        End If
'    End With
'End Sub
'
''Embora configure a ComboBox, n�o tem a lista ainda, mas n�o foi tentado com Banco de dados
'Private Sub cboEditora_Click()
'    With cboEditora
'        'verifica se foi elecionado item abaixo
'        If .ListIndex <> -1 Then
'            'se foi atribi � vari�vel vCodEditora o conte�do da propriedade ItemData
'            vCodEditora = ItemData(.ListIndex)
'        Else
'            'sen�o, zeta a vari�vel
'            vCodEditora = 0
'        End If
'    End With
'End Sub

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys CStr(text), wait
    Set WshShell = Nothing
End Sub



''tecla enter for precionada o foco vai para  proximo controle
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{Tab}"
        KeyAscii = 0
    End If
End Sub

'Evento Check
'Private Sub chkCD_Click()
'    If cnkAcompCD.Value = vbChecked Then
'        vAcompCD = True
'    Else
'        vAcompCD = False
'    End If
'End Sub

'CONFIGURA��O QUE DEPENDE DA CONEX�O COM BANCO DE DADOS
'Private Sub cboEditora_Change()
'    Dim cnnComando As New ADODB.Command
'    Dim rsTemp As New ADODB.Recordset
'    Dim i As Integer
'    On Error GoTo errComboEditoras
'    'executa consulta  editoras
'    With cnnCommand
'        .ActiveConnection = cnnBiblio
'        .CommandType = adCmdStoredProc
'        .CommandText = "EditorasEmOrdemAlfabetica"
'        Set reTemp = .Execute
'    End With
'    With rsTemp
'        'Verifica se existe alguma editora cadastrada:
'        If Not (.EOF And .BOF) Then
'        'Se existe, posicione no primeiro registro
'        .MoveFirst
'        'inicializa a vari�vel i que ser� usada como �ndice para  a propriedade ItemData
'        i = 0
'        While Not .EOF
'            'Adiciona um item � combo o nome da editora:
'            NomeCombo.AddItem !Descri��o, i
'            'Grava na propriedade itemData desse o c�digoda editora:
'            NomeCombo.ItemData(i) = !C�digo
'            'vai para o pr�ximo registro do rs:
'            .MoveNext
'            'Incrementa i:
'            i = i + 1
'        Wend
'    End With
'    End With
'End Sub
'
'Saida:
'    Set cnnComando = Nothing
'    Set reTemp = Nothing
'    Exit Sub
'errComboEditoras:
'    With Err
'        If .Number <> 0 Then
'            MsgBox "N�o foi poss�vel a leitura da tavelade Editoras:", vbInformation + vbOKOnly + vbApplicationModal, "Erro ao carregar ComboBox"
'        .Number = 0
'        GoTo Saida
'    End If
'End With
'End Sub

        

'inicio do formul�rio, est� dando um problera quando deixo o as linhas em comentadas ativas
'as variaveis  de editora devem dar pau por se comunicar com o banco de dados
Private Sub Form_Load()
'    Me.Left (frmBiblio.ScaleWidth - Me.Width) / 2
'    Me.Top (frmBiblio.Height - Me.Height) / 2
'    vCodEditora = 0
'    vCodCategoria 0
    vAcompCD = False
    vAcompDisquete = False
    vIdioma = 0
    
    'Depende Banco de dados para puxar a lista de categorias e editoras
'    ComboEditoras cboEditoras
'    Combocategorias cbocategorias
'    cboEditora.ListIndex = -1
'    cboCategoria.LisrIndex = -1
    
    
End Sub

Private Sub OptPortugu�s_Click(Index As Integer)
    vIdioma = Index
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    If Button.Index = 1 Then  'na apostila n�o o c�digoe st� em o "Then" e sem o "End If" o que dava erro na execu��o do c�digo, inserindo eles fuinciona normalmente.
        txtCodLivro.text = Empty
        txtTitulo.text = Empty
        txtAutor.text = Empty
        txtObservacoes.text = Empty
        MsgBox "Grava��o conclu�da com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Gravado"
    ElseIf Button.Index = 2 Then
        txtCodLivro.text = Empty
        txtTitulo.text = Empty
        txtAutor.text = Empty
        txtObservacoes.text = Empty
        MsgBox "Limpo, com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Limpo"
    ElseIf Button.Index = 3 Then
        vOk = MsgBox("Confirma a exclusao desse regitro?", vbApplicationModal + vbQuestion + vbYesNo, "Exclus�o")
        txtCodLivro.text = Empty
        txtTitulo.text = Empty
        txtAutor.text = Empty
        txtObservacoes.text = Empty
        MsgBox "Excluido com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Exclu�do"
    ElseIf Button.Index = 4 Then
        Unload Me
        frmBiblio.Show
    End If

End Sub

'Evento para ser testado com banco de dados
'Private Sub txtCodLivro_LostFocus()
'
'    Dim cnnComando As New ADODB.Command
'
'    Dim rsSelecao  As New ADODB.Recordset
'
'    Dim vCod       As Long
'
'    Dim i          As Integer
'
'    On Error GoTo errSelecao
'
'    'Converte o c�digo digitado para pesquisa:
'    vCod = Val(txtCodLivro.text)
'
'    'Se n�o foi digitado um c�digo v�lido, sai da sub:
'    If vCod = 0 Then Exit Sub
'    Screen.MousePointer = vbHourglass
'End Sub
'
''Tenta selecionar o registro na tabela de livros:
'With cnnComando
'    .ActiveConnection = cnnBiblio
'    .CommandType = adCmdText
'    .CommandText = "SELECT * FROM Livros WHERE CodLivro = " & vCod & ";"
'    Set rsSelecao = .Execute
'End With
'
'With rsSelecao
'
'    If .EOF And .BOF Then
'        'Se o recordset est� vazio, n�o encontrou registro com esse c�digo:
'        LimparDados
'        'Identifica a operacao como inclus�o:
'        vInclusao = True
'    Else
'        'Sen�o, atribui aos campos e vari�veis auxiliares os dados do
'        'registro:
'        txtTitulo.text = !Titulo
'        txtAutor.text = !Autor
'        vCodEditora = !CodEditora
'        vCodCategoria = !CodCategoria
'        vAcompCD = !AcompCD
'        vAcompDisquete = !AcompDisquete
'        vIdioma = !Idioma
'        'Como Observacoes n�o � um campo obrigat�rio, devemos impedir a
'        'atribui��o do valor nulo (se houver) � caixa de texto:
'        txtObservacoes = Empty & !Observacoes
'
'        'Exibe os dados das vari�veis nos controles correspondentes:
'        With cboEditora
'            'Elimina a sele��o atual:
'            .ListIndex = -1
'
'            'Como ListCount retorna o n�mero de itens da combo,
'            'ListCount - 1 � igual ao �ndice do �ltimo item. Portanto, o
'            'loop abaixo ser� executado para todos os itens da combo
'            'atrav�s de seu �ndice:
'            For i = 0 To (.ListCount - 1)
'
'                If vCodEditora = .ItemData(i) Then
'                    'Se ItemData for igual ao c�digo atual,
'                    'seleciona o item e sai do loop:
'                    .ListIndex = i
'
'                    Exit For
'
'                End If
'
'            Next i
'
'        End With
'
'        With cboCategoria
'            .ListIndex = -1
'
'            For i = 0 To (.ListCount - 1)
'
'                If vCodCategoria = .ItemData(i) Then
'                    .ListIndex = i
'
'                    Exit For
'
'                End If
'
'            Next i
'
'        End With
'
'        'Se vAcompCD = True, marca chkAcompCD, sen�o desmarca:
'        chkAcompCD.Value = IIf(vAcompCD, vbChecked, vbUnchecked)
'        chkAcompDisquete.Value = IIf(vAcompDisquete, vbChecked, vbUnchecked)
'        'Marca o bot�o de op��o correspondente ao idioma atual:
'        optIdioma(vIdioma).Value = True
'        'Habilita o bot�o Excluir:
'        Toolbar1.Buttons(3).Enabled = True
'        'Identifica a opera��o como Altera��o:
'        vInclusao = False
'    End If
'
'End With
'
''Desabilita a digita��o do c�digo:
'txtCodLivro.Enabled = False
'Saida:
''Elimina o command e o recordset da mem�ria:
'Set rsSelecao = Nothing
'Set cnnComando = Nothing
'Screen.MousePointer = vbDefault
'
'Exit Sub
'
'errSelecao:
'
'With Err
'
'    If .Number <> 0 Then
'        MsgBox "Erro na recupera��o do registro solicitado:", vbExclamation + vbOKOnly + vbApplicationModal, "Aviso"
'        .Number = 0
'        GoTo Saida
'
'    End If
'
'End With
'
'End Sub
'
'Private Sub LimparDados()
'    'Apaga o conte�do dos campos do formul�rio:
'    txtTitulo.text = Empty
'    txtAutor.text = Empty
'    txtObservacoes.text = Empty
'    'Elimina a sele��o das combos:
'    cboEditora.ListIndex = -1
'    cboCategoria.ListIndex = -1
'    'Desmarca as caixas de verifica��o:
'    chkAcompCD.Value = vbUnchecked
'    chkAcompDisquete.Value = vbUnchecked
'    'Marca a op��o Portugu�s em optIdioma:
'    optIdioma(0).Value = True
'    'Reinicializa as vari�veis auxiliares:
'    vCodEditora = 0
'    vCodCategoria = 0
'    vAcompCD = False
'    vAcompDisquete = False
'    vIdioma = 0
'End Sub

'TESTAR COM BANCO DE DADOS GRAVAR DADOS
'Private Sub GravarDados()
'
'    Dim cnnComando As New ADODB.Command
'
'    Dim vSQL       As String
'
'    Dim vCod       As Long
'
'    Dim vConfMsg   As Integer
'
'    Dim vErro      As Boolean
'
'    On Error GoTo errGravacao
'
'    'Converte o c�digo digitado para grava��o:
'    vCod = Val(txtCodLivro.text)
'    'Verifica os dados digitados:
'    vConfMsg = vbExclamation + vbOKOnly + vbApplicationModal
'    vErro = False
'
'    If vCod = 0 Then
'        MsgBox "O campo C�digo n�o foi preenchido.", vConfMsg, "Erro"
'        vErro = True
'    End If
'
'    If txtTitulo.text = Empty Then
'        MsgBox "O campo T�tulo n�o foi preenchido.", vConfMsg, "Erro"
'        vErro = True
'    End If
'
'    If txtAutor.text = Empty Then
'        MsgBox "O campo Autor n�o foi preenchido.", vConfMsg, "Erro"
'        vErro = True
'    End If
'
'    If vCodEditora = 0 Then
'        MsgBox "N�o foi selecionada uma Editora.", vConfMsg, "Erro"
'        vErro = True
'    End If
'
'    If vCodCategoria = 0 Then
'        MsgBox "N�o foi selecionada uma Categoria.", vConfMsg, "Erro"
'        vErro = True
'    End If
'
'    'Se aconteceu um erro de digita��o, sai da sub sem gravar:
'    If vErro Then Exit Sub
'    Screen.MousePointer = vbHourglass
'
'    'Constr�i o comando SQL para grava��o:
'    If vInclusao Then
'        'Se � uma inclus�o:
'        vSQL = "INSERT INTO Livros (CodLivro, Titulo, Autor, CodEditora, " & "CodCategoria, AcompCD, AcompDisquete, Idioma, Observacoes) " & "VALUES (" & vCod & ",'" & txtTitulo.text & "','" & txtAutor.text & "'," & vCodEditora & "," & vCodCategoria & "," & vAcompCD & "," & vAcompDisquete & "," & vIdioma & ",'" & txtObservacoes.text & "');"
'    Else
'        'Sen�o, altera��o:
'        vSQL = "UPDATE Livros SET Titulo = '" & txtTitulo.text & "', Autor = '" & txtAutor.text & "', CodEditora = " & vCodEditora & ", CodCategoria = " & vCodCategoria & ", AcompCD = " & vAcompCD & ", AcompDisquete = " & vAcompDisquete & ", Idioma = " & vIdioma & ", Observacoes = '" & txtObservacoes.text & "' WHERE CodLivro = " & vCod & ";"
'    End If
'
'    'Executa o comando de grava��o:
'    With cnnComando
'        .ActiveConnection = cnnBiblio
'        .CommandType = adCmdText
'        .CommandText = vSQL
'        .Execute
'    End With
'
'    MsgBox "Grava��o conclu�da com sucesso.", vbApplicationModal + vbInformation + vbOKOnly, "Grava��o OK"
'    'Chama a sub que limpa os dados do formul�rio:
'    LimparTela
'Saida:
'    Screen.MousePointer = vbDefault
'    Set cnnComando = Nothing
'
'    Exit Sub
'
'errGravacao:
'
'    With Err
'
'        If .Number <> 0 Then
'            MsgBox "Erro durante a grava��o dos dados no registro." & vbCrLf & "A opera��o n�o foi completada.", vbExclamation + vbOKOnly + vbApplicationModal, "Opera��o cancelada"
'            .Number = 0
'            GoTo Saida
'        End If
'
'    End With
'
'End Sub
'
