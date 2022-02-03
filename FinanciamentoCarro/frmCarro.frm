VERSION 5.00
Begin VB.Form frmCarro 
   Caption         =   "Simule seu financiamento"
   ClientHeight    =   5475
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7275
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCarro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "&Limpar"
      Enabled         =   0   'False
      Height          =   585
      Left            =   4800
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1335
   End
   Begin VB.CommandButton cmdCalcular 
      Caption         =   "&Calcular"
      Enabled         =   0   'False
      Height          =   600
      Left            =   1320
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1215
   End
   Begin VB.TextBox txtMeses 
      Height          =   285
      Left            =   2895
      TabIndex        =   3
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtJuros 
      Height          =   285
      Left            =   2880
      TabIndex        =   4
      Top             =   1530
      Width           =   735
   End
   Begin VB.TextBox txtEntrada 
      Height          =   285
      Left            =   2880
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   2880
      TabIndex        =   1
      Top             =   180
      Width           =   1575
   End
   Begin VB.PictureBox imgCarro 
      Height          =   1695
      Left            =   0
      Picture         =   "frmCarro.frx":030A
      ScaleHeight     =   1635
      ScaleWidth      =   7155
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2520
      Width           =   7215
   End
   Begin VB.Label lblValFinanciado 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5100
      TabIndex        =   5
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label lblPrestMensal 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   5115
      TabIndex        =   6
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Prestação Mensal R$:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   14
      Top             =   1320
      Width           =   2040
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Financiado R$:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   13
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Meses:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1530
      TabIndex        =   12
      Top             =   1005
      Width           =   1215
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Juros / Mês (%):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1260
      TabIndex        =   11
      Top             =   1530
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Entrada R$:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   1680
      TabIndex        =   10
      Top             =   585
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valor do Financiamento R$:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   2655
   End
End
Attribute VB_Name = "frmCarro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'INICIO'
'declaração de variáveis em escopo global'
Dim vValor As Currency  'declaração da variável identificando seu tipo "Currency"'
Dim vEntrada As Currency 'declaração de variável do tipo Currency"
Dim vJuros As Single 'Declaração var tipo Single1
Dim vMeses As Integer 'tipo Integer'

'Rotina de prestação, onde conterá os calculos e a ativação dos botões _Calcular_ e _Fechar_'
'não consigo saber se a conta está correta
Public Sub Prestação()

    Dim vValFinanciado As Currency

    Dim vPrestMensal   As Currency

    If vValor <> 0 And vJuros <> 0 And vMeses <> 0 Then
        If vEntrada >= vValor Then 'corrigido erro que não deixava o operador logico seguir, pois em outro campo ele estava recebendo o sinal de = para o vValor
            MsgBox "O Valor da Entrada deve ser menor que o do financiamento", vbExclamation + vbApplicationModal, "Aviso"
            lblValFinanciado.Caption = Empty
            lblPrestMensal.Caption = Empty
            txtEntrada.SetFocus

            Exit Sub

        End If

        vValFinanciado = vValor - vEntrada
       
        vPrestMensal = vValFinanciado * vJuros * (1 + vJuros) ^ vMeses / ((1 + vJuros) ^ vMeses - 1) '[forma proposta pelo exercicio]
        'vJuros * vValFinanciado + (vValFinanciado / vMeses) [uma forma de calcular]
        
        lblValFinanciado.Caption = Format(vValFinanciado, "###,##0.00")
        lblPrestMensal.Caption = Format(vPrestMensal, "###,##0.00")
        cmdCalcular.Enabled = True
        cmdLimpar.Enabled = True
    Else
        lblValFinanciado.Caption = Empty
        lblPrestMensal.Caption = Empty
        
    End If

End Sub


'dica do stackoverflow, para afuncionar no Win10'
Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
                   Set WshShell = Nothing
End Sub

'chama prestação
Private Sub cmdCalcular_Click()
    Prestação
End Sub

'ao invez de "fechar" mudei obotão para limpar, uma vez que temos o "x" no canto superior direito, e se quisermos refazer uma pesquiza bate limpar.
Private Sub cmdLimpar_Click()
    txtValor.text = Empty
    txtEntrada.text = Empty
    txtMeses.text = Empty
    txtJuros.text = Empty
    lblValFinanciado = Empty
    lblPrestMensal = Empty
    txtValor.SetFocus
End Sub


'procedure KeyPress, a ideia é identificar a tecla que o usuário está usando,  e que o foco passe de uma caixa de texto para outra quando teclar -Entrer-'
'Não funciona' sempre da erro na lina da -SendKey "{Tab}"'
' passou a funcionar depois da dica do stackoverflow acima Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)'
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Sendkeys "{Tab}"
KeyAscii = 0
    End If
End Sub



'Ao carregar o formuário atribuimos valoe 0 para esas variáveis'
Private Sub Form_Load()
    vValor = 0
    vEntrada = 0
    vJuros = 0
    vMeses = 0
End Sub


'Rotina de tratamento de erros, campo é do tipo Currency e não pode aceitar valores Text'
'A mensagem de erro aparecerar assim que o campo perder o Foco por isso LostFocus'
Private Sub txtValor_LostFocus()
    On Error GoTo Valor_Errado 'definição da rotina'
    vValor = CCur(txtValor.text) 'se a variável receber um valor Text'
    Prestação
Valor_Errado:
        If Err = vbKeyReturn Then ' se entrar no erro aparecerá essa mensagem'
            MsgBox "Dado inválido na digitação do VALOR DO FINANCIAMENTO", vbExclamation + vbApplicationModal, "Aviso"
           txtValor.text = InputBox("Informa um VALOR correto do FINANCIAMENTO: ", "Valor do Financiamento") 'janela ImputBox, aparecerá para o user informar outro valor'
        Resume 0
        End If
End Sub


'configuações txtBox Entrada
Private Sub txtEntrada_LostFocus()
    On Error GoTo Entrada_Errada
    vEntrada = CCur(txtEntrada.text)
    Prestação
Entrada_Errada:
    If Err = vbKeyReturn Then
        MsgBox "Dado inválido na digitação da ENTRADA DO FINANCIAMENTO", vbExclamation + vbApplicationModal, "Aviso"
txtEntrada.text = InputBox("Informe uma ENTRADA correta para o FINANCIAMENTO:", "Entrada do Financiamento")
        Resume 0
    End If
End Sub


'configuração txtBox Meses
Private Sub txtJuros_LostFocus()
    On Error GoTo Juros_Errado
    vJuros = CSng(txtJuros.text)
    vJuros = vJuros / 100 'VERIFICAR SE VAI DAR CERTO'
    Prestação
Juros_Errado:
      If Err = vbKeyReturn Then
        MsgBox "Dado inválido na digitação do JUROS DO FINANCIAMENTO", vbExclamation + vbApplicationModal, "Aviso"
        txtJuros.text = InputBox("Informe o JUROS correto para o FINANCIAMENTO:", "Juros do Financiamento")
     Resume 0
    End If
End Sub


'configuração txtBox Meses
Private Sub txtMeses_LostFocus()
    On Error GoTo Meses_Errado
    vMeses = CInt(txtMeses.text)
    Prestação
Meses_Errado:
    If Err = vbKeyReturn Then
        MsgBox "Dado inválido na digitação dos MESES DO FINANCIADO", vbExclamation + vbApplicationModal, "Aviso"
        txtMeses.text = InputBox("Informe os Meses corretos para o FINANCIAMENTO:", "Meses do Financiamento")

     End If
End Sub

