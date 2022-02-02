VERSION 5.00
Begin VB.Form FrmTest 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Mudando as Cores"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000007&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4185
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFundoProx 
      Caption         =   "Próxima"
      Height          =   360
      Left            =   3840
      TabIndex        =   8
      Top             =   3120
      Width           =   990
   End
   Begin VB.CommandButton cmdFundoAnt 
      Caption         =   "Anterior"
      Height          =   360
      Left            =   3840
      TabIndex        =   7
      Top             =   2520
      Width           =   990
   End
   Begin VB.CommandButton cmdTextoProx 
      Caption         =   "Próxima"
      Height          =   360
      Left            =   3840
      TabIndex        =   6
      Top             =   1920
      Width           =   990
   End
   Begin VB.CommandButton cmdTextoAnt 
      Caption         =   "Anterior"
      Height          =   360
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label LblFundo 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "15 - Branco Brilhante"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2520
      Width           =   1830
   End
   Begin VB.Label LblTexto 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0 - Preto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   795
   End
   Begin VB.Label lblTeste 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Teste das Cores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   600
      TabIndex        =   2
      Top             =   240
      Width           =   4215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cor do Fundo"
      Height          =   195
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cor do Texto?"
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   1200
      Width           =   1020
   End
End
Attribute VB_Name = "FrmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Declaração de variáveis em escopo Global'
'variáveis  que seram usadas em vários escopos'
Dim vCorTexto As Integer
Dim vCorFundo As Integer

'Fução que recebe em seu comportamento a mudança de cor com Switch'
Private Sub MudaCorTexto()
    lblTeste.ForeColor = QBColor(vCorTexto) 'A label  TESTE COR irá recebera variavel vCorTexto)'
         With LblTexto 'usando o comando With par não ter que repetir toda ver o tipo e o nome (LblTexto)'
            Select Case vCorTexto
            
                     Case 0
                        .Caption = "0 - Preto"
                    Case 1
                        .Caption = "1 - Azul"
                    Case 2
                        .Caption = "2 - Verde"
                    Case 3
                        .Caption = "3 - Ciano"
                    Case 4
                        .Caption = "4 - Vermelho"
                    Case 5
                        .Caption = "5 - Magenta"
                    Case 6
                        .Caption = "6 - Amarelo"
                    Case 7
                        .Caption = "7 - Branco"
                    Case 8
                        .Caption = "8 - Cinza"
                    Case 9
                        .Caption = "9 - Azul Claro"
                    Case 10
                        .Caption = "10 - Verde Claro"
                    Case 11
                        .Caption = "11 - Cinza Claro"
                    Case 12
                        .Caption = "12 - Vermelho Claro"
                    Case 13
                        .Caption = "13 - Magenta Claro"
                    Case 14
                        .Caption = "14 - Amarelo Claro"
                    Case Else
                        .Caption = "15 - Branco Brilhante"
            
            End Select
        
        End With
        
End Sub

'a cor do fundo muda, mas a legenda não,'
Private Sub MudaCorFundo()
    lblTeste.ForeColor = QBColor(vCorFundo)

    With lblTeste
         
        Select Case vCorFundo
                
            Case 0
                .BackColor = &H0&

            Case 1
                .BackColor = &HFF0000

            Case 2
                .BackColor = &HC000&

            Case 3
                .BackColor = &HFFFF00

            Case 4
                .BackColor = &HFF&

            Case 5
                .BackColor = &HFF00FF

            Case 6
                .BackColor = &HFFFF&

            Case 7
                .BackColor = &HFFFFFF

            Case 8
                .BackColor = &H808080

            Case 9
                .BackColor = &HFFC0C0

            Case 10
                .BackColor = &H80FF80

            Case 11
                .BackColor = &HFFFFC0

            Case 12
                .BackColor = &H8080FF

            Case 13
                .BackColor = &HFF00FF

            Case 14
                .BackColor = &H80FFFF

            Case Else
                .BackColor = &HFFFFFF
                    
        End Select
                
    End With
            
    'lblTeste.ForeColor = QBColor(vCorFundo) 'A label  TESTE COR irá recebera variavel vCorTexto)'
    With LblFundo 'usando o comando With par não ter que repetir toda ver o tipo e o nome (LblTexto)'

        Select Case vCorFundo
            
            Case 0
                .Caption = "0 - Preto"

            Case 1
                .Caption = "1 - Azul"

            Case 2
                .Caption = "2 - Verde"

            Case 3
                .Caption = "3 - Ciano"

            Case 4
                .Caption = "4 - Vermelho"

            Case 5
                .Caption = "5 - Magenta"

            Case 6
                .Caption = "6 - Amarelo"

            Case 7
                .Caption = "7 - Branco"

            Case 8
                .Caption = "8 - Cinza"

            Case 9
                .Caption = "9 - Azul Claro"

            Case 10
                .Caption = "10 - Verde Claro"

            Case 11
                .Caption = "11 - Cinza Claro"

            Case 12
                .Caption = "12 - Vermelho Claro"

            Case 13
                .Caption = "13 - Magenta Claro"

            Case 14
                .Caption = "14 - Amarelo Claro"

            Case Else
                .Caption = "15 - Branco Brilhante"
            
        End Select
        
    End With
            
End Sub



'------------------------------------------------------------------------------------------------------'
Private Sub Form_Load()
    vCorTexto = 0
    vCorFundo = 15
End Sub

'-------------------------------------------------------------------------------------------------------'
'Config botão Anterior Texto'
Private Sub cmdTextoAnt_Click()
    vCorTexto = vCorTexto - 1
    If vCorTexto < 0 Then vCorTexto = 15
    MudaCorTexto
End Sub

'config botão Próximo Texo'
Private Sub cmdTextoProx_Click()
    vCorTexto = vCorTexto + 1
    If vCorTexto > 15 Then vCorTexto = 0
    MudaCorTexto
End Sub

'-------------------------------------------------------------------------------------------------------'
'Config botão Anterior Fundo'
Private Sub cmdFundoAnt_Click()
    vCorFundo = vCorFundo - 1
    If vCorFundo < 0 Then vCorFundo = 15
    MudaCorFundo
End Sub

'Config botão proximo Fundo'
Private Sub cmdFundoProx_Click()
    vCorFundo = vCorFundo + 1
    If vCorFundo > 15 Then CvCorFndo = 0
    MudaCorFundo
End Sub

