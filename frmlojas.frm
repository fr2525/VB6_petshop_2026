VERSION 5.00
Begin VB.Form frmlojas 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Lojas"
   ClientHeight    =   3495
   ClientLeft      =   2775
   ClientTop       =   1170
   ClientWidth     =   7545
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.CheckBox ChkSenha 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      Caption         =   "Senha?"
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   6390
      TabIndex        =   5
      Top             =   1800
      Width           =   915
   End
   Begin VB.TextBox TxtTelefone 
      Enabled         =   0   'False
      Height          =   300
      Left            =   4365
      TabIndex        =   3
      Top             =   1755
      Width           =   1335
   End
   Begin VB.TextBox TxtCpf_cnpj 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1245
      TabIndex        =   4
      Top             =   1695
      Width           =   2070
   End
   Begin VB.TextBox TxtEndereco 
      Enabled         =   0   'False
      Height          =   300
      Left            =   1245
      TabIndex        =   2
      Top             =   1215
      Width           =   6090
   End
   Begin VB.TextBox TxtNome 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1275
      TabIndex        =   1
      Top             =   720
      Width           =   4140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   840
      Left            =   1980
      TabIndex        =   7
      Top             =   2340
      Width           =   2865
      Begin VB.CommandButton cmddesfaz 
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   1455
         Picture         =   "frmlojas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         Caption         =   "&Sair"
         Height          =   540
         Left            =   2145
         Picture         =   "frmlojas.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   540
         Left            =   90
         Picture         =   "frmlojas.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   765
         Picture         =   "frmlojas.frx":0366
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Informações da Loja"
      BeginProperty Font 
         Name            =   "MV Boli"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   435
      Left            =   1755
      TabIndex        =   15
      Top             =   120
      Width           =   3405
   End
   Begin VB.Label LblCodLoja 
      BackStyle       =   0  'Transparent
      Caption         =   "codloja"
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   45
      TabIndex        =   13
      Top             =   750
      Width           =   585
   End
   Begin VB.Label LblTelefone 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Telefone:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3555
      TabIndex        =   10
      Top             =   1785
      Width           =   705
   End
   Begin VB.Label lblcgc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   285
      TabIndex        =   9
      Top             =   1725
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   315
      TabIndex        =   6
      Tag             =   "SALARIO:"
      Top             =   1215
      Width           =   855
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   705
      TabIndex        =   0
      Tag             =   "NOME:"
      Top             =   750
      Width           =   465
   End
End
Attribute VB_Name = "frmlojas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodLoja = gRs("id")
   Me.TxtNome.Text = gRs("nome")
   If Not IsNull(gRs!endereco) Then Me.TxtEndereco.Text = gRs!endereco
   If Not IsNull(gRs("cpf_cnpj")) Then Me.TxtCpf_cnpj.Text = gRs("cpf_cnpj")
   If Not IsNull(gRs("telefone")) Then Me.TxtTelefone.Text = gRs("telefone")
   If Not IsNull(gRs("senha")) Then Me.ChkSenha.Value = IIf(gRs("senha") = True, 1, 0)
   
End Sub

Private Sub cmddesfaz_Click()
  lIncluir = False
  ' Carrega_tela
  cmdEditar.Enabled = True
  CmdSair.Enabled = True
  cmdUpdate.Enabled = False
  cmddesfaz.Enabled = False
  Desabilita Me
End Sub

Private Sub cmdEditar_Click()
   cmdEditar.Enabled = False
   CmdSair.Enabled = False
   cmdUpdate.Enabled = True
   cmddesfaz.Enabled = True
   Habilita Me
   'Me.TxtNome.SetFocus
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
   gSql = "UPDATE tb_lojas SET Nome = '" & Me.TxtNome.Text & "',"
   gSql = gSql & " endereco = '" & Me.TxtEndereco.Text & "', "
   gSql = gSql & " cpf_cnpj = '" & Me.TxtCpf_cnpj.Text & "',"
   gSql = gSql & " telefone = '" & Me.TxtTelefone.Text & "',"
   gSql = gSql & " senha = " & Me.ChkSenha.Value & ","
   gSql = gSql & " operador = " & f_nulo(gncodoperador, 99) & ", datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
   gSql = gSql & " WHERE id = " & Val(Me.LblCodLoja.Caption)
   cnnLocal.Execute gSql
   cmdEditar.Enabled = True
   CmdSair.Enabled = True
   cmdUpdate.Enabled = False
   cmddesfaz.Enabled = False
   Desabilita Me
     
End Sub

Private Sub Form_Activate()
   Abre_Le_rst
    
   limpa_tela Me
   
   frmlojas.LblCodLoja.Caption = ""
   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         gSql = "INSERT INTO tb_lojas (Nome,endereco,CPF_CNPJ,telefone,celular,senha, operador, datatual) "
         gSql = gSql & " VALUES ( '" & Me.TxtNome.Text & "','"
         gSql = gSql & Me.TxtEndereco.Text & "','"
         gSql = gSql & Me.TxtTelefone.Text & "','"
         gSql = gSql & Me.ChkSenha.Value & ","
         gSql = gSql & "'" & f_nulo(gncodoperador, 99) & "','" & Format(Date, "yyy-mm-dd") & "') "
         cnnLocal.Execute gSql
         
         Abre_Le_rst
         Me.LblCodLoja.Caption = gRs!loja
         cmdEditar_Click
         lPrimeiro = True
      Else
         cmdEditar.Enabled = True
         CmdSair.Enabled = True
         cmdUpdate.Enabled = False
         cmddesfaz.Enabled = False
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
      cmdEditar.Enabled = True
      CmdSair.Enabled = True
      cmdUpdate.Enabled = False
      cmddesfaz.Enabled = False
      lIncluir = False
      lPrimeiro = False
      If gRs.State = adStateOpen Then
         gRs.Close
      End If
   End If
   
   Desabilita Me
   
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
     
End Sub

Private Sub Form_Load()
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
   End Sub

Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub

Private Sub Abre_Le_rst()
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   
   gSql = "select * FROM tb_lojas"
   gRs.Open gSql, cnnLocal, adOpenKeyset, adLockOptimistic
End Sub

