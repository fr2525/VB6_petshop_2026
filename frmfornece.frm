VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmfornec 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Fornecedores"
   ClientHeight    =   6135
   ClientLeft      =   2625
   ClientTop       =   510
   ClientWidth     =   7395
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox txtCpfCnpj 
      Height          =   285
      Left            =   1620
      TabIndex        =   1
      Top             =   960
      Width           =   2055
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1605
      TabIndex        =   5
      Top             =   2430
      Width           =   5220
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2025
      Left            =   180
      TabIndex        =   11
      Top             =   2970
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   3572
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"frmfornece.frx":0000
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   795
      Left            =   1470
      TabIndex        =   10
      Top             =   5070
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmfornece.frx":00B4
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   780
         Picture         =   "frmfornece.frx":01AE
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Refresh"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1500
         Picture         =   "frmfornece.frx":0320
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Delete"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmfornece.frx":0492
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Add"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton CmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3540
         Picture         =   "frmfornece.frx":057C
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
      Begin VB.CommandButton cmddesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2880
         Picture         =   "frmfornece.frx":0676
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Update"
         Top             =   180
         Width           =   615
      End
   End
   Begin VB.TextBox TxtCelular 
      Height          =   285
      Left            =   5265
      TabIndex        =   2
      Top             =   990
      Width           =   1560
   End
   Begin VB.TextBox TxtEndereco 
      Height          =   285
      Left            =   1605
      TabIndex        =   4
      Top             =   1935
      Width           =   5220
   End
   Begin VB.TextBox TxtNome 
      Height          =   285
      Left            =   1605
      TabIndex        =   3
      Top             =   1410
      Width           =   5220
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Fornecedores"
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
      Left            =   1155
      TabIndex        =   19
      Top             =   195
      Width           =   4395
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   0
      Left            =   1035
      TabIndex        =   18
      Tag             =   "NOME:"
      Top             =   2460
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   660
      TabIndex        =   17
      Top             =   1020
      Width           =   825
   End
   Begin VB.Label LblCelular 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   4515
      TabIndex        =   9
      Top             =   1020
      Width           =   525
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   2
      Left            =   750
      TabIndex        =   7
      Tag             =   "NOME:"
      Top             =   1935
      Width           =   735
   End
   Begin VB.Label LblCodfor 
      BackStyle       =   0  'Transparent
      Caption         =   "Id"
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   165
      TabIndex        =   0
      Top             =   1035
      Width           =   390
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   1
      Left            =   1020
      TabIndex        =   6
      Tag             =   "NOME:"
      Top             =   1470
      Width           =   465
   End
End
Attribute VB_Name = "frmfornec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private lIncluir As Boolean
Private pRsFornec As New ADODB.Recordset

Private Sub Abre_Le_rst()
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   
  gSql = "select * FROM tb_fornecedores order by nome"
  gRs.Open gSql, cnnLocal, adOpenKeyset
  
End Sub

Private Sub cmdAdd_Click()
   
   limpa_tela Me
   
   Me.LblCodfor.Caption = ""
   Me.txtCpfCnpj.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()

    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar este Fornecedor ? ", vbYesNo, "Atenção") = vbYes Then
       gSql = "delete * from tb_fornecedores where id = " & Me.LblCodfor.Caption
       cnnLocal.Execute gSql
       Abre_Le_rst
       Carrega_Grid
       gRs.MoveFirst
       Carrega_tela
       Desabilita Me
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Fornecedor " & Chr(13) & "Instrucao Sql = '" & _
            cSql & "'  "
End Sub


Private Sub cmddesfaz_Click()
  
  lIncluir = False
  
  ' Carrega_tela
  Desabilita Me
  MSFlexGrid1_Click
  Me.cmdUpdate.Enabled = False
  Me.cmddesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.CmdSair.Enabled = True
  Me.cmdDelete.Enabled = True

End Sub

Private Sub cmdEditar_Click()
   ' Carrega_tela
   Habilita Me
   Me.TxtNome.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmddesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.CmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
   
End Sub

Private Sub CmdSair_Click()
   Unload Me
End Sub

Private Sub cmdUpdate_Click()
 
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   If lIncluir Then
      gSql = "SELECT cpf_cnpj from tb_fornecedores "
      gSql = gSql & " WHERE cpf_cnpj = '" & Me.txtCpfCnpj.Text & "'"
      pRsFornec.Open gSql, cnnLocal, adOpenKeyset
      If pRsFornec.BOF And pRsFornec.EOF Then
      Else
         MsgBox "CNPJ já cadastrado", vbOKOnly, "Atenção " & gOperador
         Exit Sub
      End If
      pRsFornec.Close

      gSql = "INSERT INTO tb_fornecedores (Nome,cpf_cnpj,endereco,email,celular,operador, datatual"
      gSql = gSql & ") "
      gSql = gSql & "VALUES ('" & Me.TxtNome.Text & "','"
      gSql = gSql & Me.txtCpfCnpj.Text & "','"
      gSql = gSql & Me.TxtEndereco.Text & "','"
      gSql = gSql & Me.txtEmail.Text & "','"
      gSql = gSql & Me.TxtCelular.Text & "',"
      gSql = gSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"
      cnnLocal.Execute gSql
      lIncluir = False
   Else
      gSql = "UPDATE tb_fornecedores SET Nome = '" & Me.TxtNome.Text & "',"
      gSql = gSql & "cpf_CNPJ = '" & Me.txtCpfCnpj.Text & "',"
      gSql = gSql & "endereco = '" & Me.TxtEndereco.Text & "',"
      gSql = gSql & "email = '" & Me.txtEmail.Text & "',"
      gSql = gSql & "celular = '" & Me.TxtCelular.Text & "',"
      gSql = gSql & " operador = " & f_nulo(gncodoperador, 99) & ", datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
      gSql = gSql & " WHERE id = " & Me.LblCodfor.Caption
      cnnLocal.Execute gSql
      
   End If
       
   Abre_Le_rst
   
   Carrega_Grid
   gRs.MoveFirst
   Carrega_tela
   Desabilita Me
      
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
End Sub

Private Sub Form_Load()
   Abre_Le_rst
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
   
   Me.LblCodfor.Caption = ""
    If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         'rst.AddNew
         With gRs
           .AddNew
           !nome = ""
           .Update
         End With
         cmdEditar_Click
         lPrimeiro = True
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      Carrega_tela
    
      Desabilita Me
      lIncluir = False
      lPrimeiro = False
   End If
   Carrega_Grid
   
   lIncluir = False

End Sub

Private Sub Form_Unload(Cancel As Integer)
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub MSFlexGrid1_Click()
 Dim oldrow As Long
 Dim lcColGrid As Double

  If MSFlexGrid1.row = 1 Then
     lcColGrid = MSFlexGrid1.Col
     MSFlexGrid1.Col = lcColGrid
     MSFlexGrid1.Sort = flexSortStringAscending
  End If
   
  oldrow = MSFlexGrid1.row
  
  MSFlexGrid1.row = 0
  
  With MSFlexGrid1
    .Redraw = False
    Do While True
       .row = .row + 1
       For ix = 0 To .Cols - 1
           .Col = ix: .CellBackColor = vbWhite
       Next
       If .row = .Rows - 1 Then
          Exit Do
       End If
    Loop
    .Redraw = True
    
    .row = oldrow
    
    .Col = 0:   LblCodfor.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:   TxtNome.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:   TxtCelular.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:   txtEmail.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:   txtCpfCnpj.Text = .Text: .CellBackColor = vbYellow
    .Col = 5:   TxtEndereco.Text = .Text: .CellBackColor = vbYellow
    .TopRow = .row
    
End With
     
   Desabilita Me
   Me.cmdUpdate.Enabled = False
   Me.cmddesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.CmdSair.Enabled = True
   Me.cmdDelete.Enabled = True


End Sub


Private Sub Carrega_tela()
   'Limpa as variaveis da tela se caso ficarem com dados da outra tela
   limpa_tela Me
   'Carrega a tela com os dados do registro
   Me.LblCodfor.Caption = gRs("id")
   Me.TxtNome.Text = "" & gRs("Nome")
   Me.TxtEndereco.Text = "" & gRs("endereco")
   Me.txtEmail.Text = "" & gRs("email")
   Me.TxtCelular.Text = "" & gRs("celular")
   Me.txtCpfCnpj.Text = "" & gRs("cpf_cnpj")
   
End Sub

Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Rows = 1
      MSFlexGrid1.ColWidth(4) = 0
      MSFlexGrid1.ColWidth(5) = 0
      MSFlexGrid1.ColAlignment(-1) = flexAlignLeftCenter
      
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.Col = 0:  MSFlexGrid1.Text = f_nulo(!id, "")
         MSFlexGrid1.Col = 1:  MSFlexGrid1.Text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2:  MSFlexGrid1.Text = f_nulo(!celular, "")
         MSFlexGrid1.Col = 3:  MSFlexGrid1.Text = f_nulo(!email, "")
         MSFlexGrid1.Col = 4:  MSFlexGrid1.Text = f_nulo(!cpf_CNPJ, "")
         MSFlexGrid1.Col = 5:  MSFlexGrid1.Text = f_nulo(!endereco, "")
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
          
  End With
  
  End Sub

Private Sub TxtCelular_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
Private Sub txtCpfCnpj_Exit(Cancel As Integer)
    Dim strNumeros As String
    
    ' Remove quaisquer caracteres não numéricos
    strNumeros = Replace(txtCpfCnpj.Text, ".", "")
    strNumeros = Replace(strNumeros, "-", "")
    strNumeros = Replace(strNumeros, "/", "")
    
    ' Verifica o comprimento para aplicar a máscara correta
    If Len(strNumeros) = 11 Then
        ' Aplica máscara de CPF: ###.###.###-##
        txtCpfCnpj.Text = Format(strNumeros, "000\.000\.000\-00")
    ElseIf Len(strNumeros) = 14 Then
        ' Aplica máscara de CNPJ: ##.###.###/####-##
        txtCpfCnpj.Text = Format(strNumeros, "00\.000\.000\/0000\-00")
    Else
        ' Caso não seja nem CPF nem CNPJ, você pode limpar ou exibir um aviso.
        MsgBox "Número de dígitos inválido. Digite 11 dígitos para CPF ou 14 para CNPJ."
        txtCpfCnpj.Text = ""
        Cancel = True ' Mantém o foco no campo se inválido
    End If
    
    ' Aqui você chamaria a função de validação (opcional, mas recomendado)
    ' If Not Fu_consistir_CgcCpf(strNumeros) Then
    '     MsgBox "CPF/CNPJ inválido pela regra do dígito verificador."
    '     Cancel = True
    ' End If
End Sub

Private Sub TxtCelular_LostFocus()
    TxtCelular.Text = Format(TxtCelular.Text, "(00) 00000-0000")
End Sub

Private Sub txtCpfCnpj_KeyPress(KeyAscii As Integer)
    ' Permite apenas números, Backspace, e caracteres de controle básicos
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtCpfCnpj_LostFocus()
 Dim strNumeros As String
    
    ' Remove quaisquer caracteres não numéricos
    strNumeros = Replace(txtCpfCnpj.Text, ".", "")
    strNumeros = Replace(strNumeros, "-", "")
    strNumeros = Replace(strNumeros, "/", "")
    
    ' Verifica o comprimento para aplicar a máscara correta
    If Len(strNumeros) = 11 Then
        ' Aplica máscara de CPF: ###.###.###-##
        txtCpfCnpj.Text = Format(strNumeros, "000\.000\.000\-00")
    ElseIf Len(strNumeros) = 14 Then
        ' Aplica máscara de CNPJ: ##.###.###/####-##
        txtCpfCnpj.Text = Format(strNumeros, "00\.000\.000\/0000\-00")
    Else
        ' Caso não seja nem CPF nem CNPJ, você pode limpar ou exibir um aviso.
        MsgBox "Número de dígitos inválido. Digite 11 dígitos para CPF ou 14 para CNPJ."
        txtCpfCnpj.Text = ""
        Cancel = True ' Mantém o foco no campo se inválido
    End If
    
    ' Aqui você chamaria a função de validação (opcional, mas recomendado)
    ' If Not Fu_consistir_CgcCpf(strNumeros) Then
    '     MsgBox "CPF/CNPJ inválido pela regra do dígito verificador."
    '     Cancel = True
    ' End If
End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
