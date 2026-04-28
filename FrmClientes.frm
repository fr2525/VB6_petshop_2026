VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form FrmClientes 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7230
   ClientLeft      =   1980
   ClientTop       =   1005
   ClientWidth     =   8295
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.TextBox txtEstado 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   5655
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2505
      Width           =   480
   End
   Begin VB.TextBox txtCEP 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   6690
      TabIndex        =   8
      Top             =   2505
      Width           =   1140
   End
   Begin VB.TextBox txtCidade 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      TabIndex        =   6
      Top             =   2505
      Width           =   3765
   End
   Begin VB.TextBox txtBairro 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      TabIndex        =   5
      Top             =   2100
      Width           =   6750
   End
   Begin VB.TextBox txtEndereco 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   4
      Top             =   1710
      Width           =   6765
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1785
      TabIndex        =   9
      Top             =   6210
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "FrmClientes.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "FrmClientes.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "FrmClientes.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "FrmClientes.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "FrmClientes.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "FrmClientes.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox txtNome 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1110
      MaxLength       =   100
      TabIndex        =   0
      Top             =   540
      Width           =   6750
   End
   Begin VB.TextBox txtCPF_CNPJ 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      MaxLength       =   19
      TabIndex        =   1
      Top             =   915
      Width           =   1740
   End
   Begin VB.TextBox txtCelular 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   6585
      TabIndex        =   2
      Top             =   885
      Width           =   1260
   End
   Begin VB.TextBox TXTeMAIL 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1305
      Width           =   6750
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3105
      Left            =   390
      TabIndex        =   19
      Top             =   2985
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   5477
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"FrmClientes.frx":06BC
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Clientes"
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
      Left            =   2010
      TabIndex        =   27
      Top             =   60
      Width           =   3465
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CEP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6270
      TabIndex        =   26
      Top             =   2565
      Width           =   525
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4965
      TabIndex        =   25
      Top             =   2535
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   375
      TabIndex        =   24
      Top             =   2580
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   465
      TabIndex        =   23
      Top             =   2175
      Width           =   510
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   165
      TabIndex        =   22
      Top             =   1740
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   450
      TabIndex        =   21
      Top             =   585
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Celular:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   5730
      TabIndex        =   20
      Top             =   960
      Width           =   660
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CPF/CNPJ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   975
      Width           =   915
   End
   Begin VB.Label lblcodclie 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   135
      TabIndex        =   17
      Top             =   570
      Width           =   285
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "e-mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   465
      TabIndex        =   16
      Top             =   1350
      Width           =   510
   End
End
Attribute VB_Name = "FrmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private prsCliente As New ADODB.Recordset
'Private pQd As QueryDef

Private lIncluir As Boolean
Private lPrimeiro As Boolean

Private Sub Carrega_tela()
  
   limpa_tela Me

   Me.lblcodclie = gRs!id
   Me.txtNome.text = gRs!nome
   If Not IsNull(gRs!CPF) Then Me.txtCPF_CNPJ.text = gRs!CPF
   If Not IsNull(gRs!nome) Then Me.txtNome.text = gRs!nome
   If Not IsNull(gRs!email) Then Me.TXTeMAIL.text = gRs!email
   If Not IsNull(gRs!celular) Then Me.txtCelular.text = gRs!celular
   If Not IsNull(gRs!endereco) Then Me.txtEndereco.text = gRs!endereco
   If Not IsNull(gRs!bairro) Then Me.txtBairro.text = gRs!bairro
   If Not IsNull(gRs!cidade) Then Me.txtCidade.text = gRs!cidade
   If Not IsNull(gRs!estado) Then Me.txtEstado.text = gRs!estado
   If Not IsNull(gRs!cep) Then Me.txtCEP.text = gRs!cep
   'If Not IsNull(gRs("Ultcompra")) Then Me.TxtUltimaCompra.Text = Format(gRs("Ultcompra"), "dd/mm/YYYY")
     
End Sub
Private Sub cmdAdd_Click()

   Me.lblcodclie.Caption = ""
   limpa_tela Me
   Me.txtNome.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmdDesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.cmdSair.Enabled = False
   Me.cmdDelete.Enabled = False
 
   lIncluir = True
End Sub

Private Sub cmdDelete_Click()
    
    On Error GoTo ErroDelete
    
    'this may produce an error if you delete the last
    'record or the only record in the recordset
    If MsgBox("Deseja realmente apagar o Cliente " & Me.txtNome.text & " ?", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tab_clientes where id = " & Int(Me.lblcodclie.Caption)
        CnnLocal.Execute gSql
        CnnLocal.Close
        Abre_Le_rst
        Carrega_Grid
        gRs.MoveFirst
        Carrega_tela
        Desabilita Me
     End If
     Exit Sub
     
ErroDelete:
     MsgBox "Deu erro na exclusao do Cliente" & Chr(13) & "Instrucao Sql = '" & gSql & "'  "
End Sub

Private Sub cmddesfaz_Click()
  
  lIncluir = False
  
  Desabilita Me
   
  Me.cmdUpdate.Enabled = False
  Me.cmdDesfaz.Enabled = False
  Me.cmdEditar.Enabled = True
  Me.cmdAdd.Enabled = True
  Me.cmdSair.Enabled = True
  Me.cmdDelete.Enabled = True
 
End Sub

Private Sub cmdEditar_Click()
   
   Habilita Me
   Me.txtNome.SetFocus
   Me.cmdUpdate.Enabled = True
   Me.cmdDesfaz.Enabled = True
   Me.cmdEditar.Enabled = False
   Me.cmdAdd.Enabled = False
   Me.cmdSair.Enabled = False
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
      gSql = "SELECT nome,cpf from tab_clientes WHERE cpf = '" & txtCPF_CNPJ.text & "'"
      prsCliente.Open gSql, CnnLocal, adOpenKeyset
      If prsCliente.BOF And prsCliente.EOF Then
         prsCliente.Close
         suInsert
         CnnLocal.Execute strSql
      Else
         MsgBox "Cliente com CPF/CNPJ já cadastrado", vbOKOnly, "Atenção " & gOperador
         prsCliente.Close
         txtCPF_CNPJ.SetFocus
         Exit Sub
      End If
      lIncluir = False
   Else
      gSql = "UPDATE  tab_clientes SET Nome = '" & Me.txtNome.text & "',"
      gSql = gSql & " Cpf = '" & Me.txtCPF_CNPJ.text & "',"
      gSql = gSql & " celular = '" & Me.txtCelular.text & "',"
      gSql = gSql & " email = '" & f_nulo(Me.TXTeMAIL.text, " ") & "',"
      gSql = gSql & " endereco = '" & f_nulo(Me.txtEndereco.text, " ") & "',"
      gSql = gSql & " bairro = '" & f_nulo(Me.txtBairro.text, " ") & "',"
      gSql = gSql & " cidade = '" & f_nulo(Me.txtCidade.text, " ") & "',"
      gSql = gSql & " estado = '" & f_nulo(Me.txtEstado.text, " ") & "',"
      gSql = gSql & " cep = '" & f_nulo(Me.txtCEP.text, " ") & "',"
      gSql = gSql & " operador = '" & f_nulo(gncodoperador, 99) & "', datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
      gSql = gSql & " WHERE id = " & Me.lblcodclie.Caption
      CnnLocal.Execute gSql
      
   End If
     
   Abre_Le_rst
   
   Carrega_Grid
   
   gRs.MoveFirst
   
   Carrega_tela
   'Deixa os textbox desabilitados
   Desabilita Me
   
   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True
     
End Sub

Private Sub Form_Activate()
   Abre_Le_rst
      
   Me.lblcodclie.Caption = ""

   If gRs.BOF And gRs.EOF Then
      If MsgBox("Arquivo vazio. Incluir dados agora ?", vbYesNo, "Atenção ") = vbYes Then
         '**--> função para dar o INSERT --->
         suInsert
         CnnLocal.Execute strSql
         Abre_Le_rst
         Me.lblcodclie.Caption = gRs!codcli
         cmdEditar_Click
         lPrimeiro = True
         Exit Sub
      Else
         Desabilita Me
      End If
      
   Else
      gRs.MoveFirst
      'Me.lblcodclie.Caption = gRs!idcli
      'If gRs.State = adStateOpen Then
      '   gRs.Close
      'End If
      'gsql = "Select * from tab_clientes "
      'gsql = gsql & " WHERE idCli = " & Val(Me.lblcodclie.Caption)
      'gRs.Open gsql, CnnLocal, adOpenForwardOnly
      Carrega_tela
      'Desabilita Me
      If gRs.State = adStateOpen Then
         gRs.Close
      End If
      lIncluir = False
      lPrimeiro = False
   End If
   
   Abre_Le_rst
   
   Carrega_Grid
        
   lIncluir = False

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub Form_Load()
 
   'Centraliza a tela no video
   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
 
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If gRs.State = adStateOpen Then
      gRs.Close
   End If
    Screen.MousePointer = vbDefault
End Sub
Private Sub Abre_Le_rst()
   
   sConectaLocal
   strSql = "select id,cpf,nome,email,celular,endereco,bairro,cidade,estado,cep FROM tab_clientes"
   
   gRs.Open strSql, CnnLocal, adOpenKeyset
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.Row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Redraw = False
      MSFlexGrid1.Rows = 1
        
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.Row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.ColAlignment(-1) = flexAlignLeftCenter
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.text = f_nulo(!id, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.text = f_nulo(!celular, "")
         MSFlexGrid1.Col = 3: MSFlexGrid1.text = f_nulo(!email, "")
         MSFlexGrid1.Col = 4: MSFlexGrid1.text = f_nulo(!CPF, "")
         
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
       MSFlexGrid1.Redraw = True
          
  End With
  
  End Sub

Private Sub MSFlexGrid1_Click()
  Dim oldrow As Long
  Dim lcColGrid As Double
  
  If MSFlexGrid1.Row = 1 Then
     lcColGrid = MSFlexGrid1.Col
     MSFlexGrid1.Col = lcColGrid
     MSFlexGrid1.Sort = flexSortStringAscending
  End If
 
  oldrow = MSFlexGrid1.Row
  
  MSFlexGrid1.Row = 0
  
  With MSFlexGrid1
    .Redraw = False
    Do While True
       .Row = .Row + 1
       For ix = 0 To .Cols - 1
           .Col = ix: .CellBackColor = vbWhite
       Next
       If .Row = .Rows - 1 Then
          Exit Do
       End If
    Loop
    .Redraw = True
    
    .Row = oldrow
    
    .Col = 0:  lblcodclie.Caption = .text: .CellBackColor = vbYellow
    .Col = 1:  txtNome.text = .text: .CellBackColor = vbYellow
    .Col = 2:  txtCelular.text = .text: .CellBackColor = vbYellow
    .Col = 3:  TXTeMAIL.text = .text: .CellBackColor = vbYellow
    .Col = 4:  txtCPF_CNPJ.text = .text: .CellBackColor = vbYellow
     
    .TopRow = .Row
    
    '.Refresh
 
End With
If gRs.State = adStateOpen Then
   gRs.Close
End If
gSql = "Select * from tab_clientes "
gSql = gSql & " WHERE id = " & Val(Me.lblcodclie.Caption)
gRs.Open gSql, CnnLocal, adOpenForwardOnly
Carrega_tela
Desabilita Me

   Me.cmdUpdate.Enabled = False
   Me.cmdDesfaz.Enabled = False
   Me.cmdEditar.Enabled = True
   Me.cmdAdd.Enabled = True
   Me.cmdSair.Enabled = True
   Me.cmdDelete.Enabled = True

'gRs.Close

End Sub

Private Sub suInsert()

    strSql = "INSERT INTO tab_clientes (Nome,celular,email,cpf,endereco,bairro,cidade,estado,cep,operador, datatual)"
    strSql = strSql & " VALUES ('" & Me.txtNome.text & "','"
    strSql = strSql & Me.txtCelular.text & "','"
    strSql = strSql & Me.TXTeMAIL.text & "','"
    strSql = strSql & Me.txtCPF_CNPJ.text & "','"
    strSql = strSql & Me.txtEndereco.text & "','"
    strSql = strSql & Me.txtBairro.text & "','"
    strSql = strSql & Me.txtCidade.text & "','"
    strSql = strSql & Me.txtEstado.text & "','"
    strSql = strSql & Me.txtCEP.text & "',"
    strSql = strSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"

End Sub

Private Sub TxtBairro_GotFocus()
 With txtBairro
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtCelular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
    ' Opcional: Garante que apenas números sejam digitados
    If KeyAscii <> vbKeyBack Then ' Permite Backspace
        If Not IsNumeric(Chr(KeyAscii)) Then
            KeyAscii = 0 ' Ignora caracteres não numéricos
        End If
    End If

End Sub

Private Sub TxtCelular_LostFocus()
    txtCelular.text = Format(txtCelular.text, "(00) 00000-0000")
End Sub

'Private Sub txtCelular_Change()
'    Dim TextoLimpo As String
'
'    ' Remove todos os caracteres não numéricos
'    TextoLimpo = Replace(txtCelular.Text, "(", "")
'    TextoLimpo = Replace(TextoLimpo, ")", "")
'    TextoLimpo = Replace(TextoLimpo, "-", "")
'    TextoLimpo = Replace(TextoLimpo, " ", "")
'
'    ' Formata o número se tiver dígitos suficientes
'    If Len(TextoLimpo) > 2 Then
'        txtCelular.Text = "(" & Left(TextoLimpo, 2) & ") "
'        If Len(TextoLimpo) > 6 Then ' Ajuste para 8 ou 9 dígitos
'            txtCelular.Text = txtCelular.Text & Mid(TextoLimpo, 3, 4) & "-" & Right(TextoLimpo, Len(TextoLimpo) - 6)
'        Else
'            txtCelular.Text = txtCelular.Text & Mid(TextoLimpo, 3)
'        End If
'    End If
'
'    ' Move o cursor para o final do texto
'    txtCelular.SelStart = Len(txtCelular.Text)
'End Sub

Private Sub Txtcep_GotFocus()
   With txtCEP
      .SelStart = 0
      .SelLength = Len(.text)
   End With

End Sub

Private Sub Txtcep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtCEP_LostFocus()
 txtCEP.text = Format(txtCEP.text, "00000-000")
End Sub

Private Sub TxtCidade_GotFocus()
   With txtCidade
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtCidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TXTcPF_CNPJ_KeyPress(KeyAscii As Integer)
   ' Permite apenas números, Backspace, e caracteres de controle básicos
    If Not IsNumeric(Chr(KeyAscii)) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyDelete Then
        KeyAscii = 0
    End If
End Sub

Private Sub TXTcPF_CNPJ_LostFocus()
Dim strNumeros As String
    
    ' Remove quaisquer caracteres não numéricos
    strNumeros = Replace(txtCPF_CNPJ.text, ".", "")
    strNumeros = Replace(strNumeros, "-", "")
    strNumeros = Replace(strNumeros, "/", "")
    
    ' Verifica o comprimento para aplicar a máscara correta
    If Len(strNumeros) = 11 Then
        ' Aplica máscara de CPF: ###.###.###-##
        txtCPF_CNPJ.text = Format(strNumeros, "000\.000\.000\-00")
    ElseIf Len(strNumeros) = 14 Then
        ' Aplica máscara de CNPJ: ##.###.###/####-##
        txtCPF_CNPJ.text = Format(strNumeros, "00\.000\.000\/0000\-00")
    Else
        ' Caso não seja nem CPF nem CNPJ, você pode limpar ou exibir um aviso.
        MsgBox "Número de dígitos inválido. Digite 11 dígitos para CPF ou 14 para CNPJ."
        txtCPF_CNPJ.text = ""
        Cancel = True ' Mantém o foco no campo se inválido
    End If
    
    ' Aqui você chamaria a função de validação (opcional, mas recomendado)
    ' If Not Fu_consistir_CgcCpf(strNumeros) Then
    '     MsgBox "CPF/CNPJ inválido pela regra do dígito verificador."
    '     Cancel = True
    ' End If
End Sub

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtEndereco_GotFocus()
 With txtEndereco
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtEndereco_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtEstado_GotFocus()
  With txtEstado
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub txtEstado_LostFocus()
    txtEstado.text = UCase(txtEstado.text)
End Sub

Private Sub TxtNome_GotFocus()
   With txtNome
      .SelStart = 0
      .SelLength = Len(.text)
   End With
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub




