VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form frmUsuarios 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Cadastro de Usuários"
   ClientHeight    =   7230
   ClientLeft      =   1935
   ClientTop       =   960
   ClientWidth     =   8295
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   8295
   ShowInTaskbar   =   0   'False
   Tag             =   "cadvend"
   Begin VB.ComboBox cmbNivel 
      Height          =   315
      ItemData        =   "FrmUsuarios.frx":0000
      Left            =   5760
      List            =   "FrmUsuarios.frx":000D
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   2310
      Width           =   2025
   End
   Begin VB.CheckBox chkAtivo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00808000&
      Caption         =   "Ativo?"
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
      Height          =   360
      Left            =   6765
      TabIndex        =   6
      Top             =   1845
      Width           =   1035
   End
   Begin VB.TextBox txtSenha1 
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4485
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1875
      Width           =   1770
   End
   Begin VB.TextBox txtComissao 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   4020
      MaxLength       =   2
      TabIndex        =   8
      Top             =   2325
      Width           =   480
   End
   Begin VB.TextBox txtsalario 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   2325
      Width           =   1785
   End
   Begin VB.TextBox txtSenha 
      BorderStyle     =   0  'None
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1095
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1875
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1785
      TabIndex        =   10
      Top             =   6210
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "FrmUsuarios.frx":0036
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "FrmUsuarios.frx":0130
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Refresh"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "FrmUsuarios.frx":02A2
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "FrmUsuarios.frx":0414
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "FrmUsuarios.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Picture         =   "FrmUsuarios.frx":05F8
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin VB.TextBox txtNome 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      MaxLength       =   100
      TabIndex        =   0
      Top             =   540
      Width           =   6750
   End
   Begin VB.TextBox txtLogin 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   1095
      MaxLength       =   19
      TabIndex        =   1
      Top             =   990
      Width           =   1740
   End
   Begin VB.TextBox txtCelular 
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   6585
      TabIndex        =   2
      Top             =   975
      Width           =   1260
   End
   Begin VB.TextBox TXTeMAIL 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   285
      Left            =   1095
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1425
      Width           =   6750
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2985
      Left            =   390
      TabIndex        =   19
      Top             =   3000
      Width           =   7500
      _ExtentX        =   13229
      _ExtentY        =   5265
      _Version        =   393216
      Rows            =   5
      Cols            =   6
      FixedCols       =   0
      Appearance      =   0
      FormatString    =   $"FrmUsuarios.frx":06F2
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   4575
      TabIndex        =   29
      Top             =   2355
      Width           =   375
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Redigite a senha"
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
      Left            =   2940
      TabIndex        =   28
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salário"
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
      Left            =   360
      TabIndex        =   27
      Top             =   2370
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
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
      Left            =   480
      TabIndex        =   26
      Top             =   1050
      Width           =   480
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Cadastro de Usuários"
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
      TabIndex        =   25
      Top             =   60
      Width           =   3465
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Comissão"
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
      Left            =   3090
      TabIndex        =   24
      Top             =   2340
      Width           =   840
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nível"
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
      Left            =   5145
      TabIndex        =   23
      Top             =   2355
      Width           =   465
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Senha"
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
      Left            =   360
      TabIndex        =   22
      Top             =   1905
      Width           =   555
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
      Left            =   480
      TabIndex        =   21
      Top             =   570
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
      Top             =   1035
      Width           =   660
   End
   Begin VB.Label lblcodclie 
      BackStyle       =   0  'Transparent
      Caption         =   "id"
      ForeColor       =   &H0080FFFF&
      Height          =   270
      Left            =   135
      TabIndex        =   18
      Top             =   555
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
      TabIndex        =   17
      Top             =   1470
      Width           =   510
   End
End
Attribute VB_Name = "frmUsuarios"
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

   Me.lblcodclie = gRs("id")
   Me.txtNome.Text = gRs("nome")
   If Not IsNull(gRs("login")) Then Me.txtLogin.Text = gRs("login")
   If Not IsNull(gRs("email")) Then Me.TXTeMAIL.Text = gRs("email")
   If Not IsNull(gRs("celular")) Then Me.txtCelular.Text = gRs("celular")
   If Not IsNull(gRs("senha")) Then Me.txtSenha.Text = gRs("senha")
   If Not IsNull(gRs("salario")) Then Me.txtsalario.Text = gRs("salario")
   If Not IsNull(gRs("comissao")) Then Me.txtComissao.Text = gRs("comissao")
   
   Call sConectaLocal
   strSql = ""
   strSql = strSql & "select id,descricao from tb_niveis"
   Set Rstemp = New ADODB.Recordset
   Rstemp.Open strSql, cnnLocal, 1, 2
   Rstemp.MoveFirst
   Me.cmbNivel.Clear
        
   Indcmb = 0
   Do While Not Rstemp.EOF
      
       Me.cmbNivel.AddItem (Rstemp!descricao)
       Me.cmbNivel.ItemData(Me.cmbNivel.NewIndex) = Rstemp!id
       If gRs!nivel = Rstemp!id Then
           Me.cmbNivel.ListIndex = Indcmb
       End If
       Rstemp.MoveNext
       Indcmb = Indcmb + 1
   Loop
   Rstemp.Close
   Set Rstemp = Nothing
     
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
    If MsgBox("Deseja realmente apagar o Cliente " & Me.txtNome.Text & " ?", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tb_usuarios where id = " & Int(Me.lblcodclie.Caption)
        cnnLocal.Execute gSql
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
       suInsert
       cnnLocal.Execute gSql
       lIncluir = False
   Else
      gSql = "UPDATE  tb_usuarios SET Nome = '" & Me.txtNome.Text & "',"
      gSql = gSql & " login = '" & Me.txtLogin.Text & "',"
      gSql = gSql & " celular = '" & Me.txtCelular.Text & "',"
      gSql = gSql & " email = '" & f_nulo(Me.TXTeMAIL.Text, " ") & "',"
      gSql = gSql & " senha = '" & f_nulo(Me.txtSenha.Text, " ") & "',"
      gSql = gSql & " salario = " & f_nulo(Me.txtsalario.Text, 0) & ","
      gSql = gSql & " comissao = " & f_nulo(Me.txtComissao.Text, 0) & ","
      gSql = gSql & " ativo = " & f_nulo(Me.chkAtivo.Value, 0) & ","
      gSql = gSql & " operador = '" & f_nulo(gncodoperador, 99) & "', datatual = '" & Format(Date, "yyyy-mm-dd") & "'"
      gSql = gSql & " WHERE id = " & Me.lblcodclie.Caption
      cnnLocal.Execute gSql
      
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
         cnnLocal.Execute gSql
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
      Me.lblcodclie.Caption = gRs!id
      If gRs.State = adStateOpen Then
         gRs.Close
      End If
      gSql = "Select * from tb_usuarios "
      gSql = gSql & " WHERE id = " & Val(Me.lblcodclie.Caption)
      gRs.Open gSql, cnnLocal, adOpenForwardOnly
      Carrega_tela
      Desabilita Me
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
  If KeyCode = vbKeyReturn Then Sendkeys "{TAB}"
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

   If gRs.State = adStateOpen Then
      gRs.Close
   End If
   sConectaLocal
   gSql = "select id,nome,login,email,celular,senha,ativo,nivel,salario,comissao"
   gSql = gSql & " FROM tb_usuarios"
   gRs.Open gSql, cnnLocal, adOpenKeyset
   
End Sub
Private Sub Carrega_Grid()

'Teste do MsHFlexgrid1 - eh eh eh
  MSFlexGrid1.row = 0
  
  With gRs
      '.MoveLast
      'nItem = .RecordCount
      '.MoveFirst
      MSFlexGrid1.Redraw = False
      MSFlexGrid1.Rows = 1
        
      Do While Not .EOF
         MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
         MSFlexGrid1.row = MSFlexGrid1.Rows - 1
         MSFlexGrid1.ColAlignment(-1) = flexAlignLeftCenter
            
         MSFlexGrid1.Col = 0: MSFlexGrid1.Width = 0: MSFlexGrid1.Text = f_nulo(!id, "")
         MSFlexGrid1.Col = 1: MSFlexGrid1.Text = f_nulo(!nome, "")
         MSFlexGrid1.Col = 2: MSFlexGrid1.Text = f_nulo(!login, "")
         MSFlexGrid1.Col = 3: MSFlexGrid1.Text = f_nulo(!celular, "")
         MSFlexGrid1.Col = 4: MSFlexGrid1.Text = f_nulo(!email, "")
         MSFlexGrid1.Col = 5: MSFlexGrid1.Text = f_nulo(!nivel, "")
         
         .MoveNext
         
       Loop
       MSFlexGrid1.FixedRows = 1
       MSFlexGrid1.Redraw = True
          
  End With
  
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
    
    .Col = 0:  lblcodclie.Caption = .Text: .CellBackColor = vbYellow
    .Col = 1:  txtNome.Text = .Text: .CellBackColor = vbYellow
    .Col = 2:  login.Text = .Text: .CellBackColor = vbYellow
    .Col = 3:  txtCelular.Text = .Text: .CellBackColor = vbYellow
    .Col = 4:  TXTeMAIL.Text = .Text: .CellBackColor = vbYellow
     
    .TopRow = .row
    
    '.Refresh
 
   End With
   If gRs.State = adStateOpen Then
       gRs.Close
   End If
   gSql = "Select * from tb_usuarios "
   gSql = gSql & " WHERE id = " & Val(Me.lblcodclie.Caption)
   gRs.Open gSql, cnnLocal, adOpenForwardOnly
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

    gSql = "INSERT INTO tb_usuarios (Nome,celular,email,login,senha,nivel,salario,comissao,ativo,operador, datatual)"
    gSql = gSql & " VALUES ('" & Me.txtNome.Text & "','"
    gSql = gSql & Me.txtCelular.Text & "','"
    gSql = gSql & Me.TXTeMAIL.Text & "','"
    gSql = gSql & Me.txtLogin.Text & "','"
    gSql = gSql & Me.txtSenha.Text & "','"
    gSql = gSql & Me.cmbNivel.Text & "','"
    gSql = gSql & Me.txtsalario.Text & "','"
    gSql = gSql & Me.txtComissao.Text & "','"
    gSql = gSql & Me.chkAtivo.Value & "',"
    gSql = gSql & f_nulo(gncodoperador, 99) & ",'" & Format(Date, "yyyy-mm-dd") & "')"

End Sub

Private Sub TxtBairro_GotFocus()
 With txtBairro
      .SelStart = 0
      .SelLength = Len(.Text)
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
    txtCelular.Text = Format(txtCelular.Text, "(00) 00000-0000")
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

Private Sub TxtEmail_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub

Private Sub TxtNome_GotFocus()
   With txtNome
      .SelStart = 0
      .SelLength = Len(.Text)
   End With
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub


Private Sub txtsalario_LostFocus()
  ' Verifica se o campo não está vazio
    If txtsalario.Text <> "" Then
        ' Converte o texto para valor numérico e formata como moeda
        ' A função FormatCurrency usa as configurações regionais do sistema
        txtsalario.Text = FormatCurrency(CDbl(txtsalario.Text))

    End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Or KeyAscii = 9 Then KeyAscii = 0
End Sub
