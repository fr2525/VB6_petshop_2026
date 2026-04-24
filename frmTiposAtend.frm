VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmServicos 
   BackColor       =   &H00808000&
   Caption         =   "Serviços"
   ClientHeight    =   6885
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1080
      TabIndex        =   8
      Top             =   5640
      Width           =   4245
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2181
         Picture         =   "frmTiposAtend.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   807
         Picture         =   "frmTiposAtend.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "&Refresh"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1494
         Picture         =   "frmTiposAtend.frx":026C
         Style           =   1  'Graphical
         TabIndex        =   12
         Tag             =   "&Delete"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmTiposAtend.frx":03DE
         Style           =   1  'Graphical
         TabIndex        =   11
         Tag             =   "&Add"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmTiposAtend.frx":04C8
         Style           =   1  'Graphical
         TabIndex        =   10
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdDesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2868
         Picture         =   "frmTiposAtend.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   9
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
   End
   Begin VB.TextBox txtDuracao 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5160
      MaxLength       =   3
      TabIndex        =   3
      Top             =   4860
      Width           =   750
   End
   Begin VB.TextBox txtValor 
      Alignment       =   1  'Right Justify
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   300
      MaxLength       =   14
      TabIndex        =   2
      Top             =   4860
      Width           =   1860
   End
   Begin VB.TextBox txtServico 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   240
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3960
      Width           =   6270
   End
   Begin MSComctlLib.ListView lstservicos 
      Height          =   3135
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   5530
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "minutos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   6000
      TabIndex        =   7
      Top             =   4890
      Width           =   750
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Duração :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   5160
      TabIndex        =   6
      Top             =   4440
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   270
      TabIndex        =   5
      Top             =   4500
      Width           =   840
   End
   Begin VB.Label lbl_Animal 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   255
      TabIndex        =   4
      Top             =   3630
      Width           =   1380
   End
End
Attribute VB_Name = "frmServicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer
Private Sub Nomes_Colunas()
    With lstservicos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Código", 0, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4000, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Valor", 1400, lvwColumnRight
        .ColumnHeaders.Add 4, , "Duração", 900, lvwColumnRight
    End With
End Sub

Private Sub Dados_Colunas()
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT IDservico,DESCRICAOservico,valorservico,tempoestservico FROM TAB_servicos ORDER BY DESCRICAOservico"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, CnnLocal, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            With lstservicos
                .ListItems.Add X, , Rstemp!IDservico
                
                If Not IsNull(Rstemp!DESCRICAOservico) Then
                    .ListItems(X).SubItems(1) = Rstemp!DESCRICAOservico
                Else
                    .ListItems(X).SubItems(1) = ""
                End If
                If Not IsNull(Rstemp!VALORservico) Then
                    .ListItems(X).SubItems(2) = Format(Rstemp!VALORservico, "###,##0.00")
                Else
                      .ListItems.Add(X).SubItems(2) = "0.00"
                End If
                If Not IsNull(Rstemp!TEMPOESTservico) Then
                    .ListItems(X).SubItems(3) = Format(Rstemp!TEMPOESTservico, "000")
                Else
                    .ListItems.Add(X).SubItems(3) = "00"
                End If
            End With
            Rstemp.MoveNext
        Next
    Else
        MsgBox "Sem registros", vbOKOnly
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmd_Adicionar_Click()
    txtServico.Enabled = True
    txtDuracao.Enabled = True
    txtValor.Enabled = True
    txtServico.SetFocus
    txtServico.Text = ""
    txtDuracao.Text = ""
    txtValor.Text = ""
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = False
    iTipoOperacao = 1

End Sub

Private Sub cmd_Excluir_Click()
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o Serviço: " & Chr(13) & Chr(10) & _
                            Trim(lstservicos.SelectedItem.ListSubItems.Item(1)), vbYesNo) = vbYes Then
        If fExcluir_Servico() Then
            cmd_Adicionar.Enabled = True
            cmd_Excluir.Enabled = False
            cmd_Gravar.Enabled = False
            lstservicos.ListItems.Clear
            Call Dados_Colunas
            If lstservicos.ListItems.Count > 0 Then
                lstservicos.ListItems(1).Selected = True
                txtServico.Text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
            End If
        Else
            MsgBox "Erro ao excluir o Serviço: " & Err.Description
        End If
    End If

End Sub

Private Sub cmd_Gravar_Click()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
        MsgBox "Descrição de serviço inválida. Favor corrigir", vbOKOnly
        txtServico.SetFocus
        Exit Sub
    End If
    
    If Val(txtValor.Text) = 0 Then
        If MsgBox("Campo Valor do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbYesNo) = vbNo Then
            txtValor.SetFocus
            Exit Sub
        End If
    End If
    
    If Val(txtDuracao.Text) = 0 Then
        If MsgBox("Campo Tempo de duração do serviço não está preenchido. " & Chr(13) & Chr(10) & "Deseja continuar e gravar assim mesmo? ", vbYesNo) = vbNo Then
            txtDuracao.SetFocus
            Exit Sub
        End If
    End If
    
    If fGravar_Servico() Then
        cmd_Adicionar.Enabled = True
        cmd_Gravar.Enabled = False
        cmd_Limpar.Enabled = False
        'cmd_Excluir.Enabled = true
        lstservicos.ListItems.Clear
        Call Dados_Colunas
        lstservicos.ListItems(1).Selected = True
        txtServico.Text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If

End Sub

Private Sub cmd_Limpar_Click()
    txtServico.Text = ""
    txtValor.Text = "0.00"
    txtDuracao.Text = ""
    'txtServico.SetFocus
    cmd_Adicionar.Enabled = False
    cmd_Gravar.Enabled = True

End Sub

Private Sub cmd_Sair_Click()
    Unload Me
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
    If MsgBox("Deseja realmente apagar este Pet " & Me.txtNome.Text & " ?", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tb_ where id = " & Int(Me.lblcodclie.Caption)
        CnnLocal.Execute gSql
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

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call Nomes_Colunas
    Call Dados_Colunas
    'lstServicos.ListItems = 1
    If lstservicos.ListItems.Count > 0 Then
        txtServico.Text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
    End If
End Sub

Private Sub lstservicos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtServico.Text = Trim(lstservicos.SelectedItem.ListSubItems.Item(1))
    txtServico.Enabled = True
    txtValor.Text = Format(lstservicos.SelectedItem.ListSubItems.Item(2), "###,##0.00")
    txtValor.Enabled = True
    txtDuracao.Text = lstservicos.SelectedItem.ListSubItems.Item(3)
    txtDuracao.Enabled = True
    cmd_Gravar.Enabled = True
    cmd_Excluir.Enabled = True
    cmd_Limpar.Enabled = True
    iTipoOperacao = 2
End Sub

Private Sub lstservicos_KeyPress(KeyAscii As Integer)
    If lstservicos.ListItems.Count > 0 Then
        SendKeys "{tab}"
    End If
End Sub

Private Sub txtDuracao_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtServico.SetFocus
    Else
       cmd_Gravar.SetFocus
    End If
End If

End Sub

Private Function fGravar_Servico()
    
    If Len(txtServico.Text) = 0 Or txtServico.Text = "" Then
       MsgBox "Descrição do serviço inválida. Favor corrigir", vbOKOnly
       txtServico.SetFocus
       Exit Function
    End If
    
    fGravar_Servico = True
    
    On Error GoTo Erro_fGravar_Servico
    
    'ID,DESCRICAO,valor,tempo_est
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_servicos (DESCRICAO, VALOR, TEMPO_EST, OPERADOR, DT_ATUALIZA)"
        strSql = strSql + " VALUES( '" & UCase(txtServico.Text) & "',"
        strSql = strSql + Replace(txtValor.Text, ",", ".") & "," & txtDuracao.Text & ",'"
        strSql = strSql + sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_servicos SET DESCRICAO = '" & UCase(txtServico.Text) & _
                                          "',VALOR =   " & Replace(txtValor.Text, ",", ".") & _
                                          ",tempo_est = " & txtDuracao.Text & _
                                          ",OPERADOR = '" & sysNomeAcesso & _
                                          "', DT_ATUALIZA = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = '" & lstservicos.SelectedItem.Text & "'"
                                          
    End If
    CnnLocal.Execute strSql
    Exit Function
    
Erro_fGravar_Servico:
    fGravar_Servico = False
End Function

Private Function fExcluir_Servico()
    
    fExcluir_Servico = True
    
    On Error GoTo Erro_fExcluir_Servico
    
    strSql = "DELETE from tab_servicos WHERE ID = '" & lstservicos.SelectedItem.Text & "'"
    CnnLocal.Execute strSql
    Exit Function
Erro_fExcluir_Servico:
    fExcluir_Servico = False
End Function

Private Sub txtServico_GotFocus()
     Call SelText(txtServico)
End Sub

Private Sub txtServico_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        If Len(Trim(txtServico.Text)) = 0 Then
            MsgBox "Obrigatório Informar Descrição do Serviço.", vbInformation, "Aviso"
            txtServico.SetFocus
            Exit Sub
        End If

        SendKeys "{tab}"
    End If
End Sub

Private Sub txtServico_LostFocus()
    txtValor.Text = Format(0, "###,##0.00")
End Sub


Private Sub txtValor_GotFocus()
Call SelText(txtValor)
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    Char = Chr(KeyAscii)
    KeyAscii = Asc(UCase(Char))
    If KeyAscii = 13 Then
        If Len(Trim(txtValor.Text)) = 0 Then
            MsgBox "Obrigatório Informar o Valor do Serviço.", vbInformation, "Aviso"
            txtValor.SetFocus
            Exit Sub
        End If

        SendKeys "{tab}"
    End If

End Sub
