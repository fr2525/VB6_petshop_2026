VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmTipos 
   BackColor       =   &H00808000&
   Caption         =   "Tipos de Pets"
   ClientHeight    =   5280
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   1200
      TabIndex        =   2
      Top             =   4320
      Width           =   4245
      Begin VB.CommandButton cmdDesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2868
         Picture         =   "frmTipos.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   8
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmTipos.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   7
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmTipos.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   6
         Tag             =   "&Add"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1494
         Picture         =   "frmTipos.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   5
         Tag             =   "&Delete"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   807
         Picture         =   "frmTipos.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   4
         Tag             =   "&Refresh"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2181
         Picture         =   "frmTipos.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "&Update"
         Top             =   135
         Width           =   615
      End
   End
   Begin VB.TextBox txtAnimal 
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
      Left            =   180
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3690
      Width           =   6390
   End
   Begin MSComctlLib.ListView lstTipos 
      Height          =   3165
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   5583
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
End
Attribute VB_Name = "frmTipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim iTipoOperacao As Integer

Private Sub Carrega_Colunas_Tipos()
    With lstTipos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Codigo", 300, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Descrição", 4900, lvwColumnLeft
    End With
End Sub

Private Sub MontaColunas_Tipos()
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & " SELECT ID,DESCRICAO FROM TAB_tipos_pet ORDER BY DESCRICAO"
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, CnnLocal, 1, 2
    If Rstemp.RecordCount <> 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            lstTipos.ListItems.Add X, , Rstemp!id
            
            If Not IsNull(Rstemp!DESCRICAO) Then
                lstTipos.ListItems(X).SubItems(1) = Rstemp!DESCRICAO
            Else
                lstTipos.ListItems(X).SubItems(1) = ""
            End If
'            If Not IsNull(Rstemp!RAZAO_SOCIAL) Then
'                List_Atendimentos.ListItems(X).SubItems(2) = UCase(Rstemp!RAZAO_SOCIAL)
'            Else
'                  List_Atendimentos.ListItems.Add(X).SubItems(2) = "Fornecedor n o Encontrado...!"
'            End If
'            If Not IsNull(Rstemp!VALOR_TOTAL) Then
'                List_Atendimentos.ListItems(X).SubItems(3) = Format(Rstemp!VALOR_TOTAL, "0.00")
'            Else
'                List_Atendimentos.ListItems.Add(X).SubItems(3) = ""
'            End If
            
            Rstemp.MoveNext
        Next
        'lstTipos.SetFocus
       
    Else
        MsgBox "Sem registros", vbOKOnly
        'fmeListaPedidos.Visible = False
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmdAdd_Click()
    txtAnimal.Enabled = True
    txtAnimal.SetFocus
    txtAnimal.text = ""
    cmdAdd.Enabled = False
    cmdEditar.Enabled = False
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = False
    cmdDesfaz.Enabled = True
    iTipoOperacao = 1
End Sub

Private Sub cmdDelete_Click()
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    
    If MsgBox("Tem certeza que deseja excluir o tipo de animal: " & Chr(13) & Chr(10) & _
                            Trim(lstTipos.SelectedItem.ListSubItems.Item(1)), vbYesNo) = vbYes Then
        If fExcluir_Tipo_Pet() Then
            cmdAdd.Enabled = True
            cmdDelete.Enabled = False
            cmdUpdate.Enabled = False
            lstTipos.ListItems.Clear
            Call MontaColunas_Tipos
            If lstTipos.ListItems.Count > 0 Then
                lstTipos.ListItems(1).Selected = True
                txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
            End If
        Else
            MsgBox "Erro ao excluir o tipo de PET: " & Err.Description
        End If
    End If
End Sub

Private Sub cmdUpdate_Click()
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Sub
    End If
    If fGravar_Tipo_Pet() Then
        cmdAdd.Enabled = True
        cmdUpdate.Enabled = True
        Me.cmdDelete.Enabled = True
        Me.cmdSair.Enabled = True
        Me.cmdEditar.Enabled = True
        cmdDesfaz.Enabled = False
        'cmd_Excluir.Enabled = true
        lstTipos.ListItems.Clear
        Call MontaColunas_Tipos
        lstTipos.ListItems(1).Selected = True
        txtAnimal.Enabled = False
        txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
        'Call cmd_Limpar_Click
    Else
        MsgBox "Erro ao incluir o tipo de PET: " & Err.Description
    End If
End Sub
Private Sub cmdEditar_Click()
   cmdAdd.Enabled = False
   cmdUpdate.Enabled = True
   Me.cmdDelete.Enabled = False
   Me.cmdSair.Enabled = True
   Me.cmdEditar.Enabled = False
   cmdDesfaz.Enabled = True
   'cmd_Excluir.Enabled = true
   txtAnimal.Enabled = True
End Sub
Private Sub cmddesfaz_Click()
    txtAnimal.text = ""
    'txtAnimal.SetFocus
    cmdAdd.Enabled = False
    cmdUpdate.Enabled = True
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()

   Me.Move (Screen.Width - Me.Width) / 2, _
           (Screen.Height - Me.Height) / 2
           
    Call Carrega_Colunas_Tipos
    Call MontaColunas_Tipos
    'lstTipos.ListItems = 1
    If lstTipos.ListItems.Count > 0 Then
        txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
    End If
    
End Sub

Private Function fGravar_Tipo_Pet()
    
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal invalido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
       Exit Function
    End If
    
    fGravar_Tipo_Pet = True
    
    On Error GoTo Erro_fGravar_Tipo_Pet
    If iTipoOperacao = 1 Then
        strSql = "INSERT INTO tab_tipos_pet (DESCRICAO, OPERADOR, DaTATUAL)"
        strSql = strSql + " VALUES( '" & UCase(txtAnimal.text) & "','" & sysNomeAcesso & "','" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "')"
        
    Else
        strSql = "UPDATE tab_tipos_pet SET DESCRICAO = '" & UCase(txtAnimal.text) & _
                                          "',OPERADOR = '" & sysNomeAcesso & _
                                          "', DaTATUAl = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & _
                                          "' WHERE ID = '" & lstTipos.SelectedItem.text & "'"
    End If
    CnnLocal.Execute strSql
    Exit Function
Erro_fGravar_Tipo_Pet:
    fGravar_Tipo_Pet = False
End Function

Private Function fExcluir_Tipo_Pet()
    
    fExcluir_Tipo_Pet = True
    
    On Error GoTo Erro_fExcluir_Tipo_Pet
    
    strSql = "DELETE from tab_tipos_pet WHERE ID = '" & lstTipos.SelectedItem.text & "'"
    CnnLocal.Execute strSql
    Exit Function
Erro_fExcluir_Tipo_Pet:
    fExcluir_Tipo_Pet = False
End Function

Private Sub lstTipos_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtAnimal.text = Trim(lstTipos.SelectedItem.ListSubItems.Item(1))
    txtAnimal.Enabled = True
    cmdUpdate.Enabled = True
    cmdDelete.Enabled = True
    Me.cmdDesfaz.Enabled = True
    iTipoOperacao = 2
End Sub

Private Sub lstTipos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If lstTipos.ListItems.Count > 0 Then
            SendKeys "{tab}"
        End If
    Else
        Call lstTipos_Click
    End If
End Sub

Private Sub txtAnimal_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    If Len(txtAnimal.text) = 0 Or txtAnimal.text = "" Then
       MsgBox "Tipo de Animal inválido. Favor corrigir", vbOKOnly
       txtAnimal.SetFocus
    Else
       cmdUpdate.SetFocus
    End If
End If
End Sub
