VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmCadPets 
   BackColor       =   &H00808000&
   Caption         =   "Cadastro de Pets"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   11760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   3435
      TabIndex        =   13
      Top             =   4020
      Width           =   4245
      Begin VB.CommandButton cmdDesfaz 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Desfaz"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2865
         Picture         =   "frmCadPets.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   19
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdSair 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Sair"
         Height          =   540
         Left            =   3555
         Picture         =   "frmCadPets.frx":00FA
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Incluir"
         Height          =   540
         Left            =   120
         Picture         =   "frmCadPets.frx":01F4
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "&Add"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Excluir"
         Height          =   540
         Left            =   1485
         Picture         =   "frmCadPets.frx":02DE
         Style           =   1  'Graphical
         TabIndex        =   16
         Tag             =   "&Delete"
         Top             =   150
         Width           =   615
      End
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Editar"
         Height          =   540
         Left            =   795
         Picture         =   "frmCadPets.frx":0450
         Style           =   1  'Graphical
         TabIndex        =   15
         Tag             =   "&Refresh"
         Top             =   135
         Width           =   615
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H00FFFF80&
         Caption         =   "&Salvar"
         Enabled         =   0   'False
         Height          =   540
         Left            =   2190
         Picture         =   "frmCadPets.frx":05C2
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "&Update"
         Top             =   150
         Width           =   615
      End
   End
   Begin MSComctlLib.ListView lstPets 
      Height          =   3375
      Left            =   270
      TabIndex        =   12
      Top             =   315
      Width           =   5505
      _ExtentX        =   9710
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtCuidEspec 
      Alignment       =   2  'Center
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
      Height          =   780
      Left            =   7410
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   2040
      Width           =   4110
   End
   Begin VB.TextBox TxtObserv 
      Alignment       =   2  'Center
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
      Height          =   780
      Left            =   7380
      MaxLength       =   2
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   2985
      Width           =   4110
   End
   Begin VB.TextBox txtDtNasc 
      Alignment       =   2  'Center
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
      Left            =   7410
      MaxLength       =   2
      TabIndex        =   7
      Top             =   1500
      Width           =   1170
   End
   Begin VB.ComboBox cmbTipos 
      Height          =   315
      Left            =   9450
      TabIndex        =   5
      Text            =   "Tipos"
      Top             =   1500
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.ComboBox cmbDonos 
      Height          =   315
      Left            =   7410
      TabIndex        =   2
      Text            =   "Donos"
      Top             =   330
      Visible         =   0   'False
      Width           =   4110
   End
   Begin VB.TextBox txtAnimal 
      Alignment       =   2  'Center
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
      Left            =   7410
      MaxLength       =   2
      TabIndex        =   0
      Top             =   930
      Width           =   4110
   End
   Begin VB.Label lblIdPet 
      BackStyle       =   0  'Transparent
      Caption         =   "idPet"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Cuidados Especiais :"
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
      Height          =   480
      Left            =   6165
      TabIndex        =   10
      Top             =   2070
      Width           =   1200
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Observ. :"
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
      Left            =   6300
      TabIndex        =   8
      Top             =   3015
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Dt.Nasc.  :"
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
      Left            =   6195
      TabIndex        =   6
      Top             =   1530
      Width           =   1170
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo :"
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
      Left            =   8760
      TabIndex        =   4
      Top             =   1530
      Width           =   600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Proprietário :"
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
      Left            =   5940
      TabIndex        =   3
      Top             =   360
      Width           =   1425
   End
   Begin VB.Label lbl_Animal 
      BackStyle       =   0  'Transparent
      Caption         =   "Pet  :"
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
      Left            =   6765
      TabIndex        =   1
      Top             =   900
      Width           =   600
   End
End
Attribute VB_Name = "frmCadPets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
   Me.lblcodclie.Caption = ""
   limpa_tela Me
   Me.TxtNome.SetFocus
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
    If MsgBox("Deseja realmente apagar este Pet " & Me.txtAnimal.Text & " ?", vbYesNo, "Atenção") = vbYes Then
        gSql = "delete from tab_pets where idPet = " & Int(Me.lblcodclie.Caption)
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
   Me.TxtNome.SetFocus
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
      gSql = "SELECT nome,cpf_cnpj from tb_clientes WHERE cpf_cnpj = '" & TxtCpf_cnpj.Text & "'"
      prsCliente.Open gSql, CnnLocal, adOpenKeyset
      If prsCliente.BOF And prsCliente.EOF Then
         prsCliente.Close
         suInsert
         CnnLocal.Execute gSql
      Else
         MsgBox "Cliente com CPF/CNPJ já cadastrado", vbOKOnly, "Atenção " & gOperador
         prsCliente.Close
         Txtcgc_cpf.SetFocus
         Exit Sub
      End If
      lIncluir = False
   Else
      gSql = "UPDATE  tab_clientes SET Nomecli = '" & Me.TxtNome.Text & "',"
      gSql = gSql & " Cpf_cnpj = '" & Me.TxtCpf_cnpj.Text & "',"
      gSql = gSql & " celular = '" & Me.txtCelular.Text & "',"
      gSql = gSql & " email = '" & f_nulo(Me.TXTeMAIL.Text, " ") & "',"
      gSql = gSql & " endereco = '" & f_nulo(Me.TxtEndereco.Text, " ") & "',"
      gSql = gSql & " bairro = '" & f_nulo(Me.txtBairro.Text, " ") & "',"
      gSql = gSql & " cidade = '" & f_nulo(Me.txtCidade.Text, " ") & "',"
      gSql = gSql & " estado = '" & f_nulo(Me.txtEstado.Text, " ") & "',"
      gSql = gSql & " cep = '" & f_nulo(Me.txtCEP.Text, " ") & "',"
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
         CnnLocal.Execute gSql
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
