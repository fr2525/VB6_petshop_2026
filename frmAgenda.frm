VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form frmAgenda 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8625
   ClientLeft      =   15
   ClientTop       =   -45
   ClientWidth     =   12090
   ControlBox      =   0   'False
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   12090
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNovoAgend 
      Caption         =   "Novo Agendamento"
      Height          =   945
      Left            =   300
      Picture         =   "frmAgenda.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   6960
      Width           =   1140
   End
   Begin VB.CommandButton cmdSair 
      Caption         =   "Sair"
      Height          =   945
      Left            =   10665
      Picture         =   "frmAgenda.frx":00EA
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "Gravar"
      Height          =   945
      Left            =   9385
      Picture         =   "frmAgenda.frx":01E4
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdLimpar 
      Caption         =   "Limpar"
      Height          =   945
      Left            =   8105
      Picture         =   "frmAgenda.frx":0356
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "Pets"
      Height          =   945
      Left            =   5545
      Picture         =   "frmAgenda.frx":0450
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdServicos 
      Caption         =   "Serviços"
      Height          =   945
      Left            =   6825
      Picture         =   "frmAgenda.frx":151A
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdTipos 
      Caption         =   "Tipos"
      Height          =   945
      Left            =   4265
      Picture         =   "frmAgenda.frx":225C
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   6960
      Width           =   1095
   End
   Begin VB.CommandButton cmdClientes 
      Caption         =   "Clientes"
      Height          =   945
      Left            =   2985
      Picture         =   "frmAgenda.frx":4826
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   6960
      Width           =   1095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   285
      TabIndex        =   16
      Top             =   360
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   661
      _Version        =   393216
      CalendarForeColor=   -2147483647
      CalendarTitleBackColor=   -2147483632
      CalendarTitleForeColor=   16776960
      CalendarTrailingForeColor=   128
      Format          =   129171456
      CurrentDate     =   36892
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "Detalhe"
      Height          =   6420
      Left            =   5640
      TabIndex        =   2
      Top             =   330
      Width           =   6225
      Begin VB.TextBox Text1 
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   24
         Top             =   2310
         Width           =   3780
      End
      Begin VB.ComboBox cmbServicos 
         Height          =   315
         Left            =   3840
         TabIndex        =   22
         Text            =   "Servicos"
         Top             =   3120
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox cmbDonos 
         Height          =   315
         Left            =   3120
         TabIndex        =   21
         Text            =   "Donos"
         Top             =   5130
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox cmbPets 
         Height          =   315
         Left            =   30
         TabIndex        =   20
         Text            =   "Pets"
         Top             =   5310
         Visible         =   0   'False
         Width           =   3165
      End
      Begin VB.ComboBox CmbHorario 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmAgenda.frx":58F0
         Left            =   1455
         List            =   "frmAgenda.frx":58F2
         TabIndex        =   18
         Text            =   "00:00"
         Top             =   390
         Width           =   1125
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   9
         Top             =   900
         Width           =   3780
      End
      Begin VB.TextBox txtDono 
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   8
         Top             =   1350
         Width           =   3780
      End
      Begin VB.TextBox txtTipoAtend 
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1800
         Width           =   3780
      End
      Begin VB.TextBox txtObeserva 
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
         Height          =   1680
         Left            =   1440
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   3870
         Width           =   3780
      End
      Begin VB.TextBox txtValor 
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
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   5
         Top             =   3270
         Width           =   1410
      End
      Begin VB.OptionButton OptSim 
         BackColor       =   &H00808000&
         Caption         =   "Sim"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   4
         Top             =   2805
         Width           =   645
      End
      Begin VB.OptionButton OptNao 
         BackColor       =   &H00808000&
         Caption         =   "Não"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Top             =   2820
         Width           =   705
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Especial :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   360
         TabIndex        =   23
         Top             =   2340
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Horário :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   420
         TabIndex        =   17
         Top             =   390
         Width           =   900
      End
      Begin VB.Label lbl_Animal 
         BackStyle       =   0  'Transparent
         Caption         =   "Pet :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   810
         TabIndex        =   15
         Top             =   900
         Width           =   600
      End
      Begin VB.Label lbl_Dono 
         BackStyle       =   0  'Transparent
         Caption         =   "Dono :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   660
         TabIndex        =   14
         Top             =   1350
         Width           =   780
      End
      Begin VB.Label lbl_TipoAtend 
         BackStyle       =   0  'Transparent
         Caption         =   "Serviço :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   420
         TabIndex        =   13
         Top             =   1830
         Width           =   1140
      End
      Begin VB.Label lbl_Atendido 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Atendido :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   180
         TabIndex        =   12
         Top             =   2850
         Width           =   1275
      End
      Begin VB.Label lbl_Obseerv 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Observ :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   10.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   300
         TabIndex        =   11
         Top             =   3870
         Width           =   1170
      End
      Begin VB.Label lbl_Valor 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00E0E0E0&
         Height          =   360
         Left            =   450
         TabIndex        =   10
         Top             =   3270
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      Caption         =   "Lista"
      ForeColor       =   &H80000008&
      Height          =   5805
      Left            =   240
      TabIndex        =   0
      Top             =   915
      Width           =   5205
      Begin MSComctlLib.ListView List_Atendimentos 
         Height          =   5190
         Left            =   240
         TabIndex        =   1
         Top             =   300
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   9155
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Horário"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Pet"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Atendimento"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Atendido?"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Duração"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observações"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aLista_Horario(48) As String

Private Sub Carrega_Colunas_Atendimentos()
    With List_Atendimentos
        .ListItems.Clear
        .ColumnHeaders.Clear
        .View = lvwReport
        .ColumnHeaders.Add 1, , "Horário", 700, lvwColumnLeft
        .ColumnHeaders.Add 2, , "Pet", 1900, lvwColumnLeft
        '.ColumnHeaders.Add 3, , "Dono", 1900, lvwColumnLeft
        .ColumnHeaders.Add 3, , "Atendimento", 1900, lvwColumnLeft
        '.ColumnHeaders.Add 5, , "Duração", 800, lvwColumnLeft
        '.ColumnHeaders.Add 6, , "Atendido", 850, lvwColumnLeft
        '.ColumnHeaders.Add 7, , "Observações", 4400, lvwColumnLeft
        '.ColumnHeaders.Add 8, , "Valor", 1200, lvwColumnRight
    End With
End Sub

Private Sub MontaAtendimentos(pData As Date)
    
    ldatainicio = Format(pData, "yyyy-mm-dd 00:00:00")
    ldatafim = Format(pData, "yyyy-mm-dd 23:59:59")
    
    Call sConectaLocal
    strSql = ""
    strSql = strSql & "SELECT a.dataAtend,a.idPet,a.TipoAtend,a.valor as valor_serv,tempoatend"
    strSql = strSql & " ,b.nome as nomePet,b.tipo_Pet,b.id_cli,E.nome AS dono"
    strSql = strSql & " FROM TAB_ATENDIMENTOS A"
    strSql = strSql & " INNER JOIN tab_pets B ON A.IdPet  = B.ID"
    strSql = strSql & " INNER JOIN TAB_TIPOS_pet C ON B.TIPO_Pet = C.ID"
    strSql = strSql & " INNER JOIN TAB_SERVICOS D ON A.TipoAtend = D.ID"
    strSql = strSql & " INNER JOIN TAB_clientes E ON B.ID_cli = E.ID  "
    strSql = strSql & " WHERE A.dataAtend >= '" & ldatainicio & "' and A.dataAtend <= '" & ldatafim & "'"
    
    Set Rstemp = New ADODB.Recordset
    Rstemp.Open strSql, CnnLocal, 1, 2
    If Rstemp.RecordCount > 0 Then
        Rstemp.MoveLast
        Rstemp.MoveFirst
        'fmeListaPedidos.Visible = True
        
        For X = 1 To Rstemp.RecordCount
            If Not IsNull(Rstemp!DATA_PED) Then
                List_Atendimentos.ListItems.Add X, , Format(Rstemp!DATA_PED, "DD/MM/YYYY")
            Else
                List_Atendimentos.ListItems.Add X, , ""
            End If
            If Not IsNull(Rstemp(0)) Then
                List_Atendimentos.ListItems(X).SubItems(1) = Rstemp(0)
            Else
                List_Atendimentos.ListItems(X).SubItems(1) = ""
            End If
            If Not IsNull(Rstemp!RAZAO_SOCIAL) Then
                List_Atendimentos.ListItems(X).SubItems(2) = UCase(Rstemp!RAZAO_SOCIAL)
            Else
                  List_Atendimentos.ListItems.Add(X).SubItems(2) = "Fornecedor não Encontrado...!"
            End If
            If Not IsNull(Rstemp!VALOR_TOTAL) Then
                List_Atendimentos.ListItems(X).SubItems(3) = Format(Rstemp!VALOR_TOTAL, "0.00")
            Else
                List_Atendimentos.ListItems.Add(X).SubItems(3) = ""
            End If
            
            Rstemp.MoveNext
        Next
        List_Atendimentos.SetFocus
       
    Else
        'MsgBox "Sem Atendimentos para a data selecionada", vbOKOnly
        'fmeListaPedidos.Visible = False
    End If
    
    Rstemp.Close
    Set Rstemp = Nothing
    
End Sub

Private Sub cmdClientes_Click()
    FrmClientes.Show vbModal
End Sub

Private Sub cmdNovo_Click()
   frmCadPets.Show vbModal
End Sub

Private Sub CmdSair_Click()
    Unload Me
End Sub

Private Sub cmdServicos_Click()
    frmServicos.Show vbModal
End Sub

Private Sub cmdTipos_Click()
    frmTipos.Show vbModal
End Sub

Private Sub DTPicker1_Change()
   'Print DTPicker1.Value
   Call MontaAtendimentos(DTPicker1.Value)
End Sub

Private Sub Form_Initialize()
   Dim i As Integer
   Dim sHora As String

   DTPicker1.Value = Format(Date, "dd/mm/yyyy")
   sHora = "07:00"   ' Estabelcemos um horario inicial que depopis pode ser parametrizado

   aLista_Horario(0) = sHora
   For i = 0 To 25   ' Vai até as 20:00 - Podemos ver parametrização depois
     sHora = DateAdd("n", 30, CDate(sHora))
     aLista_Horario(i) = Mid(sHora, 1, 5)
     CmbHorario.AddItem (aLista_Horario(i))
   Next

   Call Carrega_Colunas_Atendimentos
   Call MontaAtendimentos(Now)
End Sub
