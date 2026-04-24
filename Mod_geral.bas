Attribute VB_Name = "Mod_geral"
'****************************************************
'variaveis de conexao



Public Declare Sub sqlite3_open Lib "sqlite.dll" (ByVal filename As String, ByRef handle As Long)
Public Declare Sub sqlite3_close Lib "sqlite.dll" (ByVal DB_Handle As Long)
Public Declare Function sqlite3_last_insert_rowid Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Public Declare Function sqlite3_changes Lib "sqlite.dll" (ByVal DB_Handle As Long) As Long
Public Declare Function sqlite_get_table Lib "sqlite.dll" (ByVal DB_Handle As Long, ByVal SQLString As String, ByRef ErrStr As String) As Variant()
Public Declare Function sqlite_libversion Lib "sqlite.dll" () As String
Public Declare Function number_of_rows_from_last_call Lib "sqlite.dll" () As Long
' Exemplo de declaração no VB6
Public Declare Function sqlite3_exec Lib "sqlite.dll" (ByVal hDb As Long, ByVal SQL As String, ByVal callback As Long, ByVal arg As Long, ByRef errMsg As Long) As Long

Public DBz As Long
Public DBFile As String
Public minfo As String ' sql error akan store kat sini
Public row As Variant
Public query As String ' public variable untuk sql query
Public numrows As Long

Public Cnn As New ADODB.Connection
Public CnnLocal As New ADODB.Connection
Public cmd As New ADODB.Command

''**** Para usar com o Sqlite ---> e a RC6
'Public CnnLite As cConnection
''
''**** Para usar com o Sqlite ---> e a RC6
'Public RsLite   As cRecordset
'Public RsLite1  As cRecordset
'Public RsLite2  As cRecordset
'Public RsLite3  As cRecordset

'*************************
'variaveis para recordsets
Public Rstemp       As New ADODB.Recordset
Public RsTemp1      As New ADODB.Recordset
Public Rstemp2      As New ADODB.Recordset
Public gRs           As New ADODB.Recordset
'
'variaveis pra controle de registro
Global Situacao_Registro As String
Global Dias_Uso_Sistema As Integer
Global ConsultaProd_Ped As Integer
Global flagConsultaPedProd As Boolean

Public gTransacao As Boolean
Public SQL  As String
Public tmpSQL As String
'
Public strSql  As String
Public strSql1 As String
Public strSql2 As String
Public strSql3 As String

Public strPesqProdProv As Boolean

Public strFormaPgto As String

Public gMensagem As String


Global sysNomeAcesso As String

'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Inicio *
'*************************************************************************************
'
Public STR_IP_COMPUTADOR As String

Public Function BuscaIP() As String
Dim NIC As Variant
Dim NICs As Object

sysNomeAcesso = "MASTER"

On Error GoTo errError

Set NICs = GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each NIC In NICs
   If NIC.IPEnabled Then
        BuscaIP = NIC.IpAddress(0)
    End If
Next NIC

'ou
'Dim IPConfig As Variant
'Dim IPConfigSet As Object
'Set IPConfigSet = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = TRUE")
'
'For Each IPConfig In IPConfigSet
' If Not IsNull(IPConfig.IPAddress) Then MsgBox IPConfig.IPAddress(0), vbInformation
'Next IPConfig

Exit Function
    
errError:
    
    If Err.Number <> 0 Then
        Err.Clear
    End If
    BuscaIP = ""

End Function
'
'*************************************************************************************
'*** Fabio Reinert ( Alemao) 06/2017 - Inclusão de captura de IP do cliente - Fim    *
'*************************************************************************************
'
Public Sub sConectaLocal()
         
    On Error GoTo Erro_sConectaLocal
  
  
'DBFile = App.Path & "\petdb.db"
'sqlite3_open DBFile, DBz


Set CnnLocal = New ADODB.Connection
CnnLocal.ConnectionString = "DRIVER=SQLite3 ODBC Driver;Database=" & App.Path & "/PETDB.db;"
CnnLocal.Open
'*** Conexão com SQL Server
'cnnLocal.ConnectionString = "Sql_Server_Express"
'cnnLocal.Open
  
'ou
'connstring = "Provider=SQL Server Native Client 11.0;Server=DESKTOP-TN96NN4\\SQLEXPRESS;Database=db_printer3d;trusted_connection=yes"
'CnnLocal.Open

  'Conexão com Mysql que vou deixar comentada para testar a conexão com Sqlite ---> Qualquer coisa eu volto atrás
'  Set CnnLocal = New ADODB.Connection
'  'Conexão com Mysql
'  CnnLocal.Open "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=petshop;User=root;Password=oyster;"

'  strConn = "Driver={MySQL ODBC 8.0 Unicode Driver};Server=localhost;Database=db_Printer3d;User=root;Password=;"
'  cnnLocal.Open strConn
                        
Exit Sub

Erro_sConectaLocal:
    Call sMostraErro("sConectaLocal", Err.Number, Err.Description)
    'Call Fecha_Formularios
    End

End Sub

Public Sub sMostraAviso(Optional ByVal pTitulo As String, Optional ByVal pTexto1 As String, _
                        Optional ByVal pTexto2 As String, _
                        Optional ByVal pTexto3 As String, _
                        Optional ByVal pTexto4 As String)
                        
    Dim fAviso As Form
    If IsMissing(pTexto2) Then
        pTexto2 = ""
    End If
    If IsMissing(pTexto3) Then
        pTexto3 = ""
    End If
    If IsMissing(pTexto4) Then
        pTexto4 = ""
    End If
    If IsMissing(pTitulo) Then
        pTitulo = "Aviso:"
    End If
    Set fAviso = New frmAviso
    fAviso.lblAviso1.Caption = pTexto1
    fAviso.lblAviso2.Caption = pTexto2
    fAviso.lblAviso3.Caption = pTexto3
    fAviso.lblAviso4.Caption = pTexto4
    fAviso.Caption = pTitulo
    fAviso.Show vbModal
    Unload fAviso
    Set fAviso = Nothing
End Sub

Public Sub sMostraErro(Optional ByVal pModulo, Optional ByVal pErroNumero, Optional ByVal pErroDesc)
        
    If pModulo = "" Then
        pModulo = "Geral"
    End If
    If pErroNumero = "" Then
       pErroNumero = Err.Number
    End If
    If pErroDesc = "" Then
       pErroDesc = Err.Description
    End If
    Call sMostraAviso("Atenção - Erro: ", "Contate a Info Sistemas informando o erro abaixo:", _
                      "No.erro: " & pErroNumero & " Descr.: " & pErroDesc, _
                      "Módulo do erro: " & pModulo, "Sistema será encerrado")
    'Call Fecha_Formularios
    End
End Sub

Sub SelText(object As Control)
    
    With object
        .SelStart = 0
        .SelLength = Len(object)
    End With

End Sub
