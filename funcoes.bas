Attribute VB_Name = "funcoes"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" _
        (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, _
        ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" _
         (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, _
          ByVal lpFileName As String) As Long

Private Declare Function GetModuleFileName Lib "kernel32" _
         Alias "GetModuleFileNameA" _
         (ByVal hModule As Long, _
         ByVal lpFileName As String, _
         ByVal nSize As Long) As Long


'Global Sql As String
Global Titulo As String
Global vgResposta As String
Global i, X, t, gFlag As Long

'Global cPathODBC As String
'Global wFlex_id As String

'*** Variaveis para Configurações da máquina
Public gThousandSeparator As String
Public gDecimalSeparator As String
Public gListSeparator As String
Public gDateSeparator As String
Public gShortDateFormat As String
Public gLongDateFormat As String
Public gLanguage As String
Public gCountryName As String
Public gCountryId As String
Public gTimeId As String
Public gTimeSeparator As String
Public gCurrency As String
Public gCurrencyDigits As String
Public gDigits As String
Public gResposta As Integer
Public gVersao As String
Public gOrdem As Long

Public gSenha   As String
Private prsVendas As New ADODB.Recordset
Private prsItens As New ADODB.Recordset
'Public dbVr     As Database

'Num módulo:
Public Function f_ValidaData(ByVal sData _
       As String, Optional ByVal sFormato _
       As String) As Boolean
  Dim Dia As Integer, Dia_Pos As Integer
  Dim Mes As Integer, Mes_Pos As Integer
  Dim Ano As Integer, Ano_Pos As Integer
  Dim DDOk As Integer, MMOk As Integer
  Dim YYOk As Integer, i As Integer
  Dim m As Integer, Temp As String
  Dim sBst As Boolean

'P/ chamar, no evento desejado:
'VariávelBoolean = ValidaData(Data, Formato)

'Exemplo:
'  Dim bRESP As Boolean
'  bRESP = ValidaData("31/03/2000", "M/D/Y")
'  If bRESP Then
'    MsgBox "A data é válida!!!"
'  Else
'    MsgBox "A data NÃO é válida!!!"
'  End If
'Ele exibirá "A data é válida!!!"

  If IsMissing(sFormato) Then
    sFormato = "DD/MM/YYYY"
    'OU então, você pode pegar o formato
    'que estiver configurado no Windows.
  Else
     If Len(sFormato) = 0 Then
        sFormato = "DD/MM/YYYY"
     End If
  End If

  Temp = Replace(sData, "-", "/")
  sData = Temp
  Temp = Replace(sFormato, "-", "/")
  sFormato = Temp
  Temp = ""

  DDOk = 0
  MMOk = 0
  YYOk = 0

  For i = 1 To Len(sFormato)
    If UCase(Mid(sFormato, i, 1)) = "D" Then
      If DDOk > 2 Then
        f_ValidaData = False
        Exit Function
      Else
        DDOk = DDOk + 1
        If Dia_Pos = 0 Then
          Dia_Pos = Mes_Pos + Ano_Pos + 1
        End If
      End If
    ElseIf UCase(Mid(sFormato, i, 1)) = "M" Then
      If MMOk > 2 Then
        f_ValidaData = False
        Exit Function
      Else
        MMOk = MMOk + 1
        If Mes_Pos = 0 Then
          Mes_Pos = Dia_Pos + Ano_Pos + 1
        End If
      End If
    ElseIf UCase(Mid(sFormato, i, 1)) = "Y" Then
      If YYOk > 4 Then
        f_ValidaData = False
        Exit Function
      Else
        YYOk = YYOk + 1
        If Ano_Pos = 0 Then
          Ano_Pos = Dia_Pos + Mes_Pos + 1
        End If
      End If
    Else
      Select Case UCase(Mid(sFormato, i, 1))
        Case "D", "M", "Y", "/"
        Case Else
          f_ValidaData = False
          Exit Function
      End Select
    End If
  Next i

  If DDOk = 0 Or MMOk = 0 Then
    f_ValidaData = False
    Exit Function
  End If

  If YYOk = 0 Or YYOk > 4 Then
    f_ValidaData = False
    Exit Function
  End If

  If Not IsDate(sData) Then
    f_ValidaData = False
    Exit Function
  End If

  m = 0
  For i = 1 To Len(sData)
    If Mid(sData, i, 1) = "/" Or _
        i = Len(sData) Then
      If i = Len(sData) Then
        Temp = Temp & Mid(sData, i, 1)
      End If
      m = m + 1
      If m = 3 Then m = 4
      If Dia_Pos = m Then
        Dia = Temp
      ElseIf Mes_Pos = m Then
        Mes = Temp
      ElseIf Ano_Pos = m Then
        Ano = Temp
      End If
      Temp = ""
    Else
      Temp = Temp & Mid(sData, i, 1)
    End If
  Next i

  Select Case Mes
    Case 1, 3, 5, 7, 8, 10, 12
      If Dia < 1 Or Dia > 31 Then
        f_ValidaData = False
        Exit Function
      End If
    Case 4, 6, 9, 11
      If Dia < 1 Or Dia > 31 Then
        f_ValidaData = False
        Exit Function
      End If
    Case 2
      If Dia < 1 Or Dia > 29 Then
        f_ValidaData = False
        Exit Function
      ElseIf Dia = 29 Then
        sBst = False
        If Ano = 0 Then
          sBst = True
        ElseIf Ano Mod 4 = 0 Then
          sBst = True
          If Ano Mod 100 = 0 Then
            sBst = False
            If Ano Mod 400 = 0 Then
              sBst = True
            End If
          End If
        Else
          sBst = False
        End If
        If sBst = False Then
          f_ValidaData = False
          Exit Function
        End If
      End If
    Case Else
      f_ValidaData = False
      Exit Function
  End Select
  f_ValidaData = True

End Function

Private Function Replace(ByVal texto _
        As String, ByVal Isto As String, _
        ByVal PorIsto As String) As String
  Dim i As Long

  If Len(Isto) < 1 Then
    Replace = texto
    Exit Function
  End If

  For i = 1 To Len(texto)
    If Mid(texto, i, Len(Isto)) = Isto Then
      Replace = Replace & PorIsto
      i = i + (Len(Isto) - 1)
    Else
      Replace = Replace & Mid(texto, i, 1)
    End If
  Next i

'P/ usar, digamos que o TextBox Text1 contenha um texto
'como "Eu sei lá como isto funciona". Então, depois dessa linha:

'Text1.Text Replace(Text1.Text, "sei lá", "não sei")
'Text1 passará a conter "Eu não sei como isto funciona"

'Detalhe: Neste replace, você pode fazer algo como Replace("Texto","-/-","*")
' que ele substituirá sem problemas!!!!!!!!

End Function

Public Function Replead(ByVal texto _
        As String, ByVal Isto As String, _
        ByVal PorIsto As String) As String
  Dim i As Long

  If Len(Isto) < 1 Then
    Replead = texto
    Exit Function
  End If

  For i = 1 To Len(texto)
    If Val(Mid(texto, i, Len(Isto))) > 0 Then
       Replead = Replead & Mid(texto, i, Len(texto) - i)
       Exit For
    End If
    If Mid(texto, i, Len(Isto)) = Isto Then
      Replead = Replead & PorIsto
      i = i + (Len(Isto) - 1)
    Else
      Replead = Replead & Mid(texto, i, 1)
    End If
  Next i

'replace especial para tirar os zeros iniciais de um string e substituir por zeros
' por exemplo - Usado para imprimir com a formatação e objeto print #
End Function

Function UltimoDiaDoMes(Adate As Variant)

 On Error GoTo UltimoDiaErro
 Dim NextMonth As Variant
 
 NextMonth = DateAdd("m", 1, Adate)
 UltimoDiaDoMes = NextMonth - DatePart("d", NextMonth)

UltimoDiaDoMesExit:
   Exit Function

UltimoDiaErro:
   MsgBox "Erro em Ultimo dia do Mes", vbCritical
   Resume UltimoDiaDoMesExit
   
End Function

Public Function fuSomaHora(ByVal Ini As String, ByVal Fim As String, ByVal Extra As String) As String
Dim A, B, c, D, E, G, H As Double
Dim F As String

    If Val(Mid(Ini, 4, 2)) > 59 Or Val(Mid(Ini, 1, 2)) > 24 Then
        MsgBox "Hora Inicial Errada !"
        Exit Function
    End If
    If Val(Mid(Fim, 4, 2)) > 59 Or Val(Mid(Fim, 1, 2)) > 24 Then
        MsgBox "Hora Final Errada !"
        Exit Function
    End If
    If Val(Mid(Ini, 4, 2)) > 0 And Val(Mid(Ini, 1, 2)) > 24 Then
        MsgBox "Hora Inicial Errada !"
        Exit Function
    End If
    If Val(Mid(Fim, 4, 2)) > 0 And Val(Mid(Fim, 1, 2)) > 24 Then
        MsgBox "Hora Final Errada !"
        Exit Function
    End If
    
    A = Val(Mid(Ini, 4, 2)) * 60: A = A + (Val(Mid(Ini, 1, 2)) * 3600)       '*** Transforma Hora Entrada ***
    B = Val(Mid(Fim, 4, 2)) * 60: B = B + (Val(Mid(Fim, 1, 2)) * 3600)       '*** Transforma Hora Saída   ***
   ' G = Val(Mid(Almoco, 4, 2)) * 60: G = G + (Val(Mid(Almoco, 1, 2)) * 3600) '*** Transforma Hora Almoço  ***
    H = Val(Mid(Extra, 4, 2)) * 60: H = H + (Val(Mid(Extra, 1, 2)) * 3600)
    
    If B < A Then
        B = B + (86400) - A
        A = 0
        '*** Hora Final menor que a Inicial
    End If
    
    c = (B - A) + H        'Fim - Inicio - Almoco
    
    D = ((c - (c Mod 3600)) / 3600)
    If D < 0 Then D = 0
    
    E = (c Mod 3600) / 60
    If E < 0 Then E = 0
    
    F = Format(D, "00") & ":" & Format(E, "00")
    If F = "00:00" Then F = "24:00"
            
    fuSomaHora = F

End Function
'Public Sub PintaGrid(Grade As MSFlexGrid)
'  Dim oldrow As Long
'  Dim lcColGrid As Double
'
'  'If MsflexgridItens.Row = 1 Then
'  '   lcColGrid = MsflexgridItens.Col
'  '   MsflexgridItens.Col = lcColGrid
'  '   MsflexgridItens.Sort = flexSortStringAscending
'  'End If
'
'  If Grade.Rows = 1 Then
'     Exit Sub
'  End If
'
'  oldrow = Grade.Row
'
'  Grade.Row = 0
'
'  With Grade
'    .Redraw = False
'    Do While True
'       .Row = .Row + 1
'       For i = 0 To .Cols - 1
'           .Col = i: .CellBackColor = vbWhite
'       Next
'       If .Row = .Rows - 1 Then
'          Exit Do
'       End If
'    Loop
'    .Redraw = True
'
'    .Row = oldrow
'
'    For i = 0 To .Cols - 1
'        .Col = i: .CellBackColor = vbYellow
'    Next
'
'    .TopRow = .Row
'
'End With
'
'End Sub

Public Function lhoras(texto As String) As Variant
Dim wHora As Long
Dim wMinuto As Long
    
    If Len(texto) = 5 Then
        wHora = Val(Left(texto, 2)) * 3600
    ElseIf Len(texto) = 6 Then
        wHora = Val(Left(texto, 3)) * 3600
    End If
    wMinuto = Val(Right(texto, 2)) * 60
    lhoras = wHora + wMinuto
    
End Function
    
'Public Function HrCheia(strHora As Variant) As Single
'On Error GoTo NAO_FAZ
'Dim x As Integer          ' variável auxiliar
'Dim strHoraAux As String  ' var. auxiliar: manipula partes do número
'Dim Horas As Integer      ' graus
'Dim Mins As Integer       ' minutos
'Dim Segs As Integer       ' segundos
'Const Sep = ":"           ' separador (hh:mm:ss)
'
'    ' Valor das horas em formato string
'    strHora = Format$(strHora, "hh:mm")
'
'    ' Separa horas
'    x = InStr(strHora, Sep)
'
'    'Horas = IIf(x > 0, Val(Left$(strHora, x - 1)), 0)
'    Horas = IIf(x > 0, Val(Left$(strHora, x)), 0)
'
' '   If Horas > 24 Then
' '       MsgBox "Sr.(ª) " & Trim(LoginData.nmUsr) & ", " & Str$(Horas) & " - Valor incompatível com TBHHOR02.", 32, Titulo
' '       Exit Function
' '   End If
'
'    ' Separa minutos
'    strHoraAux = Right$(strHora, Len(strHora) - x)
'    Mins = Val(strHoraAux)
'
'    If Mins > 60 Then
'        MsgBox "Sr.(ª) " & Trim(LoginData.nmUsr) & ", " & Str$(Mins) & " - Valor incompatível com minutos.", 32, Titulo
'        Exit Function
'    End If    ' Valor final da função
'
'    HrCheia = Horas + (Mins / 60)
'
'Exit Function
'NAO_FAZ:
'    MsgBox "Ocorreu o erro n. " & Err & "."
'
'End Function

Public Function fuNumeros(KeyAscii As Integer, vlDecimal As Boolean, vlChrEspecial As String) As Integer
'Parâmetros:
'           KeyAscii      -> Valor ASC da tecla pressionada
'           vlDecimal     -> Flag que determina se função vai testar separador decimal
'           vlChrEspecial -> Qualquer caracter que seja válido (p.ex: '/', '-', ':', 'a', etc.)
'                            (Parâmetro pode conter mais de uma tecla especial, desde que o
'                             o conteúdo venha separado por vírgulas.)
'
   Dim i As Integer
   Dim j As Integer
   Dim vlEspeciais() As Integer    'Array contendo os caracteres especiais a serem comparados
  ' suConfigMachine
  
  'Dimensiona variáveis iniciais de busca para caracteres especiais
   j = 0
   ReDim Preserve vlEspeciais(j)
      
     'Testa se existem caracteres especiais a serem considerados
      If vlChrEspecial <> "" Then
           'Varre a string em busca dos caracteres
            For i = 1 To Len(vlChrEspecial)
              'Testa separador dos caracteres (,)
               If Mid(vlChrEspecial, i, 1) <> "," Then
                 'Guarda código ansi na array
                  vlEspeciais(j) = Asc(Mid(vlChrEspecial, i, 1))
               Else
                 'Incrementa nova posição na array para próximo caracter, se houver
                  j = j + 1
                  ReDim Preserve vlEspeciais(0 To j)
               End If
            Next
           'Varre a array para comparaçao da tecla pressionada
            For i = 0 To UBound(vlEspeciais)
              'Testa tecla pressionada comparando-a com os caracteres especiais
               If KeyAscii = vlEspeciais(i) Then
                 'Tecla é válida. Retorna seu valor
                  fuNumeros = KeyAscii
                  Exit Function
               End If
            Next
      End If
     'Não existem caracteres especial para testar
      Select Case KeyAscii
         'Teclas válidas
          Case 8, 13, 48 To 57
             fuNumeros = KeyAscii
         'Ponto decimal
          Case 44, 46
             If vlDecimal Then
                fuNumeros = Asc(gDecimalSeparator)
             Else
                fuNumeros = False
             End If
         'Teclas não válidas
          Case Else
             fuNumeros = False
      End Select
    
End Function
Public Function SoNumero(text As String) As String
    Dim i As Integer, j As String
    For i = 1 To Len(text)
        If Mid(text, i, 1) = "," Then
            j = j & "."
        Else
        If Asc(Mid(text, i, 1)) < 48 Or _
           Asc(Mid(text, i, 1)) > 57 Then
        Else
            j = j & Mid(text, i, 1)
        End If
        End If
        SoNumero = j
    Next
End Function

'Public Sub suGravaOrcamento()
'  Dim pcCodprod As String
'  Dim pnQtde As Double
'  Dim pnPreco As Double
'
'  gSql = "UPDATE tab_vendas SET "
'  gSql = gSql & " tipovenda = 0, dta_venda  = Cdate('" & Date & "')"
'  gSql = gSql & " where nsu = '" & Format(Str(gnSequencia), "000000000")
'  cnnLocal.Execute gSql
'
'  With FrmVendas.MSFlexGridItens
'     For i = 1 To .Rows - 1
'        .Col = 0
'        pcCodprod = .text
'        .Col = 2
'        pnQtde = Val(.text)
'        .Col = 3
'        pnPreco = CDbl(.text)
'        '*---> Insere nos Itens de Venda
'        gSql = "INSERT INTO tab_itemvenda (nsu,codprod,qtde,precounit,operador,datatual) "
'        gSql = gSql & " Values('" & Format(Str(gnSequencia), "000000000") & "','" & Format(pcCodprod, "000000") & "',"
'        gSql = gSql & pnQtde & "," & Replace(pnPreco, ",", ".")
'        gSql = gSql & ",'" & gOperador & "',Cdate('" & Date & "'))"
'        cnnLocal.Execute gSql
'     Next
'  End With
'
'End Sub

'Public Sub suGravaErros(rotina As String)
'Dim wtexto As String
'
'   ' wtexto = "INSERT INTO TBERR01 (Cd_erro,Dt_Erro,Ds_Erro,Ds_Funcao,Cd_Usuario)"
'   ' wtexto = wtexto & " VALUES ('" & Left(Err, 10) & "','" & futrocames(Format$(Now, "dd-mmm-yy")) & "','" & Left(strtran(Error(Err), "'", ""), 100) & "','" & Left(rotina, 100) & "','" & Left(LoginData.cdUsr, 10) & "')"
'   ' dbHr.Execute wtexto, 64
'
'    '*** Mensagem ao usuário caso ocorra erro ao gravar o cliente ***
'    wtexto = "Ao executar a operação ocorreu o erro nr. " & Err & ", " & Error(Err) & "."
'    wtexto = wtexto & Chr(13) & Chr(10)
'    wtexto = wtexto & Chr(13) & Chr(10)
'    wtexto = wtexto & "O programa continuará sua execução normal, "
'    wtexto = wtexto & "entretanto anote o erro e comunique ao administrador do sistema."
'    MsgBox "Sr.(ª) " & Trim(LoginData.nmUsr) & ", " & wtexto, 16, Titulo
'
'End Sub

Public Function strtran(ByVal cFull As String, ByVal cOld As String, ByVal cNew As String)
Dim nLoop As Single, cTemp As String

   cTemp = ""
   nLoop = 1
   cFull = UCase(cFull)
   While nLoop <= Len(cFull)
      If Mid$(cFull, nLoop, Len(cOld)) = UCase(cOld) Then
         cTemp = cTemp + cNew
         nLoop = nLoop + Len(cOld) - 1
      Else
         cTemp = cTemp + Mid$(cFull, nLoop, 1)
      End If
      nLoop = nLoop + 1
   Wend
   strtran = cTemp

End Function
Public Function Replicate(ByVal cOld As String, ByVal Tam As Double)
Dim nLoop As Single, cTemp As String

   cTemp = ""
   nLoop = 1
   While nLoop <= Tam
      cTemp = cTemp & cOld
      nLoop = nLoop + 1
   Wend
   Replicate = cTemp

End Function

Public Function f_conta(ByVal texto As String)
   f_conta = 0
   For X = 1 To Len(texto)
       If Val(Mid(texto, X, 1)) > 0 Then
          Exit For
       End If
       If Mid(texto, X, 1) = "." Then
          f_conta = f_conta + 1
       End If
       If Mid(texto, X, 1) = "0" Then
          f_conta = f_conta + 1
       End If
   Next

End Function

Public Function fuEncript(lSenha As String, lCodigo As String) As String
Dim lLen, lAlg As Integer
Dim lNova, lPasso As String
   
    lLen = Len(Trim(lCodigo))
    
    For i = 1 To lLen
        lAlg = lAlg + Asc(Mid(lCodigo, i, 1))
    Next i
    
    lAlg = Int(lAlg / 11)
    
    For i = 1 To Len(lSenha)
        lPasso = Asc(Mid(lSenha, i, 1))
        lPasso = Int((lPasso + lAlg) / 2) + lLen + Len(Trim(lSenha))
        lPasso = Chr(lPasso)
        lNova = lNova + lPasso
    Next i
    
    fuEncript = lNova
    
End Function

Function fuespacos(formControl As Control, Condicao As Integer, Retira As Integer) As String

   Dim campo As String
   
      Select Case Condicao
         'Retira espaços
         Case 0
            campo = Trim(formControl)
         'Adiciona espaços
         Case 1
            If Trim(formControl) <> "" Then
               campo = Space(Retira - Len(Trim(formControl))) & Trim(formControl)
            End If
      End Select
   fuespacos = campo

End Function

'
'RECEBE UMA STRING E A DEVOLVE SOMENTE COM OS SEUS CARACTERES ALFANUMÉRICOS
'
'EX: ? fuLimpaTexto("A B%CD*()_E F-2,45-5.78")
'      ABCDEF245578
'
Function fuLimpaTexto(ByVal vlTexto As String) As String

    Dim vlCont, vlChar As Integer
    Dim vlNewText As String
    
    vlTexto = Trim(vlTexto)
    vlNewText = ""
      For vlCont = 1 To Len(vlTexto)
          vlChar = Asc(Mid(vlTexto, vlCont, 1))
            If (vlChar > 47 And vlChar < 58) Or (vlChar > 64 And vlChar < 91) Then
               vlNewText = vlNewText & Chr(vlChar)
            End If
      Next vlCont
    fuLimpaTexto = vlNewText

End Function


Public Function f_nulo(campo As Variant, conteudo As Variant) As Variant
On Error GoTo NAO_FAZ
    
    If campo = "" Or IsNull(campo) Then
        f_nulo = conteudo
    Else
        f_nulo = campo
    End If
    
On Error GoTo 0

Exit Function
NAO_FAZ:
    On Error GoTo 0
    f_nulo = conteudo
    Exit Function

End Function

Function snulo(ByVal campo As Variant) As Variant
    
    If VarType(campo) = 0 Or VarType(campo) = 1 Then
        snulo = " "
    Else
        snulo = campo
    End If
    
End Function
'Preenche combobox e fill listbox
'como usar: CarregaControle NomeControle, "Nometabela", "CodigodoCampo","DescricaoCampo"
'
'CodigoCampo : é o identificador unico do campo . Ex: CodigoCliente
'DescricaoCampo : e o campo texto para exibir no controle. Ex: NomeCliente
'
Public Sub SuCarregaControle(Controle As Object, Tabela, CodigoCampo, DescricaoCampo As String)

On Error GoTo Erro

Dim rs As Recordset   'Declara um recorset
Dim SQL As String       'Declara uma string para a consulta SQL

Controle.Clear
'limpa o controle
SQL = ""
'limpa a string SQL
'Define a string SQL para selecionar os registros
SQL = "SELECT " & CodigoCampo & ", " & DescricaoCampo & " FROM " & Tabela
'abre o recorddset com os dados retornados
'Set rs = ConDb.OpenRecordset(sql, dbOpenForwardOnly)
Set rs = CnnLocal.Execute(SQL)
With rs
Do Until .EOF 'percorre o recordset ate o fim

  'inclui os itens correspondentes
  Controle.AddItem rs(DescricaoCampo)
  Controle.ItemData(Controle.NewIndex) = rs(CodigoCampo)
  .MoveNext

  Loop
  'fecha o recordset
  .Close
End With

Set rs = Nothing 'libera o recordset
Exit Sub

Erro: 'se houver erros faz o tratamento

If Err.Number <> 0 Then
  MsgBox ("Erro #: " & Str(Err.Number) & Err.Description)
  Exit Sub
End If

End Sub


'Public Sub suConfigMachine()
'
'    Static stRead As String
'    Static stLenprof As Integer
'
'   'Separador de milhares
'    stRead = Space(5)
'    stLenprof = GetPrivateProfileString("intl", "sThousand", ",", stRead, Len(stRead), "WIN.INI")
'    gThousandSeparator = Left$(stRead, stLenprof)

'   'Separador decimal
'    stRead = Space(5)
'    stLenprof = GetPrivateProfileString("intl", "sDecimal", ".", stRead, Len(stRead), "WIN.INI")
'    gDecimalSeparator = Left$(stRead, stLenprof)

'   'Separador de listas
'    stRead = Space(5)
'    stLenprof = GetPrivateProfileString("intl", "sList", ",", stRead, Len(stRead), "WIN.INI")
'    gListSeparator = Left$(stRead, stLenprof)

'   'Formato de datas abreviadas
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sShortDate", "dd/mm/yy", stRead, Len(stRead), "WIN.INI")
'    gShortDateFormat = Left$(stRead, stLenprof)

'   'Formato de datas completas
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sLongDate", "dd/mm/yy", stRead, Len(stRead), "WIN.INI")
'    gLongDateFormat = Left$(stRead, stLenprof)
   
'   'Separador de datas
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sDate", "/", stRead, Len(stRead), "WIN.INI")
'    gDateSeparator = Left$(stRead, stLenprof)

'   'Idioma
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sLanguage", "us", stRead, Len(stRead), "WIN.INI")
'    gLanguage = Left$(stRead, stLenprof)

'   'País
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sCountry", "eua", stRead, Len(stRead), "WIN.INI")
'    gCountryName = Left$(stRead, stLenprof)

'   'id Horário
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "iTime", "1", stRead, Len(stRead), "WIN.INI")
'    gTimeId = Left$(stRead, stLenprof)
   
'   'Separador de horário
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sTime", ":", stRead, Len(stRead), "WIN.INI")
'    gTimeSeparator = Left$(stRead, stLenprof)

'   'Moeda
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "sCurrency", "R$", stRead, Len(stRead), "WIN.INI")
'    gCurrency = Left$(stRead, stLenprof)
    
'   'Quantidade de dígitos da moeda
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "iCurrDigits", "2", stRead, Len(stRead), "WIN.INI")
'    gCurrencyDigits = Left$(stRead, stLenprof)
    
'   'Quantidade de dígitos de números
'    stRead = Space(20)
'    stLenprof = GetPrivateProfileString("intl", "iDigits", "2", stRead, Len(stRead), "WIN.INI")
'    gDigits = Left$(stRead, stLenprof)

'End Sub

'Public Function VerDataHora(formControl As Control) As Boolean
'
'   If Trim(formControl) <> "" Then
'        If IsDate(formControl) Then
'            formControl = Format(formControl, "mm/yyyy")
'            VerDataHora = True
'        Else
'            MsgBox "Sr. " & LoginData.nmUsr & ", Data inválida!!!", 48, Titulo
'            VerDataHora = False
'            SendKeys "+{Home}"
'            formControl.SetFocus
'        End If
'    End If
'
'End Function

Public Sub Habilita(frm As Form)
 Dim i
 For i = 0 To frm.Controls.Count - 1
    If TypeOf frm.Controls(i) Is TextBox Then
       frm.Controls(i).Enabled = True
    End If
    'If TypeOf frm.Controls(i) Is MaskEdBox Then
    '   frm.Controls(i).Enabled = True
    'End If
    If TypeOf frm.Controls(i) Is MSFlexGrid Then
          frm.Controls(i).Enabled = True
       End If
     If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Enabled = True
     End If
     If TypeOf frm.Controls(i) Is OptionButton Then
          frm.Controls(i).Enabled = True
     End If
     If TypeOf frm.Controls(i) Is CheckBox Then
          frm.Controls(i).Enabled = True
       End If
   
Next i

End Sub
Public Sub Desabilita(frm As Form)
'Deixa os textbox desabilitados
   Dim i
   
   For i = 0 To frm.Controls.Count - 1
       If TypeOf frm.Controls(i) Is TextBox Then
          frm.Controls(i).Enabled = False
       End If

       If TypeOf frm.Controls(i) Is MSFlexGrid Then
          frm.Controls(i).Enabled = True
       End If
       If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Enabled = False
       End If
       If TypeOf frm.Controls(i) Is OptionButton Then
          frm.Controls(i).Enabled = False
       End If
       If TypeOf frm.Controls(i) Is CheckBox Then
          frm.Controls(i).Enabled = False
       End If
   Next i
   
End Sub
Public Sub limpa_tela(frm As Form)
 Dim i
 For i = 0 To frm.Controls.Count - 1
    If TypeOf frm.Controls(i) Is TextBox Then
       frm.Controls(i).Enabled = True
       frm.Controls(i).text = ""
    End If
    'If TypeOf frm.Controls(i) Is MaskEdBox Then
    '   frm.Controls(i).Enabled = True
    '   frm.Controls(i).Text = ""
    'End If
    ' If TypeOf frm.Controls(i) Is MaskEdBox Then
    '   frm.Controls(i).Enabled = True
    '   frm.Controls(i).Text = ""
    'End If
    If TypeOf frm.Controls(i) Is MSFlexGrid Then
          frm.Controls(i).Enabled = True
       End If
     If TypeOf frm.Controls(i) Is ComboBox Then
          frm.Controls(i).Enabled = True
     End If
Next i
   
End Sub

Public Sub suCmdAdd(frm As Form)
   frm.cmdUpdate.Enabled = True
   frm.cmdDesfaz.Enabled = True
   frm.cmdEditar.Enabled = False
   frm.cmdAdd.Enabled = False
   frm.cmdSair.Enabled = False
   frm.cmdDelete.Enabled = False

End Sub
Public Sub suCmdDesfaz(frm As Form)
  frm.cmdUpdate.Enabled = False
  frm.cmdDesfaz.Enabled = False
  frm.cmdEditar.Enabled = True
  frm.cmdAdd.Enabled = True
  frm.cmdSair.Enabled = True
  frm.cmdDelete.Enabled = True
End Sub

Public Sub suCmdEditar(frm As Form)
   frm.cmdUpdate.Enabled = True
   frm.cmdDesfaz.Enabled = True
   frm.cmdEditar.Enabled = False
   frm.cmdAdd.Enabled = False
   frm.cmdSair.Enabled = False
   frm.cmdDelete.Enabled = False

End Sub
Public Sub suCmdUpdate(frm As Form)
   frm.cmdUpdate.Enabled = False
   frm.cmdDesfaz.Enabled = False
   frm.cmdEditar.Enabled = True
   frm.cmdAdd.Enabled = True
   frm.cmdSair.Enabled = True
   frm.cmdDelete.Enabled = True

End Sub

Public Sub Desabilita_menu(frm As Form)
frm.MnArquivos.Enabled = False
frm.mnMovimenta.Enabled = False
frm.mnObras.Enabled = False
frm.mncontas.Enabled = False
frm.mnRelato.Enabled = False
frm.mnutilitarios.Enabled = False
frm.mnuHelp.Enabled = False

End Sub
Public Sub Habilita_menu(frm As Form)
frm.MnArquivos.Enabled = True
frm.mnMovimenta.Enabled = True
frm.mnObras.Enabled = True
frm.mncontas.Enabled = True
frm.mnRelato.Enabled = True
frm.mnutilitarios.Enabled = True
frm.mnuHelp.Enabled = True

End Sub
Public Function CalculaCGC(Numero As String) As String

Dim i As Integer
Dim prod As Integer
Dim mult As Integer
Dim digito As Integer

If Not IsNumeric(Numero) Then
   CalculaCGC = ""
   Exit Function
End If

mult = 2
For i = Len(Numero) To 1 Step -1
  prod = prod + Val(Mid(Numero, i, 1)) * mult
  mult = IIf(mult = 9, 2, mult + 1)
Next

digito = 11 - Int(prod Mod 11)
digito = IIf(digito = 10 Or digito = 11, 0, digito)

CalculaCGC = Trim(Str(digito))

End Function
Public Function ValidaCGC(CGC As String) As Boolean
If CalculaCGC(Left(CGC, 12)) <> Mid(CGC, 13, 1) Then
   ValidaCGC = False
   Exit Function
End If

If CalculaCGC(Left(CGC, 13)) <> Mid(CGC, 14, 1) Then
   ValidaCGC = False
   Exit Function
End If

ValidaCGC = True

End Function

Function calculacpf(CPF As String) As Boolean
'Esta rotina foi adaptada da revista Fórum Access
On Error GoTo Err_CPF
Dim i As Integer 'utilizada nos FOR... NEXT
Dim strcampo As String 'armazena do CPF que será utilizada para o cálculo
Dim strCaracter As String 'armazena os digitos do CPF da direita para a esquerda
Dim intNumero As Integer 'armazena o digito separado para cálculo (uma a um)
Dim intMais As Integer 'armazena o digito específico multiplicado pela sua base
Dim lngSoma As Long 'armazena a soma dos digitos multiplicados pela sua base(intmais)
Dim dblDivisao As Double 'armazena a divisão dos digitos*base por 11
Dim lngInteiro As Long 'armazena inteiro da divisão
Dim intResto As Integer 'armazena o resto
Dim intDig1 As Integer 'armazena o 1º digito verificador
Dim intDig2 As Integer 'armazena o 2º digito verificador
Dim strConf As String 'armazena o digito verificador

lngSoma = 0
intNumero = 0
intMais = 0
strcampo = Left(CPF, 9)

'Inicia cálculos do 1º dígito
For i = 2 To 10
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11

lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig1 = 0
Else
    intDig1 = 11 - intResto
End If

strcampo = strcampo & intDig1 'concatena o CPF com o primeiro digito verificador
lngSoma = 0
intNumero = 0
intMais = 0
'Inicia cálculos do 2º dígito
For i = 2 To 11
    strCaracter = Right(strcampo, i - 1)
    intNumero = Left(strCaracter, 1)
    intMais = intNumero * i
    lngSoma = lngSoma + intMais
Next i
dblDivisao = lngSoma / 11
lngInteiro = Int(dblDivisao) * 11
intResto = lngSoma - lngInteiro
If intResto = 0 Or intResto = 1 Then
    intDig2 = 0
Else
    intDig2 = 11 - intResto
End If
strConf = intDig1 & intDig2
'Caso o CPF esteja errado dispara a mensagem
If strConf <> Right(CPF, 2) Then
    calculacpf = False
Else
    calculacpf = True
End If
Exit Function

Exit_CPF:
    Exit Function
Err_CPF:
    MsgBox Error$
    Resume Exit_CPF
End Function

'Public Function ChkData(data As String)
'If data = "" Then
'   ChkData = True
'Else
'   If InStr(data, "/") = 3 And InStrRev(data, "/") = 6 And Len(data) = 10 Then
'      If Not IsDate(data) Then
'         MsgBox "Data Inválida ", vbOKOnly, " Atenção " & gOperador
'         ChkData = False
'      Else
'         If Year(CDate(data)) < 2000 Or Year(CDate(data)) > 2060 Then
'            MsgBox "Data Inválida ", vbOKOnly, " Atenção " & gOperador
'            ChkData = False
'         Else
'            ChkData = True
'         End If
'      End If
'   Else
'      MsgBox "Data Inválida ", vbOKOnly, " Atenção " & gOperador
'      ChkData = False
'   End If
'End If
'End Function

Public Sub Centra(frm As Form)

   'Centraliza a tela no video
   frm.Move (Screen.Width - frm.Width) / 2, _
           (Screen.Height - frm.Height) / 2
   
End Sub


'No nosso caso para Ler os valores do arquivo SHOW.INI usamos o seguinte código:
'valortempo = ReadINI("Geral", "Tempo", App.Path & "\show.ini")
'valorajuda = ReadINI("Geral", "Ajuda", App.Path & "\show.ini")
'atualizaperguntas = ReadINI("Geral", "Atualiza", App.Path & "\show.ini")
' *** arquivo SHOW.INI (Show do Zecão) para guardar algumas preferências do usuário.
'Sua estrutura é a seguinte:

'[Geral]
'Tempo = 50
'Ajuda = 2
'Atualiza = SIM

Public Function ReadINI(Secao As String, Entrada As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
 Dim retlen As String
 Dim Ret As String
 Ret = String$(255, 0)
 retlen = GetPrivateProfileString(Secao, Entrada, "", Ret, Len(Ret), Arquivo)
 Ret = Left$(Ret, retlen)
 ReadINI = Ret
End Function

'2-) A função - WriteINI - escreve em um arquivo INI.
'    Precisa de quatro parâmetros : o nome da Seção , o nome da Entrada ,
'                                   o nome do Texto ( Valor ) e o nome do arquivo INI.
Public Sub WriteINI(Secao As String, Entrada As String, texto As String, Arquivo As String)
  'Arquivo=nome do arquivo ini
  'Secao=O que esta entre []
  'Entrada=nome do que se encontra antes do sinal de igual
  'texto= valor que vem depois do igual
  WritePrivateProfileString Secao, Entrada, texto, Arquivo
End Sub
Public Sub SendKeys(text As String, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.SendKeys text, wait
    Set WshShell = Nothing
End Sub

