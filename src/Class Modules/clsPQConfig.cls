Option Explicit

' =========================================================================================
' CLASSE: clsPQConfig
' OBJETIVO: Framework de gerenciamento de estado e variáveis no Power Query via código M.
' =========================================================================================

' =========================================================================================
' 1. DECLARAÇÕES E VARIÁVEIS DE ESTADO
' =========================================================================================
Private p_QueryName As String
Private p_wb As Workbook

Public VisualizarComoRegistro As Boolean ' Se True, monta o bloco [chave=chave] no final



' =========================================================================================
' 2. CONSTRUTOR
' =========================================================================================
Private Sub Class_Initialize()
    Set p_wb = ActiveWorkbook
    p_QueryName = "Variaveis" ' Define um nome padrão de fábrica
    VisualizarComoRegistro = True
End Sub



' =========================================================================================
' 3. PROPRIEDADES PÚBLICAS
' =========================================================================================
Public Property Get NomeConsulta() As String
    NomeConsulta = p_QueryName
End Property

Public Property Let NomeConsulta(ByVal vNewValue As String)
    p_QueryName = vNewValue
End Property

Public Property Get Existe() As Boolean
    Dim qry As WorkbookQuery
    On Error Resume Next
    Set qry = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    Existe = Not (qry Is Nothing)
End Property

Public Property Get CodigoMAtual() As String
    If Me.Existe Then
        CodigoMAtual = p_wb.Queries(p_QueryName).Formula
    Else
        CodigoMAtual = ""
    End If
End Property



' =========================================================================================
' 4. INFRAESTRUTURA E ESTADO
' =========================================================================================
Public Sub Inicializar()
    Dim codigoBase As String
    If Me.Existe Then Exit Sub
    
    codigoBase = "let" & vbCrLf & _
                 "    _Iniciado = true" & vbCrLf & _
                 "in" & vbCrLf & _
                 "    _Iniciado"
                 
    p_wb.Queries.Add Name:=p_QueryName, Formula:=codigoBase
End Sub

Public Sub LimparTudo()
    Dim strM As String
    If Not Me.Existe Then Exit Sub
    
    strM = "let" & vbCrLf & _
           "    _Iniciado = true" & vbCrLf & _
           "in" & vbCrLf & _
           "    _Iniciado"
           
    AtualizarBlocoIn strM
    p_wb.Queries(p_QueryName).Formula = strM
    Debug.Print "Consulta [" & p_QueryName & "] resetada com sucesso."
End Sub

Public Function ExisteVariavel(ByVal chave As String) As Boolean
    Dim strM As String, regex As Object
    If Not Me.Existe Then
        ExisteVariavel = False
        Exit Function
    End If
    
    strM = Me.CodigoMAtual
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & chave & "\b\s*="
    End With
    
    ExisteVariavel = regex.Test(strM)
    Set regex = Nothing
End Function



' =========================================================================================
' 5. CRUD: LEITURA E EXTRAÇÃO
' =========================================================================================
Public Function LerValor(ByVal chave As String) As Variant
    Dim strM As String, regex As Object, matches As Object
    strM = Me.CodigoMAtual
    If strM = "" Then Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & chave & "\b\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(strM) Then
        Set matches = regex.Execute(strM)
        LerValor = ParseValorM(Trim(matches(0).SubMatches(0)))
    Else
        LerValor = Empty
    End If
    Set regex = Nothing
End Function

Public Function ListarTodas(Optional ByVal DebugMode As Boolean = False) As Object
    Dim strM As String, apenasLet As String
    Dim regex As Object, matches As Object, match As Object
    Dim dictRetorno As Object
    Dim chave As String, Valor As Variant
    
    Set dictRetorno = CreateObject("Scripting.Dictionary")
    strM = Me.CodigoMAtual
    If strM = "" Then GoTo Fim
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
        apenasLet = .Replace(strM, "")
        
        .Global = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        If DebugMode Then Debug.Print "--- VARIÁVEIS NO POWER QUERY [" & p_QueryName & "] ---"
        
        For Each match In matches
            chave = Trim(match.SubMatches(0))
            Valor = ParseValorM(Trim(match.SubMatches(1)))
            dictRetorno.Add chave, Valor
            If DebugMode Then Debug.Print "- " & chave & " (" & TypeName(Valor) & ") = " & CStr(Valor)
        Next match
        
        If DebugMode Then Debug.Print "---------------------------------------------------"
    End If
Fim:
    Set ListarTodas = dictRetorno
    Set regex = Nothing
End Function



' =========================================================================================
' 6. CRUD: ESCRITA E EXCLUSÃO
' =========================================================================================
Public Sub DefinirValor(ByVal chave As String, ByVal Valor As Variant)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add chave, Valor
    Me.DefinirVarios dict
End Sub

Public Sub DefinirVarios(ByVal DictVariaveis As Object)
    Dim strM As String, valorFormatado As String, novaLinha As String
    Dim regex As Object, chave As Variant
    
    If DictVariaveis Is Nothing Then Exit Sub
    If Not Me.Existe Then Me.Inicializar
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
    End With
    
    For Each chave In DictVariaveis.Keys
        valorFormatado = FormatarValorM(DictVariaveis(chave))
        regex.Pattern = "^([ \t]*\b" & CStr(chave) & "\b\s*=)[^\r\n]*"
        
        If regex.Test(strM) Then
            strM = regex.Replace(strM, "$1 " & valorFormatado & ",")
        Else
            novaLinha = "let" & vbCrLf & "    " & CStr(chave) & " = " & valorFormatado & ","
            regex.Pattern = "^let\b"
            strM = regex.Replace(strM, novaLinha)
        End If
    Next chave
    
    regex.Pattern = ",(\s*\bin\b)"
    strM = regex.Replace(strM, "$1")
    
    AtualizarBlocoIn strM
    p_wb.Queries(p_QueryName).Formula = strM
    Set regex = Nothing
End Sub

Public Sub RemoverVariavel(ByVal chave As String)
    Dim strM As String, regex As Object
    If Not Me.Existe Then Exit Sub
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & chave & "\b\s*=[^\r\n]*\r?\n?"
    End With
    
    If regex.Test(strM) Then
        strM = regex.Replace(strM, "")
        regex.Pattern = ",(\s*\bin\b)"
        strM = regex.Replace(strM, "$1")
        
        AtualizarBlocoIn strM
        p_wb.Queries(p_QueryName).Formula = strM
    End If
    Set regex = Nothing
End Sub



' =========================================================================================
' 7. UTILITÁRIOS (Snapshots e Ambientes)
' =========================================================================================
Public Function Clonar(ByVal NomeOrigem As String, ByVal NomeNovo As String, Optional ByVal Sobrescrever As Boolean = False) As Boolean
    Dim qryOrigem As WorkbookQuery, qryDestino As WorkbookQuery
    Dim codigoM As String
    
    On Error Resume Next
    Set qryOrigem = p_wb.Queries(NomeOrigem)
    On Error GoTo 0
    
    If qryOrigem Is Nothing Then
        Clonar = False
        Exit Function
    End If
    
    codigoM = qryOrigem.Formula
    
    On Error Resume Next
    Set qryDestino = p_wb.Queries(NomeNovo)
    On Error GoTo 0
    
    If Not qryDestino Is Nothing Then
        If Sobrescrever Then
            qryDestino.Formula = codigoM
            Clonar = True
        Else
            Err.Raise vbObjectError + 2, "clsPQConfig", "A consulta '" & NomeNovo & "' já existe. Use Sobrescrever:=True."
            Clonar = False
        End If
    Else
        p_wb.Queries.Add Name:=NomeNovo, Formula:=codigoM
        Clonar = True
    End If
End Function



' =========================================================================================
' 8. MOTORES INTERNOS E PRIVADOS (Não chamar externamente)
' =========================================================================================
Private Sub AtualizarBlocoIn(ByRef strM As String)
    If Not Me.VisualizarComoRegistro Then Exit Sub
    
    Dim regex As Object, matches As Object, match As Object
    Dim chaves As String, chave As String, apenasLet As String
    
    Set regex = CreateObject("VBScript.RegExp")
    
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
    End With
    apenasLet = regex.Replace(strM, "")
    
    regex.Multiline = False
    regex.Pattern = "\s+$"
    apenasLet = regex.Replace(apenasLet, "")
    
    With regex
        .Global = True: .Multiline = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*="
    End With
    
    If regex.Test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        chaves = "    [" & vbCrLf
        
        For Each match In matches
            chave = match.SubMatches(0)
            chaves = chaves & "        " & chave & " = " & chave & "," & vbCrLf
        Next match
        
        If Len(chaves) > 3 Then chaves = Left(chaves, Len(chaves) - 3) & vbCrLf
        chaves = chaves & "    ]"
        strM = apenasLet & vbCrLf & "in" & vbCrLf & chaves
    End If
    Set regex = Nothing
End Sub

Private Function ParseValorM(ByVal valorBruto As String) As Variant
    Dim miolo As String, arrDate() As String
    If Right(valorBruto, 1) = "," Then valorBruto = Trim(Left(valorBruto, Len(valorBruto) - 1))
    
    If Left(valorBruto, 10) = "#datetime(" Then
        miolo = Mid(valorBruto, 11, Len(valorBruto) - 11)
        arrDate = Split(miolo, ",")
        ParseValorM = DateSerial(Trim(arrDate(0)), Trim(arrDate(1)), Trim(arrDate(2))) + _
                      TimeSerial(Trim(arrDate(3)), Trim(arrDate(4)), Trim(arrDate(5)))
    ElseIf Left(valorBruto, 6) = "#date(" Then
        miolo = Mid(valorBruto, 7, Len(valorBruto) - 7)
        arrDate = Split(miolo, ",")
        ParseValorM = DateSerial(Trim(arrDate(0)), Trim(arrDate(1)), Trim(arrDate(2)))
    ElseIf Left(valorBruto, 1) = """" And Right(valorBruto, 1) = """" Then
        ParseValorM = Mid(valorBruto, 2, Len(valorBruto) - 2)
    ElseIf LCase(valorBruto) = "true" Then ParseValorM = True
    ElseIf LCase(valorBruto) = "false" Then ParseValorM = False
    ElseIf IsNumeric(valorBruto) Then
        ParseValorM = CDbl(Replace(valorBruto, ".", ","))
    Else
        ParseValorM = valorBruto
    End If
End Function

Private Function FormatarValorM(ByVal Valor As Variant) As String
    If IsObject(Valor) Then
        If TypeName(Valor) = "Range" Then
            FormatarValorM = FormatarValorM(Valor.Value2)
            Exit Function
        Else
            FormatarValorM = "null"
            Exit Function
        End If
    End If

    If IsArray(Valor) Then
        Dim item As Variant, listaM As String
        listaM = "{"
        For Each item In Valor
            If Not IsEmpty(item) Then listaM = listaM & FormatarValorM(item) & ", "
        Next item
        If Len(listaM) > 2 Then listaM = Left(listaM, Len(listaM) - 2)
        FormatarValorM = listaM & "}"
        Exit Function
    End If

    Select Case VarType(Valor)
        Case vbString: FormatarValorM = """" & Valor & """"
        Case vbBoolean: FormatarValorM = IIf(Valor, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency: FormatarValorM = Replace(CStr(Valor), ",", ".")
        Case vbDate: FormatarValorM = "#datetime(" & Year(Valor) & ", " & Month(Valor) & ", " & Day(Valor) & ", " & Hour(Valor) & ", " & Minute(Valor) & ", " & Second(Valor) & ")"
        Case vbNull, vbEmpty: FormatarValorM = "null"
        Case Else: FormatarValorM = """" & CStr(Valor) & """"
    End Select
End Function


' =========================================================================================
' 9. RAD TOOLS (Importação e Exportação de Componentes)
' =========================================================================================
Public Sub ExportarConsultaParaTXT(ByVal NomeConsulta As String, ByVal CaminhoArquivo As String)
    Dim qry As WorkbookQuery
    Dim stream As Object
    Dim descCompleta As String, descParte As String, comoUsarParte As String
    Dim delimitadorComoUsar As String
    Dim conteudoTXT As String
    
    ' Verifica se a consulta existe
    On Error Resume Next
    Set qry = p_wb.Queries(NomeConsulta)
    On Error GoTo 0
    
    If qry Is Nothing Then
        Err.Raise vbObjectError + 3, "clsPQConfig", "A consulta '" & NomeConsulta & "' não existe."
        Exit Sub
    End If
    
    ' Separa a Descrição do "Como Usar"
    descCompleta = qry.Description
    delimitadorComoUsar = "[COMO USAR]"
    
    If InStr(1, descCompleta, delimitadorComoUsar, vbTextCompare) > 0 Then
        descParte = Trim(Split(descCompleta, delimitadorComoUsar)(0))
        comoUsarParte = Trim(Split(descCompleta, delimitadorComoUsar)(1))
    Else
        descParte = descCompleta
        comoUsarParte = ""
    End If
    
    ' Monta o texto estruturado usando tags de proteção
    conteudoTXT = "<NOME>" & vbCrLf & qry.Name & vbCrLf & "</NOME>" & vbCrLf & _
                  "<DESCRICAO>" & vbCrLf & descParte & vbCrLf & "</DESCRICAO>" & vbCrLf & _
                  "<COMOUSAR>" & vbCrLf & comoUsarParte & vbCrLf & "</COMOUSAR>" & vbCrLf & _
                  "<CODIGOM>" & vbCrLf & qry.Formula & vbCrLf & "</CODIGOM>"
                  
    ' Salva o arquivo em UTF-8 via ADODB.Stream (Garante que os acentos do Português não quebrem)
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' adTypeText
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText conteudoTXT
    stream.SaveToFile CaminhoArquivo, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub


Public Sub ImportarConsultaDeTXT(ByVal CaminhoArquivo As String, Optional ByVal Sobrescrever As Boolean = False)
    Dim stream As Object
    Dim conteudoTXT As String
    Dim strNome As String, strDesc As String, strComoUsar As String, strCodigoM As String
    Dim descFinal As String
    Dim qry As WorkbookQuery
    
    ' Lê o arquivo UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile CaminhoArquivo
    conteudoTXT = stream.ReadText
    stream.Close
    
    ' Extrai as tags usando a nossa função auxiliar Privada
    strNome = ExtrairTag(conteudoTXT, "NOME")
    strDesc = ExtrairTag(conteudoTXT, "DESCRICAO")
    strComoUsar = ExtrairTag(conteudoTXT, "COMOUSAR")
    strCodigoM = ExtrairTag(conteudoTXT, "CODIGOM")
    
    ' Validação de segurança do pacote
    If strNome = "" Or strCodigoM = "" Then
        Err.Raise vbObjectError + 4, "clsPQConfig", "Arquivo corrompido ou fora do padrão RAD."
        Exit Sub
    End If
    
    ' Une a descrição e o "como usar" para caber nativamente no Power Query
    descFinal = strDesc
    If strComoUsar <> "" Then
        descFinal = descFinal & vbCrLf & vbCrLf & "[COMO USAR]" & vbCrLf & strComoUsar
    End If
    
    ' Cria ou Atualiza a consulta no Excel
    On Error Resume Next
    Set qry = p_wb.Queries(strNome)
    On Error GoTo 0
    
    If Not qry Is Nothing Then
        If Sobrescrever Then
            qry.Formula = strCodigoM
            qry.Description = descFinal
        Else
            Err.Raise vbObjectError + 5, "clsPQConfig", "O componente '" & strNome & "' já existe. Use Sobrescrever:=True."
        End If
    Else
        ' Injeta o nome, o código M e os metadados!
        Set qry = p_wb.Queries.Add(Name:=strNome, Formula:=strCodigoM, Description:=descFinal)
    End If
End Sub


Private Function ExtrairTag(ByVal Texto As String, ByVal TagName As String) As String
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = False
    regex.IgnoreCase = True
    regex.Multiline = True
    ' Regex que lê exatamente o que está entre <TAG> e </TAG>
    regex.Pattern = "<" & TagName & ">\r?\n?([\s\S]*?)\r?\n?</" & TagName & ">"
    
    If regex.Test(Texto) Then
        Set matches = regex.Execute(Texto)
        ExtrairTag = Trim(matches(0).SubMatches(0))
    Else
        ExtrairTag = ""
    End If
    
    Set regex = Nothing
End Function

