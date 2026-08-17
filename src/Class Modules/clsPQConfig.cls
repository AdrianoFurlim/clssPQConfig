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
' ==========================================================================
' Rotina:      Class_Initialize
' Descrição:   Construtor padrão da classe. Inicializa o estado definindo o
'              ambiente de trabalho e os parâmetros padrão da consulta.
' Parâmetros:  Nenhum
' ==========================================================================
Private Sub Class_Initialize()
    ' Define o arquivo atual como base para buscar as consultas
    Set p_wb = ActiveWorkbook
    p_QueryName = "Variaveis" ' Define um nome padrão de fábrica
    ' Garante que o output final do M seja estruturado em formato de registro
    VisualizarComoRegistro = True
End Sub



' =========================================================================================
' 3. PROPRIEDADES PÚBLICAS
' =========================================================================================
' ==========================================================================
' Propriedade: NomeConsulta (Get)
' Descrição:   Retorna o nome da consulta alvo no Power Query.
' Retorno:     String - O nome atual da consulta.
' ==========================================================================
Public Property Get NomeConsulta() As String
    NomeConsulta = p_QueryName
End Property

' ==========================================================================
' Propriedade: NomeConsulta (Let)
' Descrição:   Permite alterar dinamicamente o nome da consulta alvo.
' Parâmetros:  vNewValue (String) - O novo nome a ser atribuído.
' ==========================================================================
Public Property Let NomeConsulta(ByVal vNewValue As String)
    p_QueryName = vNewValue
End Property

' ==========================================================================
' Propriedade: Existe (Get)
' Descrição:   Valida se a consulta definida em p_QueryName existe no Excel.
' Retorno:     Boolean - True se a consulta for encontrada, False caso contrário.
' ==========================================================================
Public Property Get Existe() As Boolean
    Dim qry As WorkbookQuery
    ' Desativa a interrupção de erro temporariamente para testar a atribuição do objeto
    On Error Resume Next
    Set qry = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    ' Se a variável de objeto não estiver vazia, significa que a consulta foi encontrada
    Existe = Not (qry Is Nothing)
End Property

' ==========================================================================
' Propriedade: CodigoMAtual (Get)
' Descrição:   Recupera a sintaxe M (fórmula) atual da consulta, se existir.
' Retorno:     String - O código M da consulta ou uma string vazia.
' ==========================================================================
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
' ==========================================================================
' Rotina:      Inicializar
' Descrição:   Cria a consulta base no Power Query se ela ainda não existir,
'              fornecendo um esqueleto M inicial e válido.
' Parâmetros:  Nenhum
' ==========================================================================
Public Sub Inicializar()
    Dim codigoBase As String
    If Me.Existe Then Exit Sub
    
    ' Constrói a estrutura let...in mínima necessária para o Power Query reconhecer a consulta
    codigoBase = "let" & vbCrLf & _
                 "    _Iniciado = true" & vbCrLf & _
                 "in" & vbCrLf & _
                 "    _Iniciado"
                 
    p_wb.Queries.Add Name:=p_QueryName, Formula:=codigoBase
End Sub

' ==========================================================================
' Rotina:      LimparTudo
' Descrição:   Reseta a consulta existente, sobrescrevendo todo o código M
'              com o esqueleto base inicial.
' Parâmetros:  Nenhum
' ==========================================================================
Public Sub LimparTudo()
    Dim strM As String
    If Not Me.Existe Then Exit Sub
    
    ' Define o estado virgem do código M
    strM = "let" & vbCrLf & _
           "    _Iniciado = true" & vbCrLf & _
           "in" & vbCrLf & _
           "    _Iniciado"
           
    ' Aplica as formatações de registro (se ativadas) e atualiza a consulta no Excel
    AtualizarBlocoIn strM
    p_wb.Queries(p_QueryName).Formula = strM
    Debug.Print "Consulta [" & p_QueryName & "] resetada com sucesso."
End Sub

' ==========================================================================
' Função:      ExisteVariavel
' Descrição:   Verifica via Expressão Regular se uma variável específica
'              está declarada dentro do código M da consulta.
' Parâmetros:  chave (String) - O nome da variável a ser procurada.
' Retorno:     Boolean - True se a variável existir.
' ==========================================================================
Public Function ExisteVariavel(ByVal chave As String) As Boolean
    Dim strM As String, regex As Object
    If Not Me.Existe Then
        ExisteVariavel = False
        Exit Function
    End If
    
    strM = Me.CodigoMAtual
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Configura o Regex para buscar a chave exatamente como um identificador no início da linha
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
' ==========================================================================
' Função:      LerValor
' Descrição:   Extrai e interpreta o valor de uma variável específica no
'              código M, convertendo para um tipo de dado VBA apropriado.
' Parâmetros:  chave (String) - A variável a ser lida.
' Retorno:     Variant - O valor lido e tipado, ou Empty se não encontrar.
' ==========================================================================
Public Function LerValor(ByVal chave As String) As Variant
    Dim strM As String, regex As Object, matches As Object
    strM = Me.CodigoMAtual
    If strM = "" Then Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    ' Expressão para capturar tudo que estiver após o sinal de igualdade na mesma linha
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "^[ \t]*\b" & chave & "\b\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(strM) Then
        Set matches = regex.Execute(strM)
        ' Envia o conteúdo capturado para o motor de conversão M -> VBA
        LerValor = ParseValorM(Trim(matches(0).SubMatches(0)))
    Else
        LerValor = Empty
    End If
    Set regex = Nothing
End Function

' ==========================================================================
' Função:      ListarTodas
' Descrição:   Mapeia e retorna todas as variáveis e valores contidos na
'              consulta M em um formato de dicionário VBA.
' Parâmetros:  DebugMode (Boolean) - Se True, imprime os resultados na Janela de Verificação.
' Retorno:     Object (Scripting.Dictionary) - Dicionário com chaves e valores tipados.
' ==========================================================================
Public Function ListarTodas(Optional ByVal DebugMode As Boolean = False) As Object
    Dim strM As String, apenasLet As String
    Dim regex As Object, matches As Object, match As Object
    Dim dictRetorno As Object
    Dim chave As String, Valor As Variant
    
    Set dictRetorno = CreateObject("Scripting.Dictionary")
    strM = Me.CodigoMAtual
    If strM = "" Then GoTo Fim
    
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Isola o bloco "let" removendo todo o bloco "in" para não confundir a extração
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
        
        ' Varre todas as correspondências e converte os tipos de M para VBA
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
' ==========================================================================
' Rotina:      DefinirValor
' Descrição:   Atualiza ou insere uma única variável e seu valor no código M.
' Parâmetros:  chave (String) - Nome da variável.
'              Valor (Variant) - Valor a ser inserido.
' ==========================================================================
Public Sub DefinirValor(ByVal chave As String, ByVal Valor As Variant)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add chave, Valor
    ' Encaminha para o motor em lote (DefinirVarios) para evitar duplicação de lógica
    Me.DefinirVarios dict
End Sub

' ==========================================================================
' Rotina:      DefinirVarios
' Descrição:   Insere ou atualiza um lote de variáveis fornecidas através de
'              um objeto Dictionary, diretamente no código M.
' Parâmetros:  DictVariaveis (Object) - Dicionário contendo chaves e valores.
' ==========================================================================
Public Sub DefinirVarios(ByVal DictVariaveis As Object)
    Dim strM As String, valorFormatado As String, novaLinha As String
    Dim regex As Object, chave As Variant
    
    If DictVariaveis Is Nothing Then Exit Sub
    ' Garante a existência do script base antes de tentar manipulação de strings
    If Not Me.Existe Then Me.Inicializar
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
    End With
    
    For Each chave In DictVariaveis.Keys
        ' Transforma o tipo VBA nativo para a string correspondente no formato M
        valorFormatado = FormatarValorM(DictVariaveis(chave))
        regex.Pattern = "^([ \t]*\b" & CStr(chave) & "\b\s*=)[^\r\n]*"
        
        ' Se a variável já existir, faz a substituição do valor na mesma linha
        If regex.Test(strM) Then
            strM = regex.Replace(strM, "$1 " & valorFormatado & ",")
        Else
            ' Se não existir, injeta uma nova linha de declaração logo após o 'let'
            novaLinha = "let" & vbCrLf & "    " & CStr(chave) & " = " & valorFormatado & ","
            regex.Pattern = "^let\b"
            strM = regex.Replace(strM, novaLinha)
        End If
    Next chave
    
    ' Remove a vírgula sobressalente da última variável antes da palavra 'in'
    regex.Pattern = ",(\s*\bin\b)"
    strM = regex.Replace(strM, "$1")
    
    ' Recalcula a estrutura do 'in' se necessário e salva a query
    AtualizarBlocoIn strM
    p_wb.Queries(p_QueryName).Formula = strM
    Set regex = Nothing
End Sub

' ==========================================================================
' Rotina:      RemoverVariavel
' Descrição:   Apaga a declaração de uma variável do código M, se ela existir.
' Parâmetros:  chave (String) - Nome da variável a ser excluída.
' ==========================================================================
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
        ' Remove a linha inteira da variável encontrada
        strM = regex.Replace(strM, "")
        
        ' Assegura que a sintaxe continue válida removendo a vírgula antes do 'in'
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
' ==========================================================================
' Função:      Clonar
' Descrição:   Cria uma duplicata exata de uma consulta Power Query existente.
' Parâmetros:  NomeOrigem (String) - Nome da query a ser copiada.
'              NomeNovo (String) - Nome da nova query clonada.
'              Sobrescrever (Boolean) - Se True, substitui a query de destino se já existir.
' Retorno:     Boolean - True se a clonagem foi bem-sucedida.
' ==========================================================================
Public Function Clonar(ByVal NomeOrigem As String, ByVal NomeNovo As String, Optional ByVal Sobrescrever As Boolean = False) As Boolean
    Dim qryOrigem As WorkbookQuery, qryDestino As WorkbookQuery
    Dim codigoM As String
    
    ' Valida a existência da query de origem com segurança
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
        ' Controla o comportamento caso a query de destino já exista para evitar erros do Excel
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
' ==========================================================================
' Rotina:      AtualizarBlocoIn (Private)
' Descrição:   Reconstrói dinamicamente o bloco 'in' do código M para expor
'              todas as variáveis como um Registro, facilitando o debug.
' Parâmetros:  strM (String, ByRef) - O código M sendo manipulado em memória.
' ==========================================================================
Private Sub AtualizarBlocoIn(ByRef strM As String)
    If Not Me.VisualizarComoRegistro Then Exit Sub
    
    Dim regex As Object, matches As Object, match As Object
    Dim chaves As String, chave As String, apenasLet As String
    
    Set regex = CreateObject("VBScript.RegExp")
    
    ' Divide o código na palavra 'in'
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
    End With
    apenasLet = regex.Replace(strM, "")
    
    ' Limpa espaços ocos no final do bloco let
    regex.Multiline = False
    regex.Pattern = "\s+$"
    apenasLet = regex.Replace(apenasLet, "")
    
    ' Mapeia todas as variáveis declaradas no bloco let
    With regex
        .Global = True: .Multiline = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*="
    End With
    
    If regex.Test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        chaves = "    [" & vbCrLf
        
        ' Monta a sintaxe de registro no M: [var1 = var1, var2 = var2]
        For Each match In matches
            chave = match.SubMatches(0)
            chaves = chaves & "        " & chave & " = " & chave & "," & vbCrLf
        Next match
        
        ' Remove a vírgula excedente do final da última propriedade do registro
        If Len(chaves) > 3 Then chaves = Left(chaves, Len(chaves) - 3) & vbCrLf
        chaves = chaves & "    ]"
        strM = apenasLet & vbCrLf & "in" & vbCrLf & chaves
    End If
    Set regex = Nothing
End Sub

' ==========================================================================
' Função:      ParseValorM (Private)
' Descrição:   Lê uma representação textual de valor em M e a converte de volta
'              para o tipo de dado original equivalente no ambiente VBA.
' Parâmetros:  valorBruto (String) - O fragmento de string M contendo o valor.
' Retorno:     Variant - Valor processado e tipado para uso no VBA.
' ==========================================================================
Private Function ParseValorM(ByVal valorBruto As String) As Variant
    Dim miolo As String, arrDate() As String
    ' Limpa a vírgula sintática caso tenha vindo na captura da string
    If Right(valorBruto, 1) = "," Then valorBruto = Trim(Left(valorBruto, Len(valorBruto) - 1))
    
    ' Motor de parse condicional baseado no prefixo e formato da string M
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
        ' Garante que o separador decimal seja compreendido pelo VBA regional
        ParseValorM = CDbl(Replace(valorBruto, ".", ","))
    Else
        ParseValorM = valorBruto
    End If
End Function

' ==========================================================================
' Função:      FormatarValorM (Private)
' Descrição:   Analisa um valor de tipo VBA e cria a string literal M
'              necessária para injetá-lo no Power Query.
' Parâmetros:  Valor (Variant) - O dado VBA a ser transformado.
' Retorno:     String - Fragmento de código M que representa o valor original.
' ==========================================================================
Private Function FormatarValorM(ByVal Valor As Variant) As String
    ' Trata objetos complexos extraindo os seus valores subjacentes (ex: Ranges)
    If IsObject(Valor) Then
        If TypeName(Valor) = "Range" Then
            FormatarValorM = FormatarValorM(Valor.Value2)
            Exit Function
        Else
            FormatarValorM = "null"
            Exit Function
        End If
    End If

    ' Trata matrizes transformando-as em listas de valores M (ex: {1, 2, "a"})
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

    ' Converte valores escalares identificando o VarType e gerando a sintaxe estrita em M
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
' ==========================================================================
' Rotina:      ExportarConsultaParaTXT
' Descrição:   Extrai os metadados (Descrição e Instruções) e o código M
'              da consulta para um arquivo de texto com marcação (Tags).
' Parâmetros:  NomeConsulta (String) - A query alvo da extração.
'              CaminhoArquivo (String) - Destino físico do arquivo .txt.
' ==========================================================================
Public Sub ExportarConsultaParaTXT(ByVal NomeConsulta As String, ByVal caminhoArquivo As String)
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
    
    ' Separa a Descrição do "Como Usar" com base em uma tag lógica dentro das propriedades do PQ
    descCompleta = qry.Description
    delimitadorComoUsar = "[COMO USAR]"
    
    If InStr(1, descCompleta, delimitadorComoUsar, vbTextCompare) > 0 Then
        descParte = Trim(Split(descCompleta, delimitadorComoUsar)(0))
        comoUsarParte = Trim(Split(descCompleta, delimitadorComoUsar)(1))
    Else
        descParte = descCompleta
        comoUsarParte = ""
    End If
    
    ' Monta o texto estruturado usando tags de proteção personalizadas
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
    stream.SaveToFile caminhoArquivo, 2 ' adSaveCreateOverWrite
    stream.Close
End Sub

' ==========================================================================
' Rotina:      ImportarConsultaDeTXT
' Descrição:   Lê um arquivo .txt estruturado via RAD Tools e reconstrói
'              a consulta no Power Query, injetando o código M e os metadados.
' Parâmetros:  CaminhoArquivo (String) - Caminho lógico do arquivo texto.
'              Sobrescrever (Boolean) - Permissão para substituir consulta pré-existente.
' ==========================================================================
Public Sub ImportarConsultaDeTXT(ByVal caminhoArquivo As String, Optional ByVal Sobrescrever As Boolean = False)
    Dim stream As Object
    Dim conteudoTXT As String
    Dim strNome As String, strDesc As String, strComoUsar As String, strCodigoM As String
    Dim descFinal As String
    Dim qry As WorkbookQuery
    
    ' Lê o arquivo UTF-8 para preservar a codificação especial e acentos de metadados
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile caminhoArquivo
    conteudoTXT = stream.ReadText
    stream.Close
    
    ' Extrai as tags usando a nossa função auxiliar Privada
    strNome = ExtrairTag(conteudoTXT, "NOME")
    strDesc = ExtrairTag(conteudoTXT, "DESCRICAO")
    strComoUsar = ExtrairTag(conteudoTXT, "COMOUSAR")
    strCodigoM = ExtrairTag(conteudoTXT, "CODIGOM")
    
    ' Validação de segurança do pacote lido
    If strNome = "" Or strCodigoM = "" Then
        Err.Raise vbObjectError + 4, "clsPQConfig", "Arquivo corrompido ou fora do padrão RAD."
        Exit Sub
    End If
    
    ' Une a descrição e o "como usar" para caber nativamente nas propriedades descritivas do Power Query
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
        ' Injeta o nome, o código M e os metadados simultaneamente em uma nova consulta!
        Set qry = p_wb.Queries.Add(Name:=strNome, Formula:=strCodigoM, Description:=descFinal)
    End If
End Sub

' ==========================================================================
' Rotina:      SalvarTextoBruto
' Descrição:   Armazena um texto limpo (HTML, CSS, JS) diretamente na consulta atual.
'              Pode receber o código via parâmetro ou ler direto de um arquivo físico.
' Parâmetros:  TextoConteudo (String, Opcional) - O código em string a ser salvo.
'              CaminhoArquivo (String, Opcional) - O diretório do arquivo físico.
' ==========================================================================
Public Sub SalvarTextoBruto(Optional ByVal TextoConteudo As String = "", Optional ByVal CaminhoArquivo As String = "")
    Dim strM As String
    Dim qry As WorkbookQuery
    Dim conteudoFinal As String
    Dim stream As Object
    
    ' 1. Valida se o nome da consulta foi definido antes de tentar salvar
    If p_QueryName = "" Then
        Err.Raise vbObjectError + 8, "clsPQConfig", "Defina o NomeConsulta antes de salvar."
        Exit Sub
    End If
    
    ' 2. Define a origem do texto: Arquivo físico ou String direta
    If CaminhoArquivo <> "" Then
        ' Usa ADODB.Stream para garantir a leitura perfeita em UTF-8
        On Error Resume Next
        Set stream = CreateObject("ADODB.Stream")
        stream.Type = 2 ' adTypeText
        stream.Charset = "utf-8"
        stream.Open
        stream.LoadFromFile CaminhoArquivo
        
        If Err.Number <> 0 Then
            Err.Raise vbObjectError + 6, "clsPQConfig", "Não foi possível ler o arquivo: " & CaminhoArquivo
            Exit Sub
        End If
        On Error GoTo 0
        
        conteudoFinal = stream.ReadText
        stream.Close
    ElseIf TextoConteudo <> "" Then
        conteudoFinal = TextoConteudo
    Else
        Err.Raise vbObjectError + 7, "clsPQConfig", "Forneça o TextoConteudo ou o CaminhoArquivo para salvar."
        Exit Sub
    End If
    
    ' 3. Prepara a string para o padrão estrito do Power Query
    strM = """" & Replace(conteudoFinal, """", """""") & """"
    
    ' 4. Valida se a consulta já existe para criar ou atualizar
    On Error Resume Next
    Set qry = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    
    If Not qry Is Nothing Then
        qry.Formula = strM
    Else
        p_wb.Queries.Add Name:=p_QueryName, Formula:=strM
    End If
End Sub

' ==========================================================================
' Função:      LerTextoBruto
' Descrição:   Extrai o texto puro armazenado na consulta atual, removendo
'              o encapsulamento e o escape de aspas duplas do código M.
' Parâmetros:  Nenhum.
' Retorno:     String - O código limpo original (HTML/CSS/JS/etc).
' ==========================================================================
Public Function LerTextoBruto() As String
    Dim strM As String
    Dim qry As WorkbookQuery
    
    On Error Resume Next
    Set qry = p_wb.Queries(p_QueryName)
    On Error GoTo 0
    
    If Not qry Is Nothing Then
        strM = qry.Formula
        
        ' Verifica se o texto está encapsulado entre aspas do formato M
        If Left(strM, 1) = """" And Right(strM, 1) = """" Then
            ' Remove a primeira e a última aspa
            strM = Mid(strM, 2, Len(strM) - 2)
        End If
        
        ' Desfaz o escape de aspas duplas exigido pelo Power Query
        LerTextoBruto = Replace(strM, """""", """")
    Else
        ' Retorna vazio caso a consulta não seja encontrada
        LerTextoBruto = ""
    End If
End Function

' ==========================================================================
' Função:      ExtrairTag (Private)
' Descrição:   Interpreta a sintaxe similar a XML gerada pela classe para
'              isolar o conteúdo encapsulado entre duas tags específicas.
' Parâmetros:  Texto (String) - O conteúdo bruto do arquivo.
'              TagName (String) - Nome da tag base de abertura/fechamento.
' Retorno:     String - O texto contido dentro da tag.
' ==========================================================================
Private Function ExtrairTag(ByVal Texto As String, ByVal TagName As String) As String
    Dim regex As Object, matches As Object
    Set regex = CreateObject("VBScript.RegExp")
    
    regex.Global = False
    regex.IgnoreCase = True
    regex.Multiline = True
    ' Regex que lê exatamente o que está entre <TAG> e </TAG>, suportando quebras de linha
    regex.Pattern = "<" & TagName & ">\r?\n?([\s\S]*?)\r?\n?</" & TagName & ">"
    
    If regex.Test(Texto) Then
        Set matches = regex.Execute(Texto)
        ExtrairTag = Trim(matches(0).SubMatches(0))
    Else
        ExtrairTag = ""
    End If
    
    Set regex = Nothing
End Function

