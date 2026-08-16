Option Explicit

' =========================================================
' CLASSE: clsPQConfig
' OBJETIVO: Gerenciar variáveis em uma consulta do Power Query via código M
' =========================================================

' Variáveis privadas que guardam o estado interno do objeto
Private p_QueryName As String
Private p_wb As Workbook
Public VisualizarComoRegistro As Boolean



' ---------------------------------------------------------
' 1. CONSTRUTOR (Executado ao instanciar a classe)
' ---------------------------------------------------------
Private Sub Class_Initialize()
    Set p_wb = ActiveWorkbook
    p_QueryName = "Variaveis" ' Define um nome padrão de fábrica
    VisualizarComoRegistro = True
End Sub



' ---------------------------------------------------------
' 2. PROPRIEDADES (Características da classe)
' ---------------------------------------------------------
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
    
    ' Retorna True se o objeto qry foi instanciado com sucesso
    Existe = Not (qry Is Nothing)
End Property

Public Property Get CodigoMAtual() As String
    If Me.Existe Then
        CodigoMAtual = p_wb.Queries(p_QueryName).Formula
    Else
        CodigoMAtual = ""
    End If
End Property



' ---------------------------------------------------------
' 3. MÉTODOS DE INFRAESTRUTURA
' ---------------------------------------------------------
Public Sub Inicializar()
    Dim codigoBase As String
    
    ' Se a consulta já existir, aborta para não sobrescrever
    If Me.Existe Then Exit Sub
    
    ' Cria o escopo básico para aceitar variáveis
    codigoBase = "let" & vbCrLf & _
                 "    _Iniciado = true" & vbCrLf & _
                 "in" & vbCrLf & _
                 "    _Iniciado"
                 
    p_wb.Queries.Add Name:=p_QueryName, Formula:=codigoBase
End Sub



' ---------------------------------------------------------
' 4. LER VALOR (Agora com suporte a Datas)
' ---------------------------------------------------------
Public Function LerValor(ByVal chave As String) As Variant
    Dim strM As String, regex As Object, matches As Object
    strM = Me.CodigoMAtual
    If strM = "" Then Exit Function
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        ' Lê até o fim da linha, ignorando vírgulas internas
        .Pattern = "^[ \t]*\b" & chave & "\b\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(strM) Then
        Set matches = regex.Execute(strM)
        ' Envia para o nosso novo tradutor
        LerValor = ParseValorM(Trim(matches(0).SubMatches(0)))
    Else
        LerValor = Empty
    End If
    Set regex = Nothing
End Function



' ---------------------------------------------------------
' 5. MÉTODOS DE DADOS (CRUD - Update & Create)
' ---------------------------------------------------------
Public Sub DefinirValor(ByVal chave As String, ByVal Valor As Variant)
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' Adiciona a única variável no dicionário
    dict.Add chave, Valor
    
    ' Reaproveita o motor blindado do DefinirVarios para fazer o trabalho duro!
    Me.DefinirVarios dict
End Sub



Public Sub RemoverVariavel(ByVal chave As String)
    Dim strM As String
    Dim regex As Object
    
    If Not Me.Existe Then Exit Sub
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        
        ' NOVO PADRÃO: Pega do início da linha até a quebra de linha (\r\n), ignorando vírgulas internas
        .Pattern = "^[ \t]*\b" & chave & "\b\s*=[^\r\n]*\r?\n?"
    End With
    
    If regex.Test(strM) Then
        ' Apaga a variável do bloco let
        strM = regex.Replace(strM, "")
        
        ' Correção de sintaxe crucial: Remove a vírgula orfã
        regex.Pattern = ",(\s*\bin\b)"
        strM = regex.Replace(strM, "$1")
        
        ' Reconstrói o IN
        AtualizarBlocoIn strM
        
        ' Salva no Power Query
        p_wb.Queries(p_QueryName).Formula = strM
    End If
    
    Set regex = Nothing
End Sub



' ---------------------------------------------------------
' 5. DEFINIR VARIOS (Corrige o Bug da vírgula do DateTime)
' ---------------------------------------------------------
Public Sub DefinirVarios(ByVal DictVariaveis As Object)
    Dim strM As String, valorFormatado As String, novaLinha As String
    Dim regex As Object, chave As Variant
    
    If DictVariaveis Is Nothing Then Exit Sub
    If Not Me.Existe Then Me.Inicializar
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
    End With
    
    For Each chave In DictVariaveis.Keys
        valorFormatado = FormatarValorM(DictVariaveis(chave))
        
        ' NOVO PADRÃO: Substitui a linha inteira para não bugar com vírgulas internas
        regex.Pattern = "^([ \t]*\b" & CStr(chave) & "\b\s*=)[^\r\n]*"
        
        If regex.Test(strM) Then
            ' Update (Força a vírgula no final)
            strM = regex.Replace(strM, "$1 " & valorFormatado & ",")
        Else
            ' Create
            novaLinha = "let" & vbCrLf & "    " & CStr(chave) & " = " & valorFormatado & ","
            regex.Pattern = "^let\b"
            strM = regex.Replace(strM, novaLinha)
        End If
    Next chave
    
    ' CORREÇÃO DE SEGURANÇA: Remove qualquer vírgula órfã antes do "in"
    regex.Pattern = ",(\s*\bin\b)"
    strM = regex.Replace(strM, "$1")
    
    AtualizarBlocoIn strM
    p_wb.Queries(p_QueryName).Formula = strM
    Set regex = Nothing
End Sub

' ---------------------------------------------------------
' 6. LISTAR TODAS (Extração e Log)
' ---------------------------------------------------------
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
        ' Isola apenas o bloco "let" para não ler as variáveis do bloco "in"
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
        apenasLet = .Replace(strM, "")
    
        ' Puxa todas as chaves e valores da linha
        .Global = True
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*([^\r\n]*)"
    End With
    
    If regex.Test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        
        If DebugMode Then
            Debug.Print "--- VARIÁVEIS NO POWER QUERY [" & p_QueryName & "] ---"
        End If
        
        For Each match In matches
            chave = Trim(match.SubMatches(0))
            ' Converte de volta para Tipo VBA nativo
            Valor = ParseValorM(Trim(match.SubMatches(1)))
            
            dictRetorno.Add chave, Valor
            
            If DebugMode Then
                Debug.Print "- " & chave & " (" & TypeName(Valor) & ") = " & CStr(Valor)
            End If
        Next match
        
        If DebugMode Then Debug.Print "---------------------------------------------------"
    End If

Fim:
    Set ListarTodas = dictRetorno
    Set regex = Nothing
End Function



' ---------------------------------------------------------
' RECONSTRUTOR DO BLOCO IN (Motor Privado)
' ---------------------------------------------------------
Private Sub AtualizarBlocoIn(ByRef strM As String)
    ' Se o usuário desativou a propriedade, o construtor não faz nada
    If Not Me.VisualizarComoRegistro Then Exit Sub
    
    Dim regex As Object, matches As Object, match As Object
    Dim chaves As String, chave As String
    Dim apenasLet As String
    
    Set regex = CreateObject("VBScript.RegExp")
    
    ' PASSO 1: Isolar apenas o bloco "let" apagando o "in" e tudo o que vem depois
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        .Pattern = "\bin\b[\s\S]*"
    End With
    
    apenasLet = regex.Replace(strM, "")
    
    ' --- O SEGREDO DA CORREÇÃO ESTÁ AQUI ---
    ' Desligamos o MultiLine para que o $ apague até as quebras de linha no final absoluto do texto
    regex.Multiline = False
    regex.Pattern = "\s+$"
    apenasLet = regex.Replace(apenasLet, "")
    ' ---------------------------------------
    
    ' PASSO 2: Buscar as chaves APENAS dentro da parte isolada
    With regex
        .Global = True
        .Multiline = True ' Ligamos de novo para ler os inícios de linha
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*="
    End With
    
    If regex.Test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        chaves = "    [" & vbCrLf
        
        ' Monta o bloco [Chave = Chave]
        For Each match In matches
            chave = match.SubMatches(0)
            chaves = chaves & "        " & chave & " = " & chave & "," & vbCrLf
        Next match
        
        ' Remove a última vírgula e a quebra de linha
        If Len(chaves) > 3 Then
            chaves = Left(chaves, Len(chaves) - 3) & vbCrLf
        End If
        chaves = chaves & "    ]"
        
        ' PASSO 3: Junta o "let", "in", e o novo bloco limpo
        strM = apenasLet & vbCrLf & "in" & vbCrLf & chaves
    End If
    
    Set regex = Nothing
End Sub

' ---------------------------------------------------------
' FUNÇÃO AUXILIAR PRIVADA (M -> VBA)
' ---------------------------------------------------------
Private Function ParseValorM(ByVal valorBruto As String) As Variant
    Dim miolo As String, arrDate() As String
    
    ' Remove a vírgula final se vier junto na captura
    If Right(valorBruto, 1) = "," Then valorBruto = Trim(Left(valorBruto, Len(valorBruto) - 1))
    
    ' Tipagem: Data e Hora
    If Left(valorBruto, 10) = "#datetime(" Then
        miolo = Mid(valorBruto, 11, Len(valorBruto) - 11) ' Pega apenas os números
        arrDate = Split(miolo, ",")
        ParseValorM = DateSerial(Trim(arrDate(0)), Trim(arrDate(1)), Trim(arrDate(2))) + _
                      TimeSerial(Trim(arrDate(3)), Trim(arrDate(4)), Trim(arrDate(5)))
                      
    ' Tipagem: Apenas Data
    ElseIf Left(valorBruto, 6) = "#date(" Then
        miolo = Mid(valorBruto, 7, Len(valorBruto) - 7)
        arrDate = Split(miolo, ",")
        ParseValorM = DateSerial(Trim(arrDate(0)), Trim(arrDate(1)), Trim(arrDate(2)))
        
    ' Tipagem: Texto
    ElseIf Left(valorBruto, 1) = """" And Right(valorBruto, 1) = """" Then
        ParseValorM = Mid(valorBruto, 2, Len(valorBruto) - 2)
        
    ' Tipagem: Booleano
    ElseIf LCase(valorBruto) = "true" Then ParseValorM = True
    ElseIf LCase(valorBruto) = "false" Then ParseValorM = False
        
    ' Tipagem: Número
    ElseIf IsNumeric(valorBruto) Then
        ParseValorM = CDbl(Replace(valorBruto, ".", ",")) ' Adapta ponto para vírgula do VBA
        
    ' Fallback genérico
    Else
        ParseValorM = valorBruto
    End If
End Function

' ---------------------------------------------------------
' FUNÇÃO AUXILIAR PRIVADA (Com suporte a Arrays e Ranges)
' ---------------------------------------------------------
Private Function FormatarValorM(ByVal Valor As Variant) As String
    ' 1. Trata Objetos do Excel (Ex: Um Range selecionado na planilha)
    If IsObject(Valor) Then
        If TypeName(Valor) = "Range" Then
            ' Pega os valores reais da célula e reinicia o processo
            FormatarValorM = FormatarValorM(Valor.Value2)
            Exit Function
        Else
            FormatarValorM = "null" ' Outros objetos não suportados
            Exit Function
        End If
    End If

    ' 2. Trata Arrays (Vetores do VBA ou Matrizes vindas do Range)
    If IsArray(Valor) Then
        Dim item As Variant
        Dim listaM As String
        
        listaM = "{"
        ' O "For Each" varre o array achatando automaticamente linhas e colunas!
        For Each item In Valor
            ' Ignora células ou itens completamente vazios para não gerar "null" na lista M
            If Not IsEmpty(item) Then
                ' RECURSIVIDADE: Chama a si mesma para descobrir a tipagem de cada item
                listaM = listaM & FormatarValorM(item) & ", "
            End If
        Next item
        
        ' Remove a última vírgula
        If Len(listaM) > 2 Then
            listaM = Left(listaM, Len(listaM) - 2)
        End If
        
        FormatarValorM = listaM & "}"
        Exit Function
    End If

    ' 3. Trata Variáveis Simples (Como já fazíamos)
    Select Case VarType(Valor)
        Case vbString: FormatarValorM = """" & Valor & """"
        Case vbBoolean: FormatarValorM = IIf(Valor, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency: FormatarValorM = Replace(CStr(Valor), ",", ".")
        Case vbDate: FormatarValorM = "#datetime(" & Year(Valor) & ", " & Month(Valor) & ", " & Day(Valor) & ", " & Hour(Valor) & ", " & Minute(Valor) & ", " & Second(Valor) & ")"
        Case vbNull, vbEmpty: FormatarValorM = "null"
        Case Else: FormatarValorM = """" & CStr(Valor) & """"
    End Select
End Function



' ---------------------------------------------------------
' 7. CONTROLES DE ESTADO: Verificar Existência
' ---------------------------------------------------------
Public Function ExisteVariavel(ByVal chave As String) As Boolean
    Dim strM As String, regex As Object
    
    If Not Me.Existe Then
        ExisteVariavel = False
        Exit Function
    End If
    
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        ' Padrão focado apenas em saber se a chave existe no início de alguma linha
        .Pattern = "^[ \t]*\b" & chave & "\b\s*="
    End With
    
    ExisteVariavel = regex.Test(strM)
    Set regex = Nothing
End Function


' ---------------------------------------------------------
' 8. CONTROLES DE ESTADO: Resetar / Limpar Tudo
' ---------------------------------------------------------
Public Sub LimparTudo()
    Dim strM As String
    
    If Not Me.Existe Then Exit Sub
    
    ' Cria o escopo base mínimo
    strM = "let" & vbCrLf & _
           "    _Iniciado = true" & vbCrLf & _
           "in" & vbCrLf & _
           "    _Iniciado"
           
    ' Usa nosso próprio motor para montar o bloco IN perfeitamente
    AtualizarBlocoIn strM
    
    ' Salva no Power Query
    p_wb.Queries(p_QueryName).Formula = strM
    
    ' Feedback opcional via Debug
    Debug.Print "Consulta [" & p_QueryName & "] resetada com sucesso."
End Sub

' ---------------------------------------------------------
' 9. UTILITÁRIOS: Clonar Consulta (Snapshots e Ambientes)
' ---------------------------------------------------------
Public Function Clonar(ByVal NomeOrigem As String, ByVal NomeNovo As String, Optional ByVal Sobrescrever As Boolean = False) As Boolean
    Dim qryOrigem As WorkbookQuery
    Dim qryDestino As WorkbookQuery
    Dim codigoM As String
    
    ' 1. Tenta capturar a consulta de origem
    On Error Resume Next
    Set qryOrigem = p_wb.Queries(NomeOrigem)
    On Error GoTo 0
    
    ' Se a origem não existir, aborta e retorna Falso
    If qryOrigem Is Nothing Then
        Clonar = False
        Exit Function
    End If
    
    ' Extrai a matriz inteira de código M da origem
    codigoM = qryOrigem.Formula
    
    ' 2. Verifica se o destino (o clone) já existe
    On Error Resume Next
    Set qryDestino = p_wb.Queries(NomeNovo)
    On Error GoTo 0
    
    If Not qryDestino Is Nothing Then
        ' O destino já existe. Verifica se o usuário permitiu sobrescrever
        If Sobrescrever Then
            ' Em vez de deletar e recriar (o que quebraria conexões), nós injetamos o código M por cima
            qryDestino.Formula = codigoM
            Clonar = True
        Else
            ' Trava de segurança: Aborta para não esmagar uma consulta existente
            Err.Raise vbObjectError + 2, "clsPQConfig", "A consulta '" & NomeNovo & "' já existe. Use Sobrescrever:=True para substituí-la."
            Clonar = False
        End If
    Else
        ' O destino não existe, então criamos uma consulta nova do zero
        p_wb.Queries.Add Name:=NomeNovo, Formula:=codigoM
        Clonar = True
    End If
End Function


