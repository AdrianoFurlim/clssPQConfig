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
Public Sub InicializarConsulta()
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
' 4. MÉTODOS DE DADOS (CRUD - Leitura)
' ---------------------------------------------------------
Public Function LerValor(ByVal Chave As String) As Variant
    Dim strM As String
    Dim regex As Object
    Dim matches As Object
    Dim valorBruto As String
    
    strM = Me.CodigoMAtual
    
    ' Se a consulta não existir ou estiver vazia, retorna nulo
    If strM = "" Then
        LerValor = Empty
        Exit Function
    End If
    
    ' Inicializa o motor RegEx via Late Binding (não requer ativar referências)
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        
        ' PADRÃO REGEX: Busca o nome da chave, o sinal de igual e captura tudo até achar uma vírgula ou quebra de linha
        .Pattern = Chave & "\s*=\s*(.*?)(?:,|\r|\n|$)"
    End With
    
    If regex.Test(strM) Then
        Set matches = regex.Execute(strM)
        
        ' Extrai apenas o grupo capturado (o valor em si, ignorando o nome da chave e o "=")
        valorBruto = Trim(matches(0).SubMatches(0))
        
        ' --- TIPAGEM DINÂMICA ---
        ' Analisa o formato da string extraída para converter para o tipo correto no VBA
        If Left(valorBruto, 1) = """" And Right(valorBruto, 1) = """" Then
            ' É texto: Remove as aspas do início e do fim
            LerValor = Mid(valorBruto, 2, Len(valorBruto) - 2)
            
        ElseIf LCase(valorBruto) = "true" Then
            ' É booleano verdadeiro
            LerValor = True
            
        ElseIf LCase(valorBruto) = "false" Then
            ' É booleano falso
            LerValor = False
            
        ElseIf IsNumeric(valorBruto) Then
            ' É número: Converte para Double para evitar erros com decimais
            LerValor = CDbl(valorBruto)
            
        Else
            ' Fallback para qualquer outra coisa (ex: listas no M, ou nulos)
            LerValor = valorBruto
        End If
    Else
        LerValor = Empty ' Chave não encontrada no código M
    End If
    
    Set regex = Nothing
End Function



' ---------------------------------------------------------
' 5. MÉTODOS DE DADOS (CRUD - Update, Create & Delete)
' ---------------------------------------------------------
Public Sub DefinirValor(ByVal Chave As String, ByVal Valor As Variant)
    Dim strM As String, valorFormatado As String, novaLinha As String
    Dim regex As Object
    
    If Not Me.Existe Then Me.InicializarConsulta
    strM = Me.CodigoMAtual
    
    ' Tipagem
    Select Case VarType(Valor)
        Case vbString: valorFormatado = """" & Valor & """"
        Case vbBoolean: valorFormatado = IIf(Valor, "true", "false")
        Case vbInteger, vbLong, vbDouble, vbSingle, vbCurrency: valorFormatado = Replace(CStr(Valor), ",", ".")
        Case vbDate: valorFormatado = "#datetime(" & Year(Valor) & ", " & Month(Valor) & ", " & Day(Valor) & ", " & Hour(Valor) & ", " & Minute(Valor) & ", " & Second(Valor) & ")"
        Case Else: valorFormatado = """" & CStr(Valor) & """"
    End Select
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False: .IgnoreCase = True: .Multiline = True
        .Pattern = "(\b" & Chave & "\s*=\s*)(.*?)(?=,|\r|\n|$)"
    End With
    
    If regex.Test(strM) Then
        ' UPDATE
        strM = regex.Replace(strM, "$1" & valorFormatado)
    Else
        ' CREATE
        novaLinha = "let" & vbCrLf & "    " & Chave & " = " & valorFormatado & ","
        regex.Pattern = "^let\b"
        strM = regex.Replace(strM, novaLinha)
    End If
    
    ' >>> CHAMA O MOTOR DE RECONSTRUÇÃO DO "IN" AQUI <<<
    AtualizarBlocoIn strM
    
    ' Salva
    p_wb.Queries(p_QueryName).Formula = strM
    Set regex = Nothing
End Sub

Public Sub RemoverVariavel(ByVal Chave As String)
    Dim strM As String
    Dim regex As Object
    
    If Not Me.Existe Then Exit Sub
    strM = Me.CodigoMAtual
    
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .Multiline = True
        ' Padrão para remover a linha inteira da chave
        .Pattern = "^[ \t]*\b" & Chave & "\b\s*=.*?(?:,|\r?\n|$)"
    End With
    
    If regex.Test(strM) Then
        ' Apaga a variável do bloco let
        strM = regex.Replace(strM, "")
        
        ' Correção de sintaxe crucial: Se deletamos a última variável,
        ' a variável de cima vai ficar com uma vírgula sobrando antes do "in". Isso quebra o M.
        regex.Pattern = ",(\s*\bin\b)"
        strM = regex.Replace(strM, "$1") ' Remove a vírgula orfã
        
        ' >>> CHAMA O MOTOR DE RECONSTRUÇÃO DO "IN" AQUI <<<
        AtualizarBlocoIn strM
        
        ' Salva no Power Query
        p_wb.Queries(p_QueryName).Formula = strM
    End If
    
    Set regex = Nothing
End Sub



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
        .MultiLine = True
        .Pattern = "\bin\b[\s\S]*"
    End With
    
    ' A função RTrim garante que não deixaremos quebras de linha acumuladas antes do novo "in"
    apenasLet = RTrim(regex.Replace(strM, ""))
    
    ' PASSO 2: Buscar as chaves APENAS dentro da parte isolada (apenasLet)
    With regex
        .Global = True
        ' Lê palavras no início da linha, ignorando tabulações ou espaços
        .Pattern = "^[ \t]*([A-Za-z_][A-Za-z0-9_]*)[ \t]*="
    End With
    
    If regex.test(apenasLet) Then
        Set matches = regex.Execute(apenasLet)
        chaves = "    [" & vbCrLf
        
        ' Monta o bloco [Chave = Chave] com quebra de linha para ficar elegante
        For Each match In matches
            chave = match.SubMatches(0)
            chaves = chaves & "        " & chave & " = " & chave & "," & vbCrLf
        Next match
        
        ' Remove a última vírgula e a quebra de linha (", " + CR + LF = 3 caracteres)
        If Len(chaves) > 3 Then
            chaves = Left(chaves, Len(chaves) - 3) & vbCrLf
        End If
        chaves = chaves & "    ]"
        
        ' PASSO 3: Junta o "let" original, o "in", e o novo bloco limpo
        strM = apenasLet & vbCrLf & "in" & vbCrLf & chaves
    End If
    
    Set regex = Nothing
End Sub