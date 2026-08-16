Option Explicit

' =========================================================
' CLASSE: clsPQConfig
' OBJETIVO: Gerenciar variáveis em uma consulta do Power Query via código M
' =========================================================

' Variáveis privadas que guardam o estado interno do objeto
Private p_QueryName As String
Private p_wb As Workbook

' ---------------------------------------------------------
' 1. CONSTRUTOR (Executado ao instanciar a classe)
' ---------------------------------------------------------
Private Sub Class_Initialize()
    Set p_wb = ActiveWorkbook
    p_QueryName = "Variaveis" ' Define um nome padrão de fábrica
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