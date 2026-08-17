Option Explicit

' ==========================================================================
' Rotina:      TestarClassePQ
' Descrição:   Verifica a inicialização da classe clsPQConfig, cria a consulta
'              base se necessário e testa a leitura de uma variável específica.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarClassePQ()
    Dim db As New clsPQConfig
    Dim Valor As Variant
    
    ' Verifica se a consulta base existe no Power Query; se não, cria a estrutura inicial e aborta para permitir inserção manual do usuário
    If Not db.Existe Then
        db.Inicializar
        MsgBox "Consulta criada. Adicione suas variáveis no Power Query e rode novamente!"
        Exit Sub
    End If
    
    ' Extrai o valor da variável "_iniciado" diretamente do código M para o VBA
    Valor = db.LerValor("_iniciado")
    
    ' Valida se a variável foi encontrada e retorna o valor e o tipo de dado interpretado (String, Boolean, etc.)
    If IsEmpty(Valor) Then
        MsgBox "Variável não encontrada!"
    Else
        ' Note que a tipagem do VBA já reconhece como String ou Boolean automaticamente
        MsgBox "Valor encontrado: " & Valor & vbCrLf & "Tipo do Dado: " & TypeName(Valor)
    End If
End Sub




' ==========================================================================
' Rotina:      TestarEscritaPQ
' Descrição:   Demonstra a capacidade da classe de injetar diferentes tipos
'              de dados nativos do VBA diretamente no código M.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarEscritaPQ()
    Dim db As New clsPQConfig
    
    ' Aponta para a consulta padrão "Variaveis" para realizar as injeções de código
    db.NomeConsulta = "Variaveis"
    
    ' 1. Criar/Atualizar uma String (Caminho de diretório)
    db.DefinirValor "DiretorioPDF", "C:\Arquivos\Relatorios\"
    
    ' 2. Criar/Atualizar um Booleano (Controle de fluxo/Ribbon)
    db.DefinirValor "ExportAoATT", False
    
    ' 3. Criar/Atualizar um Número inteiro (Parâmetro de impressão)
    db.DefinirValor "CopiasTerreo", 3
    
    ' 4. Criar/Atualizar uma Data (Garante timestamp da última rodada da macro)
    db.DefinirValor "UltimaAtualizacao", Now
    
    MsgBox "Variáveis salvas no Power Query com sucesso!" & vbCrLf & vbCrLf & _
           "Vá no Power Query e olhe o editor avançado da consulta 'Variaveis' para ver a mágica.", vbInformation
End Sub




' ==========================================================================
' Rotina:      TestarGeracaoDeRegistro
' Descrição:   Testa a inclusão e exclusão dinâmica de variáveis, assegurando
'              que o bloco 'in' do Power Query se ajuste automaticamente.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarGeracaoDeRegistro()
    Dim db As New clsPQConfig
    
    ' 1. Cria novas variáveis ambientais de forma sequencial (o "in" será montado lindamente)
    db.DefinirValor "DiretorioRelatorios", "Z:\PCP\Relatorios\"
    db.DefinirValor "NotificarEmail", True
    
    ' 2. Cria uma variável temporária que servirá de cobaia para exclusão
    db.DefinirValor "VariavelTesteDelete", 999
    
    ' (Pause aqui se quiser ver a variável criada no painel)
    
    ' 3. Deleta a variável temporária, forçando o motor interno a reestruturar a sintaxe M para não deixar vírgulas sobrando
    db.RemoverVariavel "VariavelTesteDelete"
    
    MsgBox "Tudo concluído! Vá olhar o card da sua consulta no Excel."
End Sub




' ==========================================================================
' Rotina:      TestarLote
' Descrição:   Processa a inserção de múltiplas variáveis de uma só vez
'              utilizando um Dicionário, reduzindo o I/O no Power Query.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarLote()
    Dim db As New clsPQConfig
    Dim lote As Object
    
    ' Usa Late Binding para instanciar o Scripting.Dictionary, dispensando a necessidade de habilitar a referência na máquina do usuário final
    Set lote = CreateObject("Scripting.Dictionary")
    
    ' Preenche o "pacote" com um mix de tipos de dados (Numérico, String, Booleano, Data)
    lote.Add "FiltroAno", 2026
    lote.Add "FiltroMes", "Agosto"
    lote.Add "CaminhoRede", "\\Servidor\PCP\"
    lote.Add "PermiteAtualizacao", True
    lote.Add "RodadoEm", Now
    
    ' Descarrega o dicionário inteiro em uma única chamada, otimizando o tempo de processamento
    db.DefinirVarios lote
    
    MsgBox "Lote de " & lote.Count & " variáveis salvo em uma única operação!"
End Sub




' ==========================================================================
' Rotina:      TestarGeracaoDeListas
' Descrição:   Testa a conversão avançada de Arrays VBA e objetos Range do
'              Excel em Listas nativas do Power Query (formato M).
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarGeracaoDeListas()
    Dim db As New clsPQConfig
    Dim arrayNativo As Variant
    
    ' TESTE 1: Injeção a partir de um Array direto do VBA
    ' A matriz heterogênea é convertida internamente para a sintaxe de lista M: {"Aprovado", "Pendente", 2026}
    arrayNativo = Array("Aprovado", "Pendente", 2026)
    db.DefinirValor "StatusPermitidos", arrayNativo
    
    ' TESTE 2: Injeção a partir de uma seleção ativa no Excel
    ' Avalia se o usuário selecionou células. Em caso positivo, extrai a propriedade Value2 da matriz e injeta no M
    If TypeName(Selection) = "Range" Then
        db.DefinirValor "EstadosSelecionados", Selection
    End If
    
    MsgBox "Listas M criadas com sucesso! Vá conferir no Power Query."
End Sub




' ==========================================================================
' Rotina:      TestarCorrecaoEDebug
' Descrição:   Avalia a extração e a fidelidade da tipagem de dados complexos
'              ao serem lidos de volta para o VBA, além de testar o Debug.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarCorrecaoEDebug()
    Dim db As New clsPQConfig
    Dim dict As Object
    
    ' 1. Cria variáveis que historicamente geram bugs de formatação (Data e Arrays)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "DataExata", Now
    dict.Add "MinhaLista", Array(10, 20, 30)
    db.DefinirVarios dict
    
    ' 2. Solicita um dump completo do PQ habilitando o log no Immediate Window (Verificação Imediata) para facilitar auditorias
    ' Abra a janela de verificação imediata (CTRL+G) para ver o relatório!
    Set dict = db.ListarTodas(DebugMode:=True)
    
    ' 3. Valida se o parser reverteu a data em M (#datetime) adequadamente para um objeto Date nativo do VBA
    MsgBox "A Variável DataExata voltou como tipo: " & TypeName(dict("DataExata"))
End Sub



' ==========================================================================
' Rotina:      TestarControlesDeEstado
' Descrição:   Demonstra verificações lógicas de ambiente baseadas na
'              existência de variáveis, bem como o reset geral da consulta.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarControlesDeEstado()
    Dim db As New clsPQConfig
    
    ' 1. Valida se a chave primária de execução (CaminhoRede) está presente no ambiente M antes de permitir o avanço do sistema
    If Not db.ExisteVariavel("CaminhoRede") Then
        MsgBox "Atenção: O Caminho da Rede ainda não foi configurado!", vbExclamation
    Else
        MsgBox "O sistema está pronto para uso."
    End If
    
    ' 2. Método destrutivo para limpar configurações residuais e retornar ao estado de fábrica (Comentado por segurança)
    ' db.LimparTudo
End Sub




' ==========================================================================
' Rotina:      TestarPerfisDeAmbiente
' Descrição:   Testa os recursos de clonagem e sobrescrita de consultas,
'              úteis para gerenciar diferentes perfis ou criar backups.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarPerfisDeAmbiente()
    Dim db As New clsPQConfig
    Dim sucesso As Boolean
    
    ' EXEMPLO 1: Cria um Snapshot exato em tempo real do código M na consulta de backup
    'sucesso = db.Clonar("Variaveis", "Variaveis_BKP", Sobrescrever:=True)
    
    If sucesso Then
        MsgBox "Backup criado com sucesso! Olhe no painel do Power Query.", vbInformation
    End If
    
    ' EXEMPLO 2: Restaura as configurações do ambiente copiando o código M do backup de volta para a produção
    ' Digamos que você bagunçou a consulta "Variaveis". Basta jogar o BKP por cima dela!
    db.Clonar "Variaveis_BKP", "Variaveis", Sobrescrever:=True
End Sub




' ==========================================================================
' Rotina:      TestarMotorRAD
' Descrição:   Verifica a funcionalidade de exportar a consulta Power Query
'              como um arquivo de texto criptografado em Tags, e sua reimportação.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarMotorRAD()
    Dim db As New clsPQConfig
    Dim caminhoBase As String
    
    ' Descobre dinamicamente o diretório da área de trabalho do usuário logado via Windows Scripting Host
    caminhoBase = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    ' 1. EXPORTAÇÃO
    ' Consolida o código M e os metadados descritivos em um único arquivo de texto padronizado
    ' (Não esqueça de ir no PQ e colocar uma descrição na consulta "Variaveis" para testar)
    db.ExportarConsultaParaTXT "Variaveis", caminhoBase & "Componente_Variaveis.txt"
    
    MsgBox "O arquivo foi gerado no seu Desktop. Abra o TXT e veja o formato das Tags!", vbInformation
    
    ' 2. IMPORTAÇÃO
    ' Faz o parse das tags do arquivo exportado e gera/atualiza a estrutura de forma espelhada no Excel
    db.ImportarConsultaDeTXT caminhoBase & "Componente_Variaveis.txt", Sobrescrever:=True
End Sub


' ==========================================================================
' Rotina:      TestarInjecaoTexto
' Descrição:   Verifica a gravação de um texto bruto (HTML) passado via string
'              diretamente para o Power Query, validando sua re-leitura.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarInjecaoTexto()
    Dim db As New clsPQConfig
    Dim htmlOriginal As String
    Dim htmlLido As String
    
    ' Define o nome da consulta alvo
    db.NomeConsulta = "App_HTML_Teste"
    
    ' 1. Cria uma string simulando um bloco HTML digitado na IDE
    htmlOriginal = "<div class='container'>" & vbCrLf & _
                   "    <h1>Bem-vindo ao OpenVBA-Web</h1>" & vbCrLf & _
                   "    <button onclick='alert(""Sucesso!"")'>Clique Aqui</button>" & vbCrLf & _
                   "</div>"
                   
    ' 2. GRAVAÇÃO: Salva no Power Query enviando pelo parâmetro TextoConteudo
    db.SalvarTextoBruto TextoConteudo:=htmlOriginal
    
    ' 3. LEITURA: Puxa o código de volta já limpo (sem as aspas de escape do M)
    htmlLido = db.LerTextoBruto()
    
    ' 4. Validação
    If htmlOriginal = htmlLido Then
        MsgBox "Teste de Injeção via String: SUCESSO!" & vbCrLf & vbCrLf & _
               "O código foi salvo e recuperado do Power Query com integridade total.", vbInformation, "OpenVBA-Web"
    Else
        MsgBox "Teste falhou. O texto recuperado não é idêntico ao original.", vbCritical, "Erro de Integridade"
    End If
End Sub


' ==========================================================================
' Rotina:      TestarInjecaoArquivo
' Descrição:   Abre um seletor de arquivos nativo (FileDialog) para o usuário
'              escolher um arquivo físico, consome o conteúdo mapeando o fluxo
'              através da classe configurada e valida a persistência na query.
' Parâmetros:  Nenhum
' ==========================================================================
Sub TestarInjecaoArquivo()
    Dim db As New clsPQConfig
    Dim fd As FileDialog
    Dim CaminhoArquivo As String
    Dim htmlLido As String
    
    ' 1. Configura o estado da classe definindo a consulta alvo
    db.NomeConsulta = "App_HTML_Arquivo"
    
    ' 2. Instancia e configura o Seletor de Arquivos do Office
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "OpenVBA-Web: Selecione o arquivo para injetar no Power Query"
        .AllowMultiSelect = False
        .Filters.Clear
        .Filters.Add "Arquivos OpenVBA-Web", "*.html; *.css; *.js", 1
        .Filters.Add "Todos os Arquivos", "*.*", 2
        
        ' Verifica se o usuário selecionou algo ou cancelou a janela
        If .Show = -1 Then
            CaminhoArquivo = .SelectedItems(1)
        Else
            MsgBox "Operação cancelada. Nenhum arquivo foi selecionado.", vbExclamation, "OpenVBA-Web RAD"
            Exit Sub
        End If
    End With
    
    ' 3. GRAVAÇÃO: A classe gerencia o arquivo apontado e injeta na query configurada
    db.SalvarTextoBruto CaminhoArquivo:=CaminhoArquivo
    
    ' 4. LEITURA: Recupera o código armazenado para checar o parser de escape
    htmlLido = db.LerTextoBruto()
    
    ' 5. Validação do sucesso do motor de arquivos
    If htmlLido <> "" Then
        MsgBox "Teste de Arquivo [" & db.NomeConsulta & "]: SUCESSO!" & vbCrLf & vbCrLf & _
               "O arquivo selecionado (" & Dir(CaminhoArquivo) & ") foi processado e salvo com sucesso no Power Query.", vbInformation, "OpenVBA-Web"
    Else
        MsgBox "Ocorreu uma falha na extração ou gravação do arquivo.", vbCritical, "Erro"
    End If
    
    Set fd = Nothing
End Sub


