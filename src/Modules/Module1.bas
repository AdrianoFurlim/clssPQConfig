Option Explicit

Sub TestarClassePQ()
    Dim db As New clsPQConfig
    Dim Valor As Variant
    
    ' Verifica se a consulta base existe, se não, cria.
    If Not db.Existe Then
        db.Inicializar
        MsgBox "Consulta criada. Adicione suas variáveis no Power Query e rode novamente!"
        Exit Sub
    End If
    
    ' Lê uma variável da sua consulta (substitua pelo nome real de uma variável sua)
    Valor = db.LerValor("_iniciado")
    
    If IsEmpty(Valor) Then
        MsgBox "Variável não encontrada!"
    Else
        ' Note que a tipagem do VBA já reconhece como String ou Boolean automaticamente
        MsgBox "Valor encontrado: " & Valor & vbCrLf & "Tipo do Dado: " & TypeName(Valor)
    End If
End Sub

Sub TestarEscritaPQ()
    Dim db As New clsPQConfig
    
    ' Aponta para a consulta que você usa (o padrão da classe já é "Variaveis")
    db.NomeConsulta = "Variaveis"
    
    ' 1. Criar/Atualizar uma String
    db.DefinirValor "DiretorioPDF", "C:\Arquivos\Relatorios\"
    
    ' 2. Criar/Atualizar um Booleano (como os seus botões do Ribbon)
    db.DefinirValor "ExportAoATT", False
    
    ' 3. Criar/Atualizar um Número
    db.DefinirValor "CopiasTerreo", 3
    
    ' 4. Criar/Atualizar uma Data (Bônus corporativo!)
    db.DefinirValor "UltimaAtualizacao", Now
    
    MsgBox "Variáveis salvas no Power Query com sucesso!" & vbCrLf & vbCrLf & _
           "Vá no Power Query e olhe o editor avançado da consulta 'Variaveis' para ver a mágica.", vbInformation
End Sub

Sub TestarGeracaoDeRegistro()
    Dim db As New clsPQConfig
    
    ' 1. Cria novas variáveis (o "in" será montado lindamente)
    db.DefinirValor "DiretorioRelatorios", "Z:\PCP\Relatorios\"
    db.DefinirValor "NotificarEmail", True
    
    ' 2. Cria uma variável temporária
    db.DefinirValor "VariavelTesteDelete", 999
    
    ' (Pause aqui se quiser ver a variável criada no painel)
    
    ' 3. Deleta a variável temporária (o "in" vai se ajustar sozinho)
    db.RemoverVariavel "VariavelTesteDelete"
    
    MsgBox "Tudo concluído! Vá olhar o card da sua consulta no Excel."
End Sub

Sub TestarLote()
    Dim db As New clsPQConfig
    Dim lote As Object
    
    ' Usa Late Binding para criar o Dicionário sem precisar marcar Referências
    Set lote = CreateObject("Scripting.Dictionary")
    
    ' Preenche o "pacote" de variáveis
    lote.Add "FiltroAno", 2026
    lote.Add "FiltroMes", "Agosto"
    lote.Add "CaminhoRede", "\\Servidor\PCP\"
    lote.Add "PermiteAtualizacao", True
    lote.Add "RodadoEm", Now
    
    ' Envia o pacote inteiro para a classe processar em milissegundos
    db.DefinirVarios lote
    
    MsgBox "Lote de " & lote.Count & " variáveis salvo em uma única operação!"
End Sub

Sub TestarGeracaoDeListas()
    Dim db As New clsPQConfig
    Dim arrayNativo As Variant
    
    ' TESTE 1: Passando um Array direto do VBA
    ' Pode ter números e textos misturados, a classe vai se virar perfeitamente
    arrayNativo = Array("Aprovado", "Pendente", 2026)
    db.DefinirValor "StatusPermitidos", arrayNativo
    
    ' TESTE 2: Passando um Range do Excel (Seleção atual)
    ' A classe descobre que é Range, extrai os valores e converte em lista M
    If TypeName(Selection) = "Range" Then
        db.DefinirValor "EstadosSelecionados", Selection
    End If
    
    MsgBox "Listas M criadas com sucesso! Vá conferir no Power Query."
End Sub


Sub TestarCorrecaoEDebug()
    Dim db As New clsPQConfig
    Dim dict As Object
    
    ' 1. Cria variáveis problemáticas (O Bug da vírgula foi extinto)
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "DataExata", Now
    dict.Add "MinhaLista", Array(10, 20, 30)
    db.DefinirVarios dict
    
    ' 2. Testa a extração com o modo Debug Ativado
    ' Abra a janela de verificação imediata (CTRL+G) para ver o relatório!
    Set dict = db.ListarTodas(DebugMode:=True)
    
    ' 3. Valida se a Data voltou pro VBA como tipo Date e não String
    MsgBox "A Variável DataExata voltou como tipo: " & TypeName(dict("DataExata"))
End Sub

Sub TestarControlesDeEstado()
    Dim db As New clsPQConfig
    
    ' 1. Verifica se uma configuração crucial já existe antes de tentar exportar algo
    If Not db.ExisteVariavel("CaminhoRede") Then
        MsgBox "Atenção: O Caminho da Rede ainda não foi configurado!", vbExclamation
    Else
        MsgBox "O sistema está pronto para uso."
    End If
    
    ' 2. (Opcional) Descomente a linha abaixo para testar o Reset
    db.LimparTudo
End Sub


Sub TestarPerfisDeAmbiente()
    Dim db As New clsPQConfig
    Dim sucesso As Boolean
    
    ' EXEMPLO 1: Criando um Snapshot/Backup das variáveis atuais
    'sucesso = db.Clonar("Variaveis", "Variaveis_BKP", Sobrescrever:=True)
    
    If sucesso Then
        MsgBox "Backup criado com sucesso! Olhe no painel do Power Query.", vbInformation
    End If
    
    ' EXEMPLO 2: O "Rollback" (Restaurando o ambiente)
    ' Digamos que você bagunçou a consulta "Variaveis". Basta jogar o BKP por cima dela!
    db.Clonar "Variaveis_BKP", "Variaveis", Sobrescrever:=True
End Sub


Sub TestarMotorRAD()
    Dim db As New clsPQConfig
    Dim caminhoBase As String
    
    ' Pega o caminho do desktop automaticamente
    caminhoBase = CreateObject("WScript.Shell").SpecialFolders("Desktop") & "\"
    
    ' 1. EXPORTAÇÃO
    ' (Não esqueça de ir no PQ e colocar uma descrição na consulta "Variaveis" para testar)
    'db.ExportarConsultaParaTXT "Variaveis", caminhoBase & "Componente_Variaveis.txt"
    
    'MsgBox "O arquivo foi gerado no seu Desktop. Abra o TXT e veja o formato das Tags!", vbInformation
    
    ' 2. IMPORTAÇÃO
    ' Ele lê o arquivo e cria um pacote novo
    db.ImportarConsultaDeTXT caminhoBase & "Componente_Variaveis.txt", Sobrescrever:=True
End Sub

