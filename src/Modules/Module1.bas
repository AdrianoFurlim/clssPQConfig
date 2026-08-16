Option Explicit

Sub TestarClassePQ()
    Dim db As New clsPQConfig
    Dim Valor As Variant
    
    ' Verifica se a consulta base existe, se não, cria.
    If Not db.Existe Then
        db.InicializarConsulta
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
