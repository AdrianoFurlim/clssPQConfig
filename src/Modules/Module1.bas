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

Sub TestarEscritaPQ1()
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
