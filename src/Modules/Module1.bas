Option Explicit

Sub TestarClassePQ()
    Dim db As New clsPQConfig
    Dim valor As Variant
    
    ' Verifica se a consulta base existe, se não, cria.
    If Not db.Existe Then
        db.InicializarConsulta
        MsgBox "Consulta criada. Adicione suas variáveis no Power Query e rode novamente!"
        Exit Sub
    End If
    
    ' Lê uma variável da sua consulta (substitua pelo nome real de uma variável sua)
    valor = db.LerValor("_iniciado")
    
    If IsEmpty(valor) Then
        MsgBox "Variável não encontrada!"
    Else
        ' Note que a tipagem do VBA já reconhece como String ou Boolean automaticamente
        MsgBox "Valor encontrado: " & valor & vbCrLf & "Tipo do Dado: " & TypeName(valor)
    End If
End Sub