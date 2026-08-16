# Power Query Config Framework (VBA)

![VBA](https://img.shields.io/badge/Language-VBA-blue) ![Excel](https://img.shields.io/badge/Platform-Excel-green) ![Power Query](https://img.shields.io/badge/Integration-Power_Query-yellow) ![License](https://img.shields.io/badge/License-MIT-lightgrey)

## Sobre o Projeto
O objetivo desta classe (`clsPQConfig`) é fornecer um framework de gerenciamento de estado e variáveis no Power Query via código M[cite: 1]. Através do VBA, é possível manipular os dados de forma programática, permitindo a comunicação dinâmica entre macros do Excel e as transformações de consultas[cite: 1].

## Funcionalidades
* **CRUD de Variáveis:** Verificação, leitura, atualização, inserção e exclusão de variáveis injetadas diretamente na linguagem M[cite: 1].
* **Visualização como Registro:** O sistema pode reestruturar o bloco `in` da consulta para apresentar as variáveis declaradas em um formato de registro (`[chave=chave]`), facilitando o debug visual[cite: 1].
* **Conversor de Tipos Nativo:** Converte tipos de dados VBA (incluindo literais, booleanos, números e datas `#datetime`) para a sintaxe rigorosa exigida pelo Power Query[cite: 1]. Transforma Matrizes (Arrays) e Ranges em Listas M nativas (Ex: `{1, 2, "a"}`)[cite: 1].
* **Utilitários de Snapshot:** Possui recursos embutidos para clonar consultas existentes (criando backups) ou limpar todo o código residual retornando ao estado de fábrica[cite: 1].
* **Ferramentas RAD:** Permite a exportação e importação do código M e metadados (como Descrição) usando arquivos de texto com proteção por marcação de Tags estilo XML[cite: 1].

## Pré-requisitos
* **Microsoft Excel (Macro-Enabled):** Interage nativamente com as consultas através do objeto `ActiveWorkbook`[cite: 1].
* **Componentes COM (Late Binding):** A classe não exige ativação manual de referências na máquina do usuário[cite: 2]. Internamente, ela instancia os seguintes objetos em tempo de execução:
  * `VBScript.RegExp`: Para leitura e manipulação de strings via Expressão Regular[cite: 1].
  * `Scripting.Dictionary`: Para passagem de parâmetros em lotes[cite: 1, 2].
  * `ADODB.Stream`: Para extração e gravação de arquivos de texto preservando a acentuação (UTF-8)[cite: 1].

## Como Usar (Exemplos de Código)

### 1. Inicializando e Lendo uma Variável
~~~vba
Sub ExemploLeitura()
    Dim db As New clsPQConfig
    Dim Valor As Variant
    
    ' Verifica se a consulta alvo existe; caso contrário, cria o esqueleto M[cite: 1, 2]
    If Not db.Existe Then[cite: 2]
        db.Inicializar[cite: 2]
        MsgBox "Consulta base criada. Adicione suas variáveis!"[cite: 2]
        Exit Sub[cite: 2]
    End If[cite: 2]
    
    ' Extrai a variável e interpreta seu tipo de dado (Ex: String, Boolean)[cite: 1, 2]
    Valor = db.LerValor("_Iniciado")[cite: 2]
    
    If Not IsEmpty(Valor) Then[cite: 2]
        MsgBox "Valor: " & Valor & vbCrLf & "Tipo: " & TypeName(Valor)[cite: 2]
    End If[cite: 2]
End Sub
~~~

### 2. Injetando Variáveis e Arrays (Escrita)
~~~vba
Sub ExemploEscrita()
    Dim db As New clsPQConfig
    ' Define o nome da consulta onde a injeção ocorrerá[cite: 2]
    db.NomeConsulta = "Variaveis"[cite: 2]
    
    ' Insere/Atualiza variáveis escalares convertendo-as para M[cite: 1, 2]
    db.DefinirValor "DiretorioPDF", "C:\Arquivos\Relatorios\"[cite: 2]
    db.DefinirValor "ExportAoATT", False[cite: 2]
    db.DefinirValor "CopiasTerreo", 3[cite: 2]
    db.DefinirValor "UltimaAtualizacao", Now[cite: 2]
    
    ' Injeta uma matriz (Array) que virará uma lista no Power Query[cite: 1, 2]
    Dim arrayNativo As Variant[cite: 2]
    arrayNativo = Array("Aprovado", "Pendente", 2026)[cite: 2]
    db.DefinirValor "StatusPermitidos", arrayNativo[cite: 2]
End Sub
~~~

### 3. Escrita em Lote para Alta Performance
~~~vba
Sub ExemploEmLote()
    Dim db As New clsPQConfig
    Dim lote As Object
    
    ' Cria o dicionário usando Late Binding[cite: 2]
    Set lote = CreateObject("Scripting.Dictionary")[cite: 2]
    
    ' Agrupa todas as propriedades[cite: 2]
    lote.Add "FiltroAno", 2026[cite: 2]
    lote.Add "FiltroMes", "Agosto"[cite: 2]
    lote.Add "CaminhoRede", "\\Servidor\PCP\"[cite: 2]
    lote.Add "PermiteAtualizacao", True[cite: 2]
    
    ' Salva o lote inteiro no código M do Power Query reduzindo o I/O da macro[cite: 1, 2]
    db.DefinirVarios lote[cite: 2]
End Sub
~~~

## Estrutura / Referência da API

| Recurso | Tipo | Descrição do Comportamento |
| :--- | :--- | :--- |
| `NomeConsulta` | Propriedade | Aponta ou retorna o nome da consulta alvo (Padrão de fábrica: "Variaveis")[cite: 1]. |
| `Existe` | Propriedade | Valida se a consulta referenciada foi criada e está disponível no Excel[cite: 1]. |
| `Inicializar()` | Rotina | Injeta a sintaxe base obrigatória (`let...in`) na consulta, caso seja vazia[cite: 1]. |
| `LimparTudo()` | Rotina | Sobrescreve toda a instrução da consulta existente, retornando-a ao estado virgem[cite: 1]. |
| `LerValor(chave)` | Função | Captura o conteúdo de uma variável após o `=` via Regex e aplica tipagem para VBA[cite: 1]. |
| `DefinirValor(c, v)`| Rotina | Encaminha um único par de chave/valor para injeção programática via Regex[cite: 1]. |
| `DefinirVarios(Dict)`| Rotina | Analisa um objeto dicionário e injeta todas as propriedades em uma única gravação estruturada[cite: 1]. |
| `RemoverVariavel(c)`| Rotina | Deleta fisicamente a declaração de uma variável, ajustando vírgulas sintáticas do M automaticamente[cite: 1]. |
| `ListarTodas(Debug)`| Função | Retorna todas as variáveis presentes do bloco let como `Dictionary`, suportando log imediato[cite: 1]. |
| `Clonar(...)` | Função | Cria a duplicata de uma consulta e de seus códigos M. Útil para snapshots e ambientes[cite: 1]. |
| `ExportarConsulta...`| Rotina | Extrai o `M` e a Descrição do Power Query e cria um arquivo `.txt` formatado com Tags RAD (UTF-8)[cite: 1]. |
| `ImportarConsulta...`| Rotina | Realiza o processo de leitura do `.txt`, interpretando as tags para reconstruir a consulta automaticamente no Excel[cite: 1]. |

## Licença

Este projeto está licenciado sob a licença MIT.

---

**MIT License**

Copyright (c) 2026 [Adriano Furtado Lima]

A permissão é por meio deste concedida, gratuitamente, a qualquer pessoa que obtenha uma cópia 
deste software e dos arquivos de documentação associados (o "Software"), para negociar 
no Software sem restrições, incluindo, sem limitação, os direitos de usar, copiar, modificar, 
mesclar, publicar, distribuir, sublicenciar e/ou vender cópias do Software, e permitir 
às pessoas a quem o Software é fornecido que o façam, sujeito às seguintes condições:

O aviso de direitos autorais acima e este aviso de permissão devem ser incluídos em todas as 
cópias ou partes substanciais do Software.

O SOFTWARE É FORNECIDO "COMO ESTÁ", SEM GARANTIA DE QUALQUER TIPO, EXPRESSA OU IMPLÍCITA, 
INCLUINDO, MAS NÃO SE LIMITANDO ÀS GARANTIAS DE COMERCIALIZAÇÃO, ADEQUAÇÃO A UM DETERMINADO 
FIM E NÃO VIOLAÇÃO. EM NENHUM CASO OS AUTORES OU DETENTORES DE DIREITOS AUTORAIS SERÃO 
RESPONSÁVEIS POR QUALQUER RECLAMAÇÃO, DANOS OU OUTRA RESPONSABILIDADE, SEJA EM UMA AÇÃO 
DE CONTRATO, ATO ILÍCITO OU DE OUTRA FORMA, DECORRENTE DE, FORA DE OU EM CONEXÃO COM O 
SOFTWARE OU O USO OU OUTRAS NEGOCIAÇÕES NO SOFTWARE.
