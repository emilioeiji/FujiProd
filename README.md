# ExcelMacroRunner

Aplicativo WinForms em C# para executar um `.bat`, abrir planilhas Excel `.xlsm`, rodar macros VBA em sequência e monitorar a atualização de um arquivo HTML gerado no processo.

O foco do projeto é substituir scripts manuais e automações `.vbs` por um executável único, com interface simples, log em arquivo e controle mais robusto do Excel via `Microsoft.Office.Interop.Excel`.

## O que o app faz

- Lê as configurações de `appsettings.json`
- Executa um arquivo `.bat`
- Abre até 3 workbooks `.xlsm`
- Executa 3 macros em sequência
- Monitora um arquivo HTML gerado pelo processo
- Mata `EXCEL.EXE` e reinicia o ciclo quando o HTML para de atualizar
- Roda em loop contínuo sem travar a interface
- Impede múltiplas instâncias usando `Mutex`
- Salva logs em arquivo `.log`

## Estrutura do projeto

- [ExcelMacroRunner.sln](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner.sln)
- [ExcelMacroRunner/Program.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/Program.cs:1)
- [ExcelMacroRunner/MainForm.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/MainForm.cs:1)
- [ExcelMacroRunner/MainForm.Designer.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/MainForm.Designer.cs:1)
- [ExcelMacroRunner/MacroRunnerService.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/MacroRunnerService.cs:1)
- [ExcelMacroRunner/ServiceConfig.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/ServiceConfig.cs:1)
- [ExcelMacroRunner/FileLogger.cs](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/FileLogger.cs:1)
- [ExcelMacroRunner/appsettings.json](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/appsettings.json:1)
- [ExcelMacroRunner/Assets/AppIcon.ico](/C:/Users/emili/source/repos/FujiProd/ExcelMacroRunner/Assets/AppIcon.ico)

## Interface

O painel principal mostra:

- Status atual
- Última execução
- Última atualização do HTML
- Próxima ação
- Mensagem atual
- Caminho do `appsettings.json`
- Pasta de log
- Log recente na tela

Botões disponíveis:

- `Iniciar`
- `Parar`
- `Reiniciar agora`
- `Matar Excel`
- `Abrir pasta do log`

## appsettings.json

Exemplo:

```json
{
  "ServiceConfig": {
    "BatPath": "C:\\Scripts\\atualizar.bat",
    "ExcelJobs": [
      {
        "Name": "Atualiza Base 1",
        "ExcelFilePath": "C:\\Relatorios\\AtualizadorBase1.xlsm",
        "MacroName": "Modulo1.AtualizarBase1"
      },
      {
        "Name": "Atualiza Base 2",
        "ExcelFilePath": "C:\\Relatorios\\AtualizadorBase2.xlsm",
        "MacroName": "Modulo1.AtualizarBase2"
      },
      {
        "Name": "Consolida e Gera HTML",
        "ExcelFilePath": "C:\\Relatorios\\ConsolidadorDashboard.xlsm",
        "MacroName": "Modulo1.ConsolidarEGerarHtml"
      }
    ],
    "HtmlPath": "C:\\Relatorios\\dashboard.html",
    "LoopDelaySeconds": 5,
    "HtmlMaxAgeMinutes": 10,
    "KillExcelOnTimeout": true,
    "ShowExcel": false,
    "LogFolder": "C:\\Logs\\ExcelMacroRunner"
  }
}
```

## Como preencher o MacroName

Nao, a macro nao precisa obrigatoriamente estar nesse formato.

O app chama:

```csharp
excelApplication.Run(job.MacroName);
```

Ou seja: o valor de `MacroName` precisa estar no formato que o Excel aceita no `Application.Run`.

Os formatos mais comuns sao:

1. Macro no workbook aberto e publica:

```json
"MacroName": "ConsolidarEGerarHtml"
```

2. Macro publica dentro de modulo padrao:

```json
"MacroName": "Modulo1.ConsolidarEGerarHtml"
```

3. Macro em outro workbook, ou quando voce quer ser mais explicito:

```json
"MacroName": "'ConsolidadorDashboard.xlsm'!Modulo1.ConsolidarEGerarHtml"
```

### Qual eu recomendo

Se a macro esta no mesmo `.xlsm` que o app abriu naquele job, normalmente estes dois funcionam:

- `ConsolidarEGerarHtml`
- `Modulo1.ConsolidarEGerarHtml`

Eu prefiro `Modulo1.ConsolidarEGerarHtml` porque fica mais claro e costuma ajudar quando ha varias macros com nomes parecidos.

Se a sua macro chama outra macro em outro arquivo, e esse outro arquivo precisa estar aberto, o mais seguro costuma ser usar o nome completo:

```json
"MacroName": "'OutroArquivo.xlsm'!ModuloInterno.MinhaMacro"
```

## Dica importante sobre erro em macro em cascata

Se o app encontra o `.xlsm`, abre o Excel e falha no `Run`, os motivos mais comuns sao:

- a macro esta com nome diferente do configurado
- a macro nao e `Public`
- a macro esta em `ThisWorkbook` ou em modulo de planilha, nao em modulo padrao
- a macro chamada abre outro workbook e esse segundo arquivo nao existe
- o segundo workbook abre com caixa de dialogo, prompt de seguranca ou erro interno
- a macro depende de caminho local diferente no computador de destino
- a macro espera que o Excel esteja visivel ou com alguma janela ativa

Para automacao, o ideal e:

- colocar a macro principal em um modulo padrao
- deixar a macro como `Public Sub NomeDaMacro()`
- tratar erros dentro do VBA
- registrar log dentro do proprio VBA, se possivel
- evitar `MsgBox`, `InputBox` e prompts manuais

Exemplo recomendado em VBA:

```vb
Public Sub ConsolidarEGerarHtml()
    On Error GoTo TratarErro

    Call AtualizarBase1
    Call AtualizarBase2
    Call GerarHtml
    Exit Sub

TratarErro:
    Err.Raise Err.Number, "ConsolidarEGerarHtml", Err.Description
End Sub
```

## Compatibilidade x86 e x64

Para `Interop` com Excel, a arquitetura do app deve acompanhar a arquitetura do Office instalado.

- `Office 32 bits` -> publicar `win-x86`
- `Office 64 bits` -> publicar `win-x64`

Para conferir no Excel:

- `Arquivo`
- `Conta`
- `Sobre o Excel`

## Build

Comando:

```powershell
dotnet build "C:\Users\emili\source\repos\FujiProd\ExcelMacroRunner\ExcelMacroRunner.csproj"
```

## Publicacao single-file via command line

### Office 32 bits

```powershell
dotnet publish "C:\Users\emili\source\repos\FujiProd\ExcelMacroRunner\ExcelMacroRunner.csproj" -c Release -r win-x86 --self-contained true /p:PublishSingleFile=true /p:EnableCompressionInSingleFile=true
```

### Office 64 bits

```powershell
dotnet publish "C:\Users\emili\source\repos\FujiProd\ExcelMacroRunner\ExcelMacroRunner.csproj" -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true /p:EnableCompressionInSingleFile=true
```

### Publicacao com pasta de saida definida

```powershell
dotnet publish "C:\Users\emili\source\repos\FujiProd\ExcelMacroRunner\ExcelMacroRunner.csproj" -c Release -r win-x86 --self-contained true /p:PublishSingleFile=true /p:EnableCompressionInSingleFile=true -o "C:\Users\emili\source\repos\FujiProd\publish\ExcelMacroRunner-x86"
```

## Pasta final publicada

O executavel publicado costuma sair em:

- `bin\Release\net9.0-windows\win-x86\publish`
- `bin\Release\net9.0-windows\win-x64\publish`

Importante:

- mantenha o `appsettings.json` na mesma pasta do `.exe`
- o `single-file` vale para o executavel, mas o `appsettings.json` continua externo

## Logs

Os logs sao gravados na pasta configurada em `LogFolder`.

Exemplo:

```text
C:\Logs\ExcelMacroRunner
```

O painel tambem mostra as mensagens recentes em tempo real.

## Fluxo de execucao

Ao clicar em `Iniciar`, o loop faz:

1. Validar os caminhos
2. Executar o `.bat`
3. Abrir o Excel
4. Abrir os workbooks configurados
5. Rodar as macros em sequencia
6. Fechar workbook e Excel
7. Liberar objetos COM
8. Verificar o `LastWriteTime` do HTML
9. Reiniciar o ciclo se o HTML estiver velho
10. Aguardar o delay configurado

## Requisitos

- Windows
- .NET 9 SDK para build local
- Microsoft Excel instalado
- Arquivos `.xlsm` acessiveis no caminho configurado
- Permissao de leitura e escrita nas pastas usadas

## Pacote NuGet usado

Pacote principal:

- `Microsoft.Office.Interop.Excel`

Referencia oficial:

- [NuGet Microsoft.Office.Interop.Excel](https://www.nuget.org/packages/microsoft.office.interop.excel/)

## Observacoes praticas

- O programa nao e um Windows Service real
- O recomendado e iniciar pelo Agendador de Tarefas no login do Windows
- O Agendador deve apenas iniciar o app
- O loop interno do app cuida das repeticoes
- O app foi pensado para ser simples de manter e facil de diagnosticar por log

## Proximos passos sugeridos

- ajustar os caminhos reais no `appsettings.json`
- testar primeiro com `ShowExcel: true`
- confirmar o nome exato da macro manualmente no VBA
- testar x86 se o Office for 32 bits
- se a macro em cascata continuar falhando, logar erros dentro do proprio VBA
