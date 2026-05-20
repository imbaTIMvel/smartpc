# SmartPC

Programa de criação de planilhas de Prestação de Contas p/ o time financeiro do HBR.

![Logo do SmartPC](assets/icons/smartpc.ico)

## 1. Requisitos

Para uso adequado do programa, o usuário deve possuir:

- Sistema Operacional: Windows 10 ou 11

Para uso da funcionalidade `Abrir planilha quando estiver pronta`, é recomendável que o usuário possua o *Microsoft Excel* instalado na versão mais recente.

## 2. Guia de Uso

### 2.1 Baixando e Instalando o Programa

Para usar o `SmartPC`, primeiro, você deve baixar o arquivo .exe disponível [aqui](link). Procure pela versão mais recente (*Latest*) e clique no arquivo .exe para fazer o download.

> [!Warning]
> Caso você ainda tenha o executável de uma versão antiga do programa, recomenda-se excluí-lo.

Baixado o programa, você pode colocar o arquivo .exe onde achar melhor.

### 2.2 Abrindo o Programa

Feito isso, clique no arquivo .exe para abrir o programa.

![Tela inicial](assets/images/exe_in_downloads.png)

> [!Warning]
> É possível que o *Windows Defender* acuse o programa como "software perigoso". Neste caso, para executá-lo, você deve clicar em `Mais Informações` e, depois, no botão `Executar assim mesmo`.

![Windows Defender acusando o programa](assets/images/windows_defender_01.png)

![Executar assim mesmo](assets/images/windows_defender_02.png)

### 2.3 Interface do Programa

![Interface do programa - superior](assets/images/ui_01.png)

O programa possui três campos para inserção de arquivos de entrada. São eles:

| Campo         | Extensões de arquivo aceitas | Aceita mais de um arquivo? |
| ------------- | ---------------------------- | -------------------------- |
| Planilha Base | .xlsx                        | Não                        |
| Extrato CC    | .xlsx e .xls                 | Não                        |
| Relatório(s)  | .xlsx                        | Sim                        |

Para cada um dos campos, 

![Interface do programa - superior](assets/images/ui_02.png)

Além disso, o programa possui:
- `Título PC`: Uma **caixa de texto** para inserir o título da Prestação de Contas (PC) que será usado para preencher a coluna "PC" na planilha de saída e a célula associada à "PC:" na folha de rosto;
- `Abrir planilha quando estiver pronta`: Um **toggle switch** que permite que o usuário defina se a planilha de saída será aberta ou não após a execução do programa;
- `Gerar Planilha`: O botão que inicia a execução do programa.

### 2.4 Modos de Execução

Por convenção, o programa pode ser executado em três "modos" diferentes, a depender dos arquivos de entrada inseridos em cada um dos campos. Trataremos abaixo a execução do programa em cada um desses modos:

#### 2.4.1 Criação de Planilha de Prestação de Contas

| Arquivo de Entrada | Deve ser adicionado? |
| ------------------ | -------------------- |
| Planilha Base      | Não                  |
| Extrato CC         | Sim                  |
| Relatório(s)       | Sim (pelo menos um)  |



[...]

## 3. Releases

[...]

## 4. Desenvolvimento

**Autor:**
Timóteo Altoé (*handle*: [imbaTIMvel](github.com/imbaTIMvel))

**Datas:**
`29/04/2026` Início do projeto

`05/05/2026` Lançamento da versão *alfa* - para testes internos

`20/05/2026` Publicação da primeira versão oficial no GitHub

`21/05/2026` Lançamento da versão *beta* - para testes
