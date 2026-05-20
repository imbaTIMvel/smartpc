# SmartPC

Programa de criação de planilhas de Prestação de Contas p/ o time financeiro do Instituto Hardware BR - HBR.

![Logo do SmartPC](assets/icons/smartpc.ico)

## 1. Requisitos

Para uso adequado do programa, o usuário deve possuir:

- Sistema Operacional: Windows 10 ou 11

Para uso da funcionalidade `Abrir planilha quando estiver pronta`, é recomendado que o usuário possua o *Microsoft Excel* instalado na versão mais recente.

## 2. Guia de Uso

### 2.1 Baixando e Instalando o Programa

Para usar o `SmartPC`, primeiro, você deve baixar o arquivo `.exe` disponível [aqui](link). Procure pela versão mais recente (*Latest*) e clique no arquivo .exe para fazer o download.

> [!Warning]
> Caso você ainda tenha o executável de uma versão antiga do programa, recomenda-se excluí-lo.

Baixado o programa, você pode colocar o arquivo `.exe` onde achar melhor.

### 2.2 Abrindo o Programa

Feito isso, clique no arquivo `.exe` para abrir o programa.

![Tela inicial](assets/images/exe_in_downloads.png)

> [!Warning]
> É possível que o *Windows Defender* acuse o programa como "software perigoso". Neste caso, para executá-lo, você deve clicar em `Mais Informações` e, depois, no botão `Executar assim mesmo`.

![Windows Defender acusando o programa](assets/images/windows_defender_01.png)

![Executar assim mesmo](assets/images/windows_defender_02.png)

### 2.3 Interface do Programa

![Interface do programa - superior](assets/images/ui_01.png)

#### 2.3.1 Campos de Arquivos

O programa possui três campos para inserção de arquivos (planilhas Excel) de entrada. São eles:

| Campo         | Extensões de arquivo aceitas | Padronização do arquivo                                           | Aceita mais de um arquivo? |
| ------------- | ---------------------------- | ----------------------------------------------------------------- | -------------------------- |
| Planilha Base | .xlsx                        | [standard_sheet](assets/standard_sheet/standard_sheet.xlsx)       | Não                        |
| Extrato CC    | .xlsx e .xls                 | Extratos CC do Itaú ou do Banco BRB                               | Não                        |
| Relatório(s)  | .xlsx                        | Relatórios de Pagamentos e/ou Recebimentos, retirados do Octalink | Sim                        |

Para cada um dos campos, há dois botões: `Selecionar` e `Remover`. Ao clicar em `Selecionar`, o programa abre um diálogo do *Explorador de Arquivos*, permitindo que o usuário selecione o arquivo Excel correspondente ao campo (mais de um arquivo, se forem Relatórios).

![Diálogo de seleção de arquivo](assets/images/file_select_01.png)

Após selecionar o(s) arquivo(s), o campo de arquivo inserido é atualizado.

![Interface com arquivos selecionados](assets/images/file_select_02.png)

Ao clicar em `Remover`, se houver arquivo(s) selecionado(s) no campo correspondente, o programa remove o arquivo selecionado, deixando o campo vazio.

![Remoção de arquivo](assets/images/file_removal_01.png)

![Interface sem arquivos selecionados](assets/images/file_removal_02.png)

#### 2.3.2 Demais Recursos

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

Inserindo um Extrato CC (do Itaú ou do Banco BRB) e ao menos um Relatório (de Pagamentos ou Recebimentos, do Octalink), ao clicar em `Gerar Planilha`, o programa deve gerar uma planilha no formato [standard_sheet](assets/standard_sheet/standard_sheet.xlsx), com as colunas:
- `Nº`: Contagem do item da planilha gerada, serve como referência para o espelhamento de dados na folha de rosto;
- `RESUMO DO GASTO`: Preenchido com "-", deve ser preenchido posteriormente pelo usuário (operador);
- `RUBRICA`: Retirado da coluna "Rubrica" do(s) Relatório(s) do Octalink;
- `FORNECEDOR`: Retirado da coluna "Razão Social" do(s) Relatórios do Octalink;
- `PC`: Texto inserido na caixa de texto "Título PC" do programa;
- `CNPJ ou CPF`: Retirado da coluna "CPF/CNPJ" do(s) Relatórios do Octalink;
- `DATA DA NF`: Retirado da coluna "Emissão" do(s) Relatórios do Octalink;
- `(número) NF/RECIBO`: Retirado da coluna "Nº Documento" do(s) Relatórios do Octalink;
- `Nº EXTRATO`: Retirado da coluna "Numero Documento" ou "DOC" do Extrato CC;
- `DATA DO PAGAMENTO`: Retirado da coluna "Data" ou "DATA" do Extrato CC. Faz-se a correspondência com a coluna "Vencimento/Mov." do(s) Relatórios do Octalink;
- `VALOR PAGO`: Retirado da coluna "Valor R$" ou "Valor" do Extrato CC. Faz-se a correspondência com a coluna "Valor Total" do(s) Relatórios do Octalink. Imediatamente acima do nome da coluna, há uma célula com a soma dos itens;
- `CÓD. PEDIDO`: Retirado da coluna "Cód. Pedido" do(s) Relatórios do Octalink;
- `CÓD. RMS`: Retirado da coluna "Cód. Controle" do(s) Relatórios do Octalink;
- `Mês_pagto`: Formatado como *mês/ano* a partir do dado inserido na coluna `DATA DO PAGAMENTO`;
- `HISTÓRICO`: Retirado da coluna "Histórico" ou "Descrição" do Extrato CC. Faz-se a correspondência com a coluna "Histórico" do(s) Relatórios do Octalink;
- `EXCEÇÕES DO ALGORITMO`: Erros do algoritmo em caso de não haver correspondências entre o Extrato CC e o(s) Relatório(s), ou em caso de haver múltiplas correspondências sem desempate por **histórico**.

Para associar adequadamente as informações entre o Extrato CC e o(s) Relatório(s), o programa procura por correspondências de **valor** e **data**, recorrendo a desempate por **histórico** (como fallback) em casos de múltiplas ocorrências. Em caso de múltiplas correspondências (sem possibilidade de tratamento por **histórico**) ou nenhuma correspondência, o programa reporta os erros na coluna `EXCEÇÕES DO ALGORITMO`.

> [!Note]
> Com os arquivos de entrada inseridos, ao clicar no botão `Gerar Planilha`, o programa deve fazer as correspondências entre o Extrato CC e o(s) Relatório(s), gerando uma planilha de saída (com folha de rosto) e permitindo que o usuário escolha o local de salvamento do arquivo após o processamento.

![Gerando a planilha](assets/images/create_sheet_01.png)

![Caixa de aviso](assets/images/create_sheet_02.png)

![Salvando o arquivo](assets/images/create_sheet_03.png)

#### 2.4.2 Emenda de Planilhas

| Arquivo de Entrada | Deve ser adicionado? |
| ------------------ | -------------------- |
| Planilha Base      | Sim                  |
| Extrato CC         | Sim                  |
| Relatório(s)       | Sim (pelo menos um)  |

Inserindo uma Planilha Base (no formato [standard_sheet](assets/standard_sheet/standard_sheet.xlsx)), um Extrato CC (do Itaú ou do Banco BRB) e ao menos um Relatório (de Pagamentos ou Recebimentos, do Octalink), ao clicar em `Gerar Planilha`, o programa deve adicionar as informações associadas e processadas do Extrato CC e do(s) Relatório(s) à base da Planilha Base, associando as colunas conforme a ordem da padronização [standard_sheet](assets/standard_sheet/standard_sheet.xlsx), e preservando os dados da Planilha Base (inclusive anotações nas células abaixo da tabela em si). Adicionalmente, caso a Planilha Base não possua folha de rosto, o programa deve adicionar uma à planilha de saída.

> [!Note]
> Com os arquivos de entrada inseridos, ao clicar no botão `Gerar Planilha`, o programa deve fazer as correspondências entre o Extrato CC e o(s) Relatório(s) e emendar os dados novos à Planilha Base, gerando uma planilha de saída (com folha de rosto) e permitindo que o usuário escolha o local de salvamento do arquivo após o processamento.

![Gerando a planilha](assets/images/splice_sheet_01.png)

![Caixa de aviso](assets/images/splice_sheet_02.png)

![Salvando o arquivo](assets/images/splice_sheet_03.png)

#### 2.4.3 Adição de Folha de Rosto

| Arquivo de Entrada | Deve ser adicionado? |
| ------------------ | -------------------- |
| Planilha Base      | Sim                  |
| Extrato CC         | Não                  |
| Relatório(s)       | Não                  |

Inserindo apenas uma Planilha Base (no formato [standard_sheet](assets/standard_sheet/standard_sheet.xlsx)) **SEM FOLHA DE ROSTO**, ao clicar em `Gerar Planilha`, o programa deve criar uma folha de rosto e adicioná-la à planilha de saída, mantendo os dados da Planilha Base inalterados.

> [!Note]
> Com os arquivos de entrada inseridos, ao clicar no botão `Gerar Planilha`, o programa deve adicionar uma folha de rosto à Planilha Base, gerando uma planilha de saída (com folha de rosto e dados preservados) e permitindo que o usuário escolha o local de salvamento do arquivo após o processamento.

![Gerando a planilha](assets/images/add_face_01.png)

![Caixa de aviso](assets/images/add_face_02.png)

![Salvando o arquivo](assets/images/add_face_03.png)

## 3. Releases

`v0.1.0` SmartPC (*beta release*)

> [!Warning]
> O lançamento beta (*beta release*) foi desenvolvido para testes internos, visando identificar e corrigir bugs antes do lançamento de uma versão estável.

Data de lançamento: `20/05/2026`

Para fazer o download desta versão, clique [aqui](link).

*Release* inicial do programa de criação de planilhas de Prestação de Contas p/ o time financeiro do Instituto Hardware BR - HBR.

**Features:**
- Compatível com planilhas Excel, dos tipos:
  - `Planilha Base`: Na padronização [standard_sheet](assets/standard_sheet/standard_sheet.xlsx), no formato .xlsx;
  - `Extrato CC`: Como exportados pelos bancos Itaú e BRB, nos formatos .xlsx ou .xls;
  - `Relatório(s)`: Relatórios de Pagamentos e/ou Relatórios de Recebimentos, conforme exportados pelo Octalink, no formato .xlsx;
- Possui 3 (três) modos de operação, escolhidos de acordo com os arquivos de entrada inseridos:
  - `Criação de Planilha de Prestação de Contas`: Com Extrato CC e, no mínimo, um Relatório;
  - `Emenda de Planilhas`: Com Planilha Base, Extrato CC e, no mínimo, um Relatório;
  - `Adição de Folha de Rosto`: Apenas com Planilha Base;
- Permite que o usuário escolha o diretório de salvamento para a planilha (.xlsx) de saída.

Clique [aqui](link) para acessar o **changelog completo**.

## 4. Desenvolvimento

**Autor:**

Timóteo Altoé (*handle*: [imbaTIMvel](github.com/imbaTIMvel))

**Datas:**

`29/04/2026` Início do projeto

`05/05/2026` Lançamento da versão *alfa* - para testes internos

`20/05/2026` Publicação da primeira versão oficial no GitHub

`21/05/2026` Lançamento da versão *beta* - para testes
