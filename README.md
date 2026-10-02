# Macros de Ponto (AuroraCoop)

Quatro macros que transformam as extrações do sistema de ponto em abas organizadas e,
no fim, no CSV que alimenta o painel dos gestores.

| Pasta | Macro | Entrada → Saída |
|---|---|---|
| `CartaoPonto/` | `ConsolidarCartaoPonto` | Relatório HRCP102 (aba ativa) → aba `Cartao_Consolidado`, uma linha por colaborador/dia |
| `PausasTermicas/` | `FormatarPausasTermicas` | Relatório HRES114 (aba ativa) → a mesma aba, só com as duas tabelas (colaboradores e total) |
| `PuroParaEditado/` | `PURO_Para_EDITADO` | Aba `PURO` → nova aba `EDITADO_MESAA` |
| `GerarCSVPonto/` | `GerarAbaCSV` / `ExportarCSVPorGestor` | Abas de Tratamento, RE, Cartão e Pausas → aba `CSV` → arquivos `.csv` |
| `Dashboard/` | `index.html` | CSV → painel no navegador (gestores e supervisor) |

Contexto completo para o Claude Code em [`CLAUDE.md`](CLAUDE.md). As macros em `.txt`, prontas para colar no VBA, estão em `txt/`; as ferramentas de conferência, em `ferramentas/`.

## Como instalar

Abra o editor do VBA (`Alt + F11`), crie um módulo (**Inserir > Módulo**) e **cole** o conteúdo
do `.bas`. Colar é mais seguro que **Arquivo > Importar**: os arquivos estão em UTF-8 e o
importador do VBA lê em ANSI, o que estraga os acentos das mensagens e dos cabeçalhos
(e, com isso, a busca de colunas como "Matrícula" e "Situação").

## O que mudou nesta versão

### Cartão Ponto: horário de cada dia
- Nova coluna **Cód. Horário**: o código do horário que a pessoa fez naquele dia (coluna "Hor"
  do cartão original, ex.: `0985`, `3303`, `9999`), gravado como texto para manter o zero à esquerda.
- A coluna **Horário** agora mostra a descrição **do horário daquele dia**, tirada da lista
  "Horários:" do próprio colaborador. Antes ela repetia, em todos os dias, só o primeiro horário
  da lista: no HRCP102 de 28/08 a 29/09, 76 dos 202 colaboradores têm mais de um horário no
  período e apareciam com o horário errado em parte dos dias.
- Códigos que não estão na lista do colaborador (9999 DSR, 9998 Compensado, 9997 Feriado,
  9996) ficam com "Horário" em branco. O que aconteceu no dia já aparece em "Descrição Marcação".
- Ordem das colunas, igual à do cartão original:
  `Matrícula · Nome · Cargo · Setor · DT · Sem · Cód. Horário · Horário · Ponto 1..6 ·
  Descrição Marcação · Trabalho · BH · Just. · Injus. · 50% · 60% · 100% · 120% · Ad. Not. ·
  (5 colunas de auditoria)`.
- O setor passou a ser lido direto da linha "Localização:".

### Pausas Térmicas: só as tabelas
A aba formatada fica apenas com:
1. a **tabela de colaboradores**, com o cabeçalho na linha 1;
2. uma linha em branco;
3. a **tabela de total** ("Total Geral").

Saíram o título que ficava em cima e a linha "Soma dos itens..." que ficava embaixo.

### PURO → EDITADO: sem espaços sobrando
Todo texto (nome, marcações, flags e os próprios cabeçalhos) sai sem os espaços que a exportação
coloca no fim (`"FULANO DE TAL      "` → `"FULANO DE TAL"`). Célula que só tinha espaços sai vazia de
verdade. Datas, horas e números não são alterados.

### GerarCSVPonto: Tratamento Internos e Externos
- Lê **todas** as abas cujo nome começa com "Tratamento" (`Tratamento - Internos`,
  `Tratamento - Externos`) com as mesmas regras de antes. Continua funcionando com a planilha
  antiga, que tem uma aba `Tratamento` só.
- Nova coluna **`grupo`** no fim do CSV: `Interno` / `Externo`, conforme a aba de origem. Fica em
  branco na planilha antiga. As linhas de resumo (auditoria, ocorrências curtas, pausas) herdam o
  grupo da matrícula.
- A aba de cadastro é achada pelo prefixo **"RE "** (`RE 08.09`, `RE 30.09`...). Não precisa mais
  editar o código quando o nome muda.
- As abas de Cartão Ponto são achadas pelo prefixo **"Cartão"** e pelo cabeçalho (`Matrícula` e
  `DT` na linha 2). A versão anterior procurava "Cartão ponto" e não achava abas chamadas
  `Cartão Julho` / `Cartão Agosto` / `Cartão atual`. Com isso a auditoria de 3 meses e as
  ocorrências curtas saíam vazias.
- A tabela de Pausas Térmicas é achada pelo **cabeçalho** (`Matrícula` + `Pausa corretas`) em
  qualquer aba, e as colunas são lidas pelo nome. A tabela de total não entra no CSV.
- A mensagem final mostra quantas linhas cada aba de Tratamento gerou, ou por que alguma foi
  pulada (coluna obrigatória faltando), e qual aba RE foi usada.

## Revisão do painel e do CSV (01/10)

Conferido com a planilha de 28/08 a 27/09: cada número do painel foi recalculado direto das abas
de origem (Tratamento, Cartão, Pontonet, Pausas Térmicas) e comparado com o que o painel mostra
no navegador. Com o GerarCSV corrigido, os 15 itens batem (KPIs, os 5 gráficos de ocorrências,
top colaboradores, banco de horas, extra e falta no mesmo dia, ocorrências curtas, pausas
térmicas, demora no PontoNet e histórico).

### GerarCSVPonto
- **Filtros**: antes de ler, tira os filtros de todas as abas e procura a última linha sem usar
  `End(xlUp)`. Com o filtro que estava na aba "Cartão atual", 36 linhas sumiam do CSV
  (28 faltas < 15 min, 6 extras < 15 min, 1 linha de auditoria e 1 hora extra 100%). A mensagem
  final lista as abas que estavam filtradas.
- **Check = "S"**: a linha é ignorada por inteiro, inclusive a hora extra 100%.
- **"Problema horário"** (com ou sem acento) não entra no CSV (`SITUACOES_IGNORAR`).
- **Ocorrência que só tem hora extra**: agora gera também a linha da ocorrência. Antes ela virava
  só "Hora Extra", que o painel não conta, e sumia dos gráficos ("Falta ponto intervalo" e
  parte de "Problema feriado"). Agora cada linha válida da Tratamento é exatamente 1 ocorrência.
- **`#N/A` do PROCX** (ex.: gestor não achado) sai vazio, e não mais como "Erro 2042".
  Quebras de linha dentro das justificativas viram espaço.

### Painel (index.html)
- Saíram os cards "Total de horas extra" e "Total de horas falta/atraso".
- Novo filtro **Grupo** (Interno/Externo), que aparece quando o CSV traz a coluna `grupo`.
- Ignora "Problema horário" e valores de erro do Excel mesmo vindos de um CSV antigo.
- Ocorrências curtas: separadas em Faltas / Extras / Total, contando as linhas como o macro
  gerou (sem descartar por arredondamento).
- Pausas Térmicas: nova coluna "Pausa longa (>0:20)" (o dado já vinha no CSV e não aparecia).
- O topo mostra "N linhas" em vez de "N ocorrências" (o arquivo tem linhas de resumo).

## Ordem de uso no mês

1. Rodar `ConsolidarCartaoPonto` no HRCP102 e colar o resultado numa aba `Cartão <Mês>`.
2. Rodar `FormatarPausasTermicas` no HRES114 (numa cópia) e trazer a aba para a pasta de Tratamento.
3. Rodar `PURO_Para_EDITADO` e tratar as ocorrências em `Tratamento - Internos` / `Tratamento - Externos`.
4. Rodar `GerarAbaCSV` e depois `ExportarCSVPorGestor`.

> Estes macros foram escritos e revisados fora do Excel: a lógica foi conferida simulando-a em
> Python sobre os arquivos reais. Rode sempre numa **cópia** da planilha na primeira vez.
