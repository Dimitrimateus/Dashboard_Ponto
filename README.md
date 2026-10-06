# Macros de Ponto (AuroraCoop)

Quatro macros que transformam as extrações do sistema de ponto em abas organizadas e,
no fim, no CSV que alimenta o painel dos gestores.

| Pasta | Macro | Entrada → Saída |
|---|---|---|
| `CartaoPonto/` | `ConsolidarCartaoPonto` | Relatório HRCP102 (aba ativa) → aba `Cartao_Consolidado`, uma linha por colaborador/dia |
| `PausasTermicas/` | `FormatarPausasTermicas` | Relatório HRES114 (aba ativa) → a mesma aba, só com as duas tabelas (colaboradores e total) |
| `PuroParaEditado/` | `PURO_Para_EDITADO` | Aba `PURO` → nova aba `EDITADO_MESAA` |
| `MarcacoesDigitadas/` | `FormatarMarcacoesDigitadas` | Relatório "Marcações digitadas" (aba ativa) → a mesma aba, como tabela (`Marcações Digitadas`) |
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

## Painel: horas do Cartão, período, card do colaborador e modo explicação (06/10)

### GerarCSVPonto
- Nova linha **"Resumo Horas Cartão"**, uma por colaborador: soma o **Cartão atual inteiro**
  (hora extra = colunas 50% + 100%; banco de horas = coluna BH). Não depende da Tratamento,
  do Check nem do período.
- O CSV ganhou 2 colunas no fim (agora são 25): `minutos_hora_extra_cartao` e
  `minutos_banco_horas_cartao`. Só essa linha nova as preenche.
- A mensagem final diz quantos colaboradores entraram nesse resumo.

### Painel (index.html)
- **Banco de horas e hora extra por colaborador**: a barra tem a parte **verde** (hora extra)
  e a **vermelha** (banco de horas), e à direita vem o **saldo** (extra − banco), em verde com
  "+" ou em vermelho com "−". Os números vêm do resumo novo do Cartão. O botão "Mostrar todos"
  abre a lista inteira. Com um CSV antigo (sem o resumo), o gráfico usa as linhas da
  Tratamento e avisa que o CSV precisa ser gerado de novo.
- **Período em destaque** no centro da faixa preta do topo. Ele acompanha o filtro de período
  (aparece "Período filtrado" quando o filtro muda) e agora vai até o último dia do Cartão
  (ex.: 28/08 a 27/09).
- **Filtros limpos automaticamente** quando o arquivo sai ("Carregar outro arquivo") e quando
  um arquivo novo é aberto. Antes, o colaborador digitado e a busca do histórico continuavam
  valendo.
- **Card do colaborador**: clicar num nome (gráficos de top colaboradores e de banco de horas,
  e em qualquer tabela: histórico, extra e falta, curtas, pausas, demora) abre um card com
  matrícula, gestor, cargo, setor, ocorrências no período (por tipo e linha a linha), horas do
  Cartão e saldo, extra e falta no mesmo dia, curtas, demora no PontoNet e pausas.
- **Entenda o painel** (botão no topo): liga o modo explicação. É o mesmo painel, com uma
  caixa laranja em cada número, gráfico e tabela dizendo de onde vem a informação. Funciona
  também sem arquivo carregado.

A conferência das ferramentas (`ferramentas/`) foi atualizada: 15/15 itens batem com a
planilha de 28/08 a 27/09 (versão _2), incluindo extra, banco e saldo de cada colaborador.

## Lista de colaboradores do ORG (06/10)

### GerarCSVPonto
- Lê a aba **ORG** (relação colaborador × gestor; achada pelo nome "ORG..." e pelas colunas
  Matricula / COLABORADOR / GESTOR) e grava uma linha **"Cadastro ORG"** por pessoa: nome,
  matrícula, gestor, cargo e a área do ORG (no campo `situacao`). Não muda as colunas do CSV.
- Matrícula repetida no ORG conta uma vez. Linha com a matrícula em `#N/A` entra sem matrícula,
  a não ser que o mesmo nome já esteja no ORG com matrícula válida.
- O gestor das linhas-resumo (horas do Cartão, curtas, auditoria, pausas) de quem não aparece
  na Tratamento passa a vir do ORG. Antes essas pessoas saíam "Sem gestor".

### Painel
- Nova seção **Lista de colaboradores (ORG)**: todo mundo do ORG, com gestor, cargo, área,
  ocorrências no filtro e saldo de horas do Cartão. Tem busca e um seletor (com/sem ocorrência,
  sem gestor, sem matrícula). O nome abre o card do colaborador.
- Aviso em vermelho com **quem tem dados no arquivo mas não está no ORG**, onde a pessoa aparece
  e, quando há, a linha parecida do ORG (mesmo nome com matrícula em `#N/A`, ou nome escrito
  com uma letra diferente).
- O card do colaborador mostra a área do ORG e avisa "Não está no ORG".

## Marcações digitadas (06/10)

Marcação digitada é a batida que não veio do relógio: o colaborador incluiu no PontoNet com
justificativa, ou o RH lançou no sistema.

### FormatarMarcacoesDigitadas (nova macro)
- Rode com o relatório "Marcações digitadas" aberto (aba ativa, numa cópia). A aba vira uma
  tabela com cabeçalho na linha 1 e é renomeada para `Marcações Digitadas`:
  `Matrícula · Colaborador · Cód. Cargo · Cargo · Cód. Local · Local · Origem · Data · Hora ·
  Dia da semana · Coletor · Função · Motivo · Justificativa`.
- Descarta os cabeçalhos de página repetidos, separa o código e a descrição do local e junta o
  motivo que o relatório quebra em duas linhas ("Problemas no dispositivo de registro de" +
  "ponto"). As colunas são achadas pelo cabeçalho do relatório.
- Depois, traga a aba para a planilha de Tratamento (Mover ou copiar).

### GerarCSVPonto
- Acha a aba pelo cabeçalho (Matrícula, Data, Hora, Motivo, Justificativa) e grava uma linha
  **"Marcação Digitada"** por batida: data com a hora (arredondada ao minuto), motivo em
  `situacao`, gestor (Tratamento ou ORG), setor e cargo.
- O CSV ganhou 2 colunas no fim (agora são 27): `justificativa_marcacao` e `origem_marcacao`.

### Painel
- Nova seção **Marcações digitadas**: indicadores (total, pessoas, % por esquecimento, quantas
  caem num dia que também tem ocorrência na Tratamento, % em horário redondo), motivo, dia da
  semana, data, top 10 colaboradores, por gestor, horário da batida, esquecimento repetido,
  justificativas mais usadas e a lista completa com busca. Respeita os filtros de gestor, grupo,
  cargo, colaborador e período.
- A lista do ORG ganhou a coluna "Digitadas", o card do colaborador mostra as batidas digitadas
  dele e o aviso "fora do ORG" passa a considerar as digitadas.

### Ferramentas
- `formatar_marcacoes.py`: a mesma formatação da macro, em Python (para conferir sem Excel).
- `adicionar_aba_xlsx.py`: acrescenta a aba formatada à planilha de Tratamento sem regravar o
  resto do arquivo (só mexe nos índices do pacote e acrescenta estilos no fim).
- Conferência: 17/17.

## Ordem de uso no mês

1. Rodar `ConsolidarCartaoPonto` no HRCP102 e colar o resultado numa aba `Cartão <Mês>`.
2. Rodar `FormatarPausasTermicas` no HRES114 (numa cópia) e trazer a aba para a pasta de Tratamento.
3. Rodar `PURO_Para_EDITADO` e tratar as ocorrências em `Tratamento - Internos` / `Tratamento - Externos`.
4. Rodar `FormatarMarcacoesDigitadas` no relatório de Marcações digitadas e trazer a aba.
5. Rodar `GerarAbaCSV` e depois `ExportarCSVPorGestor`.

> Estes macros foram escritos e revisados fora do Excel: a lógica foi conferida simulando-a em
> Python sobre os arquivos reais. Rode sempre numa **cópia** da planilha na primeira vez.
