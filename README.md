# Macros de Ponto (AuroraCoop)

Quatro macros que transformam as extrações do sistema de ponto em abas organizadas e,
no fim, no CSV que alimenta o painel dos gestores.

| Pasta | Macro | Entrada → Saída |
|---|---|---|
| `CartaoPonto/` | `ConsolidarCartaoPonto` | Relatório HRCP102 (aba ativa) → aba `Cartao_Consolidado`, uma linha por colaborador/dia |
| `PausasTermicas/` | `FormatarPausasTermicas` | Relatório HRES114 (aba ativa) → a mesma aba, só com as duas tabelas (colaboradores e total) |
| `PuroParaEditado/` | `PURO_Para_EDITADO` | Aba `PURO` → nova aba `EDITADO_MESAA` |
| `MarcacoesDigitadas/` | `FormatarMarcacoesDigitadas` | Relatório "Marcações digitadas" (aba ativa) → a mesma aba, como tabela (`Marcações Digitadas`) |
| `Afastamentos/` | `FormatarAfastamentos` | Relatório "Histórico de Afastamentos" HRCL006 (aba ativa) → a mesma aba, como tabela (`Afastamentos`) |
| `BancoHoras/` | `FormatarBancoHoras` | Relatório de banco de horas HRES105 ("BH", aba ativa) → a mesma aba, como tabela (`Banco de Horas`) |
| `GerarCSVPonto/` | `GerarAbaCSV` / `ExportarCSVPorGestor` | Abas de Tratamento, RE, Cartão e Pausas → aba `CSV` → arquivos `.csv` |
| `Dashboard/` | `index.html` | CSV → painel no navegador (gestores e supervisor) |

Contexto completo para o Claude Code em [`CLAUDE.md`](CLAUDE.md). As macros em `.txt`, prontas para colar no VBA, estão em `txt/`; as ferramentas de conferência, em `ferramentas/`.

## Jeito mais fácil (03/11): um código só

Cole **só** o `txt/PainelPonto_TudoEmUm.txt` num módulo da planilha de Tratamento (apague os módulos
antigos) e rode **GerarCSVsDaPasta** (Alt+F8). Ela pede a pasta dos relatórios, confere se estão lá os
12 arquivos (`Cartão ponto 1`, `Cartão ponto 2`, `Cartão ponto 3`, `Pausas térmicas 3 meses`,
`Digitadas 3 meses`, `Afastamentos 3 meses`, `Pontonet 3 meses`, `Excedentes 3 meses`,
`Intrajornada 3 meses`, `Interjornada 3 meses`, `Interjornada semanal 3 meses` e, desde 04/11, `BH`),
avisa os que faltam, pergunta onde salvar os CSVs e faz o resto. Os CSVs saem com o **nome do gestor
no começo** (`<GESTOR>.csv`) mais o `TODOS OS GESTORES.csv`. O módulo é montado por
`ferramentas/montar_modulo_unico.py` a partir dos `.bas` de cada macro.

## Como instalar

Abra o editor do VBA (`Alt + F11`), crie um módulo (**Inserir > Módulo**) e **cole** o conteúdo
do `.bas`. Colar é mais seguro que **Arquivo > Importar**: os arquivos estão em UTF-8 e o
importador do VBA lê em ANSI, o que estraga os acentos das mensagens e dos cabeçalhos
(e, com isso, a busca de colunas como "Matrícula" e "Situação").

## Ajustes e conferência ampliada (13/11)

- Saiu o filtro de período (cada seção tem o seu seletor de mês).
- No PontoNet, "Esquecimento do crachá" é um motivo separado de "Esquecimento".
- Quem não está na RE (cadastro de ativos) aparece como **desligado**; sem gestor, fica no grupo "Desligados" do
  filtro de gestor. A macro GerarCSV marca essas linhas e acha o gestor pelo nome de quem está no ORG sem matrícula.
- Corrigidos: a "Média de todos" da demora no PontoNet; "<1%" nos pedaços de barra; o mapa de afastamentos com
  "Só saída médica e horas" sem nenhum dia inteiro.
- A conferência agora cobre o painel inteiro, inclusive o resumo de cada pessoa, cada gestor, cada mês dos
  seletores e as janelas "de onde vem" (ferramentas `auditar_mais.py`, `extrair_extra_painel.js`, `auditar_extra.py`).
  Cole de novo o `txt/PainelPonto_TudoEmUm.txt`.

## Painel para todas as unidades (11/11)

- Saíram as seções **Ocorrências**, **Colaboradores em destaque** e **Histórico**: dependiam da planilha de
  tratamento, que só existe na unidade de Recife/Cabo. Saíram também os filtros de **grupo** e de **tipo de
  ocorrência** e, no resumo da pessoa, o que vinha das ocorrências.
- "Extra e falta no mesmo dia" e "Ocorrências curtas" (vêm do cartão ponto) estão agora no fim da seção **Jornada**.
- A **Lista de colaboradores** mostra só o cadastro, como está na lista de equipes: matrícula, colaborador, área,
  cargo e gestor.
- Ao abrir um arquivo, o painel vai para **Arquivos carregados** e mostra um aviso: cada quadro tem um botão **?**
  com a explicação ("Não mostrar de novo" vale para aquele computador).
- Macro **GerarCSVPonto**: não lê mais as abas "Tratamento - Internos/Externos" (constante `LER_ABAS_TRATAMENTO`,
  `True` volta ao jeito antigo). Funciona numa planilha sem essas abas. Cole de novo o `txt/PainelPonto_TudoEmUm.txt`.

## Conferência completa do painel (11/11)

Para provar que cada número do painel está certo, sem depender do código do painel:

1. `ferramentas/extrair_tudo_painel.js` abre o painel no navegador, carrega o CSV e anota tudo o que a tela mostra.
2. `ferramentas/auditar_csv_painel.py` refaz cada conta direto do CSV, em Python, e compara (OK/FALHA, item a item).
3. `ferramentas/relatorio_conferencia.js` monta o relatório em Word, seção por seção (o PDF sai do Word).

Os comandos estão no CLAUDE.md, seção 10. Todo o código (painel e macros) tem comentários em português simples;
no começo do script do painel há um "Guia rápido" com os limites que o RH pode mudar.

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
  caixa laranja em cada número, gráfico e tabela dizendo de onde vem a informação. Só aparece
  depois de carregar o arquivo (sem arquivo, a tela inicial traz o "Como usar o painel").

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
  **"Marcação Digitada"** por batida de Origem "D" (as de Origem "E" são ignoradas): data com a hora (arredondada ao minuto), motivo em
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

## Absenteísmo — atestados (06/10)

### GerarCSVPonto
- Lê a coluna "Descrição Marcação" dos até 3 Cartões. Por enquanto só **atestados** (texto que
  contém "atest": "Atestado ate 15 dias", o noturno e, quando aparecer, o odontológico). Para
  incluir acidente de trabalho, suspensão e faltas justificadas/injustificadas, acrescente os
  textos na constante `ABSENTEISMO_CONTEM`.
- Dias seguidos de atestado = 1 atestado. Grava uma linha **"Absenteísmo"** por atestado: data do
  1º dia, descrição, setor (coluna Setor do Cartão), **dias** (corridos: o Cartão marca sábado e
  domingo do atestado) e se ele **encosta numa folga** (o dia antes do início ou depois do fim é
  DSR, folga, compensado ou feriado na escala da própria pessoa).
- O CSV ganhou 2 colunas no fim (agora são 29): `dias_ausencia` e `emenda_folga`.

### Painel: seção "Absenteísmo (atestados)"
- Seletor de **ciclo** (28 a 27) ou "Últimos 3 meses".
- **Evolução mês a mês** (dias, atestados e pessoas por ciclo).
- **Por gestor / setor / cargo** (dias no ciclo).
- **Encostados em folga**: quantos, em que dia da semana os atestados começam e a lista.
- **Ranking dos 3 meses** com reincidência em vermelho (atestado em 2+ meses ou 3+ atestados).
- **Mapa colaborador × dia** do ciclo.
- **Card do colaborador**: atestados dos 3 meses (início, fim, dias, encostado em folga).
- Respeita gestor, grupo, cargo e colaborador; o filtro de período não se aplica.
- Planilha de 28/08 a 27/09: 31 atestados, 107 dias, 23 pessoas (jul 33, ago 27, set 47 dias);
  19 encostam numa folga; 6 reincidentes. Conferência: 18/18.

## Ajustes no painel (06/10)
- Botão **"Esconder filtros"** no topo: tira o painel da direita (Carregar dados e Filtros) e
  os gráficos ocupam a tela toda; os filtros continuam valendo. A escolha fica guardada no
  navegador.
- Card do colaborador: as tabelas (batidas digitadas, atestados, ocorrências) não ficam mais
  espremidas; no computador cada uma rola sozinha, no celular o card inteiro rola.
- Lista de colaboradores (ORG): saiu o aviso em vermelho de quem não está no ORG.

## Afastamentos pelo relatório do sistema, seções e Pausas (07/10)

### FormatarAfastamentos (nova macro)
- Rode com o relatório "Histórico de Afastamentos" (HRCL006) aberto (aba ativa, numa cópia). A aba
  vira uma tabela e é renomeada para `Afastamentos`: `Matrícula · Colaborador · Admissão · Cód.
  Situação · Situação · Início · Hora início · Término · Hora término · Dias · Horas · Prev.
  Término · Exame`. "Horas" só nos afastamentos de poucas horas (ex.: Saída Médico/Empresa).
- Depois, traga a aba para a planilha de Tratamento.

### GerarCSVPonto
- Com a aba `Afastamentos`, grava uma linha **"Afastamento"** por registro (todos os tipos), com
  dias, horas e se encosta numa folga (o Cartão continua servindo para isso). Sem a aba, os
  atestados continuam vindo do Cartão, como antes.

### Painel
- **Menu de seções** (botão ☰ no topo): o painel mostra uma seção por vez (Visão geral,
  Colaboradores em destaque, Pausas térmicas, Marcações digitadas, Afastamentos, Demora no
  PontoNet, Lista de colaboradores, Histórico) ou "Todas as seções numa página só".
- **Afastamentos — atestados**: agora a partir do relatório do sistema (33 atestados, 105 dias,
  21 pessoas no período de teste). Cada registro do sistema é 1 atestado.
- **Evolução mês a mês**: com um ciclo escolhido no seletor, mostra as semanas desse ciclo e
  destaca o ciclo na tabela (antes não mudava nada).
- Card do colaborador: todos os afastamentos dos 3 meses (atestado, férias, curso, saída médico...).
- **Pausas Térmicas**: a tabela fica na própria seção (não é mais uma janela), com colunas
  estreitas e o cabeçalho em duas linhas, cabendo na tela sem rolar para o lado.
- Conferência: 18/18.

## Revisão do painel com o Dimitri (08/10)
- Tela inicial (sem arquivo) com o passo a passo "Como usar o painel"; o botão "Entenda o
  painel" só aparece com arquivo carregado.
- Colaboradores em destaque: ocorrências curtas mostram 20 pessoas e "Mostrar todos".
- Marcações digitadas: saíram os cards "No dia de uma ocorrência" e "Horário redondo"; o top de
  colaboradores mostra 20 e "Mostrar todos".
- Afastamentos: a lista de atestados encostados em folga mostra o dia da folga (data, dia da
  semana, o que o Cartão diz) e se ela fica antes do início ou depois do fim; o mapa colaborador
  × dia lista todos que tiveram atestado nos 3 meses, com dias no ciclo e em 3 meses.
- Demora no PontoNet: aviso em amarelo quando o gestor (ou o próprio colaborador) esteve de
  férias entre a ocorrência e a tratativa, pelas linhas "Afastamento" com situação "Férias".
- Lista do ORG: o seletor agora é só de cadastro (sem matrícula, sem gestor, sem cargo, sem área).
- Filtro de status: o "Pendente" vinha das linhas-resumo; agora só conta ocorrências, e o
  filtro some quando há um só status.

## Limites de jornada (10/10)
- **Macro nova `FormatarJornada`** (pasta `Jornada/`): uma macro só para os 4 relatórios do sistema
  — Intrajornada (intervalo menor que o mínimo), Interjornada (menos de 11h entre um dia e outro),
  Interjornada semanal (trabalhou no DSR/feriado) e Horas excedentes (mais de 10h no dia). Abra o
  relatório, deixe a aba dele ativa e rode a macro: ela descobre qual é, transforma em tabela e
  renomeia a aba. Copie a aba para a planilha de Tratamento.
- **GerarCSVPonto**: acha as 4 abas pelo cabeçalho e gera uma linha por ocorrência
  (tipo_ocorrencia = "Intrajornada", "Interjornada", "Interjornada Semanal", "Horas Excedentes").
  O CSV passa a ter 32 colunas: `minutos_trabalhados`, `minutos_descanso` e `marcacoes_dia` no fim.
- Painel: por enquanto essas linhas só ficam fora das contagens gerais; a seção própria depende
  das vistas que o RH escolher.

## Subtítulos que acompanham o filtro (11/10)

- Card **Colaboradores envolvidos**: o "de X" embaixo agora é o tamanho da equipe do filtro. Sem
  filtro, todas as pessoas do ORG; com um gestor, "de 22 na equipe de <gestor>"; com vários, "nas
  equipes dos N gestores" (o mesmo para grupo e cargo). Quem tem ocorrência e não está no ORG
  também conta, para o total nunca ficar menor que a parte.
- Seção **Marcações digitadas**: o % de pessoas usa a mesma equipe do filtro.

## Gráficos mais fáceis de ler (13/10)

- **Tela cheia**: botão ⛶ no canto de cada card (Esc ou ✕ fecham).
- **Filtrar dentro do gráfico**: clique num item da legenda para ver só ele (ex.: só "Inclusão por
  esquecimento" nas digitadas); clique em outros para somar; "Mostrar todos" volta. Vale para os
  gráficos da seção; os números do topo e os filtros da lateral não mudam.
- **Período de cada card**: etiqueta amarela 📅 embaixo do título (período do filtro, mês do
  Cartão, relatório de pausas, ciclo ou últimos 3 meses dos afastamentos, cadastro do ORG).
- **Mapa colaborador × dia**: ao passar o mouse aparece a pessoa, o dia e o tipo de afastamento; a
  legenda também filtra.
- Ao abrir, os filtros sempre aparecem; ao carregar o arquivo, o painel vai para a Visão geral.
- Cores: 5 cores bem diferentes (conferidas para daltonismo) e cinza para "Outros"; o vermelho
  ficou só para alerta, falta, banco de horas e atestado.
- Depois (13/10): o filtro da legenda vale só para o gráfico clicado; dá para filtrar em tela
  cheia sem ela fechar; listas longas (curtas, top de digitadas, banco de horas) aparecem inteiras
  só em tela cheia ("Ver todos em tela cheia"); no mapa colaborador × dia a legenda fica em cima, há
  tela cheia, e com filtro some quem não tem nada no ciclo.
- **De onde vem**: clicar numa barra, num pedaço colorido dela, numa fatia, num ponto da linha ou
  num quadrado do mapa abre a lista de pessoas e registros que formam aquele número (ex.: o pedaço
  vermelho da barra de um gestor nos afastamentos mostra quem teve atestado). O nome abre o card.
- **Entenda o painel** reescrito para quem não conhece os processos do RH, sem termos técnicos.
- Na janela "de onde vem", os nomes de cima filtram a lista (o resumo da pessoa abre pelo nome na
  lista) e a janela rola inteira. Revisão geral: topo do celular sem rolagem lateral, botão
  "Esconder filtros" não fica mais laranja no modo explicação, subtítulos sem termos técnicos.

- Correções (13/10): a janela "de onde vem" voltava a não rolar no computador (regra de estilo
  sobrescrita) e o cabeçalho da tabela dela agora fica fixo no topo ao rolar.

## Escolhas da página de ideias construídas (14/10)

As 15 ideias marcadas na página de ideias que ainda faltavam:

- **Jornada (seção nova no menu)**: J1 quatro números (mais de 10h no dia, intervalo de almoço
  curto, trabalho na folga ou feriado, menos de 11h de descanso) e gráfico por ciclo com uma cor por
  tipo; J2 ranking de mais de 10h por pessoa (dias e horas acima de 10h; clique mostra as batidas de
  cada dia); J7 risco trabalhista por gestor (gráfico e tabela com a % da equipe atingida). Seletor
  de ciclo próprio. As horas do dia são o "Total" do relatório (com a conversão da hora noturna).
- **Pegar no pulo (afastamentos)**: P1 atestado sempre no mesmo dia da semana; P2 atestado de 1 dia
  na segunda ou na sexta (por pessoa e por equipe); P3 atestado a até 3 dias das férias.
- **Pegar no pulo (digitadas)**: P10 "Esquecimento crônico": esquecimentos por pessoa em cada ciclo,
  com a tendência quando o arquivo tem mais de um ciclo; vermelho = 5 ou mais no último ciclo.
- **Pausas térmicas**: T1 quatro números no topo; T3 trabalho acima de 1h40 sem pausa por pessoa
  (vermelho acima de 30% dos períodos) e por gestor; T4 pausas certas/curtas/longas por gestor ou
  cargo; T8 pausas no card do colaborador, comparadas com a média da equipe; T9 evolução mês a mês.
- **T9 na macro GerarCSVPonto**: agora lê **todas** as abas de pausas. A aba sem mês no nome é a do
  mês atual; as dos meses anteriores precisam do nome do mês no nome da aba (ex.: "Pausas Térmicas
  Agosto"). Duas abas do mesmo mês: só a primeira entra (a mensagem final avisa).
- **M3**: no gráfico por gestor, o seletor do título troca "total" por "por pessoa da equipe"
  (ocorrências ÷ pessoas da equipe na lista de equipes).
- **M4**: na janela "de onde vem" aparecem botões "Filtrar o painel: gestor X / tipo Y / dia Z"
  (o clique na barra continua abrindo "de onde vem", como decidido em 13/10).
- **M7**: porcentagem do total ao lado do número nas barras e na legenda das roscas.
- Conferência: 26 itens (os 18 de antes + jornada, sinais de atestado, esquecimento e pausas).

## Tendências e 3 meses de digitadas e de pausas (18/10)

- **Seção nova "Tendências"** (logo depois da Visão geral): a unidade ciclo a ciclo (batidas
  digitadas, esquecimentos, dias de atestado, os 4 limites de jornada e as pausas), com barrinhas
  dos 3 ciclos e seta vermelha (piorou) ou verde (melhorou); embaixo, o assunto escolhido por
  gestor e as pessoas que mais pioraram e mais melhoraram. Segue os filtros de gestor, grupo,
  cargo e colaborador: com um gestor escolhido, é a tendência da equipe dele.
- **Em cada seção**: digitadas ciclo a ciclo por motivo e por pessoa; pausas (acima de 1h40) por
  gestor e por pessoa; dias de atestado por gestor; jornada por gestor e mais de 10h por pessoa;
  "Extra e falta no mesmo dia" com a tendência (este mês x média dos 2 anteriores).
- **Card do colaborador**: "Evolução da pessoa, ciclo a ciclo".
- O painel continua abrindo no mês atual: as digitadas dos meses anteriores só entram nas
  tendências (e quando se aumenta o período no filtro).
- Texto "pegar no pulo" retirado dos títulos (pedido do Dimitri).
- **FormatarPausasTermicas**: renomeia a aba com o mês do relatório ("Pausas Térmicas
  Setembro"), lido da linha "Período" antes de apagar o cabeçalho.
- Ferramentas: `formatar_pausas.py` (versão Python da macro) e `adicionar_aba_xlsx.py
  --substituir` (troca o conteúdo de uma aba sem regravar o resto da planilha).
- Conferência: 27 itens (o novo confere a tabela da unidade ciclo a ciclo com as abas).

## Pausas térmicas: batidas dia a dia (18/10)

- **FormatarPausasTermicas** agora deixa três tabelas na aba: totais por colaborador (como antes),
  Total Geral (como antes) e, embaixo, **as batidas** (uma linha por batida: colaborador, data, dia
  da semana, hora, nº da batida no dia, tipo — entrada/saída, refeição ou pausa térmica — e as
  exceções do dia), como Tabela do Excel com filtro próprio. A GerarCSVPonto continua lendo só a
  tabela de totais, então o CSV e o painel não mudam.
- Letras do relatório: P = entrada/saída da jornada, G = refeição, F = pausa térmica.

### Depois: as visões das batidas no painel (B1, B2, B5, B6, B7, B9, B12)

- **GerarCSVPonto** lê também a tabela de batidas de cada aba de pausas (`GerarDiasPausas`) e grava
  uma linha **"Pausa Térmica Dia"** por pessoa e dia: `situacao` = exceções do dia,
  `marcacoes_dia` = as batidas com a letra ("21:02 P 21:30 F 21:50 F ..."). Na planilha de 3 meses:
  3.002 dias (958 / 1.105 / 939 por ciclo); CSV com 6.084 linhas.
- **Painel, seção Pausas Térmicas**, bloco "As batidas de cada dia" com seletor do mês:
  horário em que mais se passa de 1h40 (B1), dia da semana (B2), tempo até a 1ª pausa (B7), dias
  com mais de 1h40 e nenhuma pausa (B5), maior tempo seguido sem pausa por pessoa (B6), qual
  batida fica faltando (B12) e o mapa colaborador × dia (B9). Clique em barra ou quadrado abre o
  "de onde vem" com as batidas do dia em palavras.
- A conta é refeita batida a batida (trabalho e pausa alternando; acima de 100 min = acima de
  1h40) e pode diferir em poucos períodos do total do sistema (setembro: 284 × 288).
- Conferência: 32/32 (5 itens novos, recalculados da tabela de batidas da aba).

## Ajustes pedidos em 19/10

- Tendências: o minigráfico "Evolução (mês a mês)" mostra o mês embaixo de cada barra, o número
  ao passar o mouse e abre os registros daquele mês ao clicar. A coluna "Comparado ao mês
  anterior" diz em palavras o que mudou ("17 a menos que em Ago/26"; "subiu de 3,7% para 9,2%").
- "Extra e falta no mesmo dia" sem a coluna de tendência, ao lado de "Colaboradores com mais
  ocorrências".
- Afastamentos: "Doação de sangue" virou um tipo próprio, e o seletor "Mostrar" tem "Só <tipo>"
  para cada tipo do arquivo.
- Legenda clicável (filtro) no mapa das pausas, no "Pausas por gestor/cargo" e na lista de jornada.
- Textos da tela revistos para quem não é da área (sem "DSR", "ORG", "#N/A", nomes de macro/aba).

## Ajustes pedidos em 26/10

- Seções na ordem: Ocorrências, Marcações digitadas, Afastamentos, Jornada, Colaboradores em
  destaque, Tendências, Pausas térmicas, PontoNet, Lista de colaboradores, Histórico.
- Botão "?" em cada gráfico: mostra a explicação só daquele gráfico.
- "Demora média no PontoNet" foi para a seção PontoNet.
- Tendências e Pausas: o seletor ("Assunto", "Mês das batidas") fica numa caixa com os gráficos que
  dependem dele.
- Sem nome do arquivo no topo; sem menu de seções antes de carregar o arquivo; sem status
  (Regularizado/Pendente); sem a visão "Qual batida fica faltando".

## Ajustes pedidos em 27/10

- Topo: ☰ seções, logo, período e o nome do painel; o botão "Entenda o painel" saiu (cada gráfico
  tem o seu "?" e cada seção tem "Entenda esta seção").
- PontoNet com 3 meses: a aba "Pontonet" pode ter o relatório "Ausência de marcação" de 3 meses.
  O GerarCSVPonto gera uma linha "Justificativa PontoNet" por pedido (CSV com 37 colunas) e a seção
  PontoNet mostra os pedidos mês a mês, como terminaram, o motivo, o tempo do colaborador e do
  gestor (os dias que o PontoNet conta) e a tendência por gestor. Corrigido: um pedido cancelado
  não apaga mais a justificativa aprovada do mesmo dia.
- Mesma ordem dos gráficos em todas as seções: números do topo, mês a mês, o quê, quando, por
  gestor, por pessoa, mapa, sinais de atenção, tendências e lista completa.

## Ajustes pedidos em 28/10

- PontoNet: só os pedidos aprovados (cancelados e batidas achadas no relógio não entram).
- A demora (do dia da batida que faltou até a aprovação) agora segue o mês escolhido na seção.
- Topo: seções e logo à esquerda, período no meio, nome do painel e "Esconder filtros" à direita.
- Menu: Tendências depois de PontoNet.
- Dá para abrir vários arquivos juntos (ex.: o de dois gestores) e juntar mais um depois; o filtro de
  gestor mostra todos. Linhas repetidas entre os arquivos (ex.: TODOS + um gestor) contam uma vez.

## Ajustes pedidos em 29/10

- **PURO → EDITADO cola nas abas de tratamento.** Depois de montar a `EDITADO_<MES><AA>`, a macro
  procura cada pessoa no ORG (pela matrícula; se não achar, pelo nome) e olha a coluna
  "Tipo de trabalho": internos vão para `Tratamento - Internos` e externos para `Tratamento -
  Externos`, coluna por coluna pelo nome do cabeçalho (Data ← Dt.Apuraç., Status ← Tratamento,
  Pontonet só na de externos). Trabalho, Cargo e Gestor: se a aba já tem fórmula (PROCX), a
  fórmula é estendida; se não, o valor vem do ORG. A macro é usada todo dia: **nunca apaga nada**
  das abas de tratamento, só cola embaixo. Pergunta onde colar: 1 = internos e externos, 2 = só
  internos, 3 = só externos, Cancelar = nada (30/10). Checa repetidas por pessoa + dia (30/10).
  **SIMPLIFICADA em 04/11 (pedido do Dimitri):** a macro só cria a aba `EDITADO_<MÊS><AA>` e joga
  as linhas nas abas `Tratamento - Internos` / `Tratamento - Externos` que já existem, embaixo do
  que já está lá. Só pergunta onde colar (1 = internos e externos, 2 = só internos, 3 = só externos; escolha de 04/11), não compara se o dia já foi colado, não cria aba de
  relatório nem cópia "(2)" (se a EDITADO do mês já existe, é apagada e refeita). Quem não está no
  ORG (ou está sem "Tipo de trabalho") fica amarelo na EDITADO, não é colado, e a mensagem final diz
  quantos são. `DistribuirEditadoNasAbasTratamento` refaz só a colagem (com a aba EDITADO aberta).
- Painel: o filtro de período virou um calendário próprio em que **só os dias com informação**
  podem ser clicados.
- Topo sem o nome "Painel de Ponto · Ocorrências dos Colaboradores".
- Todo número do topo das seções mostra o período a que se refere (📅) e tem o seu "?".
- Afastamentos: o resumo do ciclo virou frases simples numa caixa.
- Colaboradores em destaque: busca por nome no banco de horas e dois gráficos novos, "quem ficou
  mais no positivo" e "quem ficou mais no negativo" no saldo do mês (hora extra − banco).
- Pausas térmicas: "período de trabalho" virou "trecho de trabalho" (o tempo entre uma pausa e a
  próxima), com a explicação; cabeçalhos da tabela em palavras simples.
- PontoNet: o card "Demora média" saiu; a média de todos virou uma linha no card por gestor e as
  tabelas de demora vêm logo depois da lista de justificativas aprovadas.
- "De onde vem" revisado em todos os gráficos: o pedaço "Curtas" (ou "Longas", "Certas") das
  pausas por gestor abre só as pessoas que tiveram aquele tipo, com a coluna em destaque; as barras
  de afastamento em dias contam dias na janela; o gráfico de atestados por dia da semana separa os
  que encostam numa folga (vermelho) dos que não encostam (cinza).

## Gerar os CSVs a partir de uma pasta (02/11)

- **GerarCSVsDaPasta** (no módulo PrepararAbasDoMes): na planilha de tratamento do mês, rode esta
  macro. Ela pede (1) a pasta onde estão os relatórios baixados do sistema e (2) a pasta onde os
  CSVs vão ser gravados. Depois abre cada arquivo da pasta (sem alterar os arquivos), traz as abas
  para a planilha, formata tudo, gera a aba CSV e grava `dados_<gestor>.csv` e `dados_TODOS.csv` na
  pasta escolhida. A aba "Relatório da preparação" mostra o que aconteceu com cada arquivo.
- Pode rodar de novo no mesmo mês: as abas já formatadas do mesmo relatório (ex.: "Marcações
  Digitadas", "Cartão Outubro", "Pausas Térmicas Setembro") são substituídas, sem duplicar.
- A RE nova pode ir na pasta (entra como "RE dd.mm" e passa a ser a usada). ORG e as abas
  Tratamento - Internos/Externos continuam na planilha.
- Arquivo que não é nenhum dos relatórios esperados é ignorado e aparece no relatório.

## Uma macro só para preparar o mês (31/10)

- **PrepararAbasDoMes** (nova): copie para a planilha de Tratamento as abas dos relatórios baixados
  do sistema (Cartão Ponto de cada mês, Pausas Térmicas, Marcações digitadas, Afastamentos, os 4 de
  jornada e o PontoNet), com qualquer nome, e rode só esta macro. Ela reconhece cada relatório pelo
  conteúdo, formata com a macro dele (sem as mensagens de cada uma) e cria a aba "Relatório da
  preparação" com o que foi feito em cada aba, o que deu certo, o que deu erro e o que faltou baixar.
  Precisa dos módulos de todas as macros colados na planilha (as macros de formatação ganharam um
  "modo automático"; usadas sozinhas continuam iguais).
- Pausas Térmicas de vários meses num relatório só: sai uma aba por mês, com os totais recontados
  das batidas de cada mês.
- GerarCSVPonto: acha a RE também quando o cabeçalho vem "Matricula" sem acento (RE 23.10).
- Painel: o período padrão termina no fim do espelho de ponto (dia 27), mesmo que a planilha de
  tratamento traga ocorrências do dia 28 (que já são do ciclo seguinte).

## Ordem de uso no mês

1. Rodar `ConsolidarCartaoPonto` no HRCP102 e colar o resultado numa aba `Cartão <Mês>`.
2. Rodar `FormatarPausasTermicas` no HRES114 (numa cópia) e trazer a aba para a pasta de Tratamento.
3. Rodar `PURO_Para_EDITADO` (ele já cola os internos e os externos nas abas `Tratamento - Internos` /
   `Tratamento - Externos`, pelo "Tipo de trabalho" do ORG) e tratar as ocorrências nessas abas.
4. Rodar `FormatarMarcacoesDigitadas` no relatório de Marcações digitadas e
   `FormatarAfastamentos` no Histórico de Afastamentos, e trazer as duas abas.
5. Rodar `GerarAbaCSV` e depois `ExportarCSVPorGestor`.

> Estes macros foram escritos e revisados fora do Excel: a lógica foi conferida simulando-a em
> Python sobre os arquivos reais. Rode sempre numa **cópia** da planilha na primeira vez.

## Banco de horas e CSVs com o nome do gestor (04/11)

- **Arquivo novo na pasta do mês: `BH`** (relatório de banco de horas HRES105 do sistema). A
  `GerarCSVsDaPasta` agora confere **12** arquivos (os 11 de antes + `BH`) e para, sem mexer em nada,
  se faltar algum.
- **Macro nova `FormatarBancoHoras`** (pasta `BancoHoras/`, também dentro do módulo único): reescreve
  o relatório como tabela, uma linha por pessoa: matrícula, nome, filial, local, início e fim do
  período, quantos meses (ciclos de 28 a 27) o período tem, se o banco fechou, saldo anterior,
  crédito, débito, horas anistiadas (perdoadas), horas pagas e saldo final. A aba vira `Banco de
  Horas`. A `PrepararAbasDoMes`/`GerarCSVsDaPasta` reconhece o relatório sozinha.
- **CSV com 44 colunas**: linhas novas `Banco de Horas Período` (uma por pessoa) com as colunas
  `bh_inicio_periodo`, `bh_saldo_anterior_minutos`, `bh_credito_minutos`, `bh_debito_minutos`,
  `bh_anistiadas_minutos`, `bh_pagas_minutos`, `bh_saldo_final_minutos`.
- **Nomes dos arquivos**: `<GESTOR>.csv` (antes `dados_<gestor>.csv`) e `TODOS OS GESTORES.csv`
  (antes `dados_TODOS.csv`).
- **Painel, seção "Colaboradores em destaque"**, depois do saldo do mês: card **"Como cada pessoa
  fechou o banco de horas"**. O banco fecha de 4 em 4 meses: o painel conta os meses do período do
  relatório e, com 4, mostra o banco **fechado** (quem teve horas pagas, perdoadas ou fechou zerado,
  com crédito e débito de cada pessoa); com menos de 4, mostra o banco **aberto** e o saldo até agora,
  avisando que ainda pode mudar. Seletor "Mostrar", busca por nome, tela cheia e o resultado também
  no card do colaborador.
- Relatório de 28/06 a 27/10 (216 pessoas): crédito 4542h47, débito 4369h44; 65 pessoas com horas
  pagas (196h15), 38 com horas perdoadas (23h12), 113 zeradas. Bate com o "Total Geral" do relatório.
- Ferramentas: `formatar_bh.py` (o mesmo da macro, em Python), simulador com as linhas de banco de
  horas (`BH_PKL`), conferência com 34 itens (o item de digitadas e os da jornada foram ajustados às
  mudanças de 03/11).

### Correção (04/11, depois): erro "Variável não definida" no código único

- O código único dava **"Erro de compilação: Variável não definida"** em `substituirAbas` ao rodar a
  GerarCSVsDaPasta. Causa: ao juntar as macros, as declarações de cada parte (`Private ... As`,
  `Const`) ficavam no meio do módulo, onde cada parte começava; o VBA só aceita declaração antes da
  primeira Sub. Agora todas ficam no topo e o script de montagem confere.
- A **PURO_Para_EDITADO saiu do código único**: é outro assunto (dia a dia do tratamento) e fica só no
  `txt/PURO_Para_EDITADO.txt`, colado num módulo separado.

## Banco de horas dia a dia, períodos fixos e PontoNet só do ORG (05/11)

- **Seção nova "Banco de horas"** (depois de "Colaboradores em destaque"): escolha a equipe, a pessoa e os
  meses; o gráfico mostra o saldo do banco no fim de cada dia (a soma da equipe, ou a pessoa), com o
  fechamento marcado. Embaixo, uma tabela por pessoa (saldo no começo e no fim, crédito, débito, menor
  saldo, dias com extra e com débito, folgas no banco e um resumo em frase), os cards de crédito e débito
  por colaborador (vieram de "Colaboradores em destaque") e o fechamento do banco.
- O saldo parte do relatório de banco de horas (HRES105) e anda dia a dia com o espelho de ponto:
  **crédito = hora extra 50%**, **débito = coluna BH** (a 100% é paga à parte). Conferido com a
  "Análise de compensação do BH" do RH.
- **Períodos do banco**: novembro a fevereiro, março a junho e julho a outubro; fecha no último mês. O
  relatório de banco de horas pode ter 1 mês, 2, 3, o período ou o ano: o painel diz se ele inclui um
  fechamento. No fechamento, o saldo negativo (coluna "Hr. Anist.") aparece como **descontado na folha**.
- CSV: linhas novas **"BH Dia"** (uma por pessoa e dia com movimento no banco; colunas 24/25).
- **PontoNet só dos gestores que estão no ORG** (macro e painel); as equipes deles também saem.
- Botões **"?"** em laranja.


## Arquivos carregados e banco de horas a partir do zero (05/11, depois)

- **Seção nova "Arquivos carregados"** (a primeira do menu, antes de Ocorrências): para cada CSV aberto,
  que informação do sistema veio nele (espelho de ponto, BH, pausas, digitadas, afastamentos, PontoNet,
  jornada, lista de equipes), com o **período** e os **meses** (28 a 27); e, somando os arquivos, o que
  **cada seção usa** e o que **falta**, com o nome do arquivo da pasta do mês (ex.: "Falta Jul/26 em
  Digitadas 3 meses", "Falta o espelho de ponto de Jul/26 (arquivo Cartão ponto)").
- **Banco de horas a partir do zero**: no começo do período (nov, mar, jul) todo mundo está zerado. Com o
  espelho de ponto de **todos os meses do período** no arquivo, o saldo é só a soma dos dias; o relatório
  de banco de horas (um só: o do último mês fechado do período) serve para **conferir**. Card novo
  **"Conferência: espelho de ponto × relatório de banco de horas"**: crédito, débito e saldo de cada pessoa
  pelos dois lados; diferença em vermelho. Faltando um mês do espelho, o painel diz qual e o saldo volta a
  partir do relatório.
- Dá para abrir junto o CSV do mês anterior para completar os meses: cada mês do espelho vem de um
  arquivo só (o mais novo que tem aquele mês), para um dia que o sistema mudou não contar duas vezes.
- **GerarCSVPonto**: as linhas "BH Dia" vêm dos **4** Cartões mais recentes (antes 3). No último mês do
  período (fevereiro, junho, outubro), coloque na pasta também o **"Cartão ponto 4"** (o 1º mês do
  período); a pasta não exige esse arquivo, mas sem ele o painel avisa que falta o mês.

## Pedidos de 05/11 (foto): topo, banco de horas e janelas de detalhe

- Topo **sem o período**; o **logo da Aurora** fica no meio. Ao abrir os arquivos, o painel vai para
  **"Arquivos carregados"**.
- Em toda a tela, **"cartão ponto"** no lugar de "espelho de ponto".
- Cada seção diz **quem está sendo mostrado**, conforme os filtros ("👥 Mostrando: 22 pessoas na equipe de
  FULANO", "34 pessoas nas equipes dos 2 gestores").
- **Banco de horas:** "horas extras" e "horas falta" (sem crédito/débito); sem a conta horas extras − horas
  falta (saíram os cards "Quem colocou/tirou mais" e o saldo à direita do gráfico por colaborador); no
  fechamento, só "fecharam devendo" (sem falar em desconto); os textos dizem se você está vendo o **período
  inteiro** ou só parte dele; os **meses escolhidos são sempre seguidos**; o gráfico tem uma **linha no fim de
  cada mês** (a do fechamento em marrom); com um mês só aparecem **todos os dias**, e **clicar num dia** abre o
  resumo dele (quem fez hora extra, quem teve hora falta e o banco de cada um no fim do dia).
- **Janelas de detalhe** ("de onde vem" e o card da pessoa): nenhum texto cortado, barras de rolagem visíveis;
  os nomes de cima ficam numa caixa com busca e "Ver todos os nomes" (antes, com muita gente, ocupavam a janela).
