# CONTEXTO DO PROJETO — Tratamento de Ponto e Painel de Ocorrências (AuroraCoop)

> Arquivo de contexto para o Claude Code. Estado em **08/10/2026**. Leia inteiro antes de propor
> qualquer mudança: várias regras abaixo já foram discutidas e decididas com o Dimitri, e as
> "armadilhas" já custaram retrabalho.
>
> **Chegou agora, sem histórico?** Comece pela seção **12 (Estado atual e como continuar)**: ela diz
> onde está o código, o que o Dimitri está testando, o que falta decidir e como trabalhar com ele.

---

## 1. Quem e o quê

**Dimitri** (RH, AuroraCoop — Cooperativa Central Aurora Alimentos, unidades de Recife/Cabo e
Teresina) trata as ocorrências de cartão ponto dos colaboradores todo mês e entrega um **painel
(index.html)** para os **gestores** e o **supervisor administrativo**. O ciclo de apuração vai do
dia **28 de um mês ao dia 27 do mês seguinte**.

Fluxo mensal:

```
HRCP102 (cartão ponto, sistema Senior) ──► ConsolidarCartaoPonto ──► aba "Cartão <mês>"
HRES114 (pausas térmicas, Senior)      ──► FormatarPausasTermicas ──► aba "Pausas Térmicas"
Marcações digitadas (Senior)           ──► FormatarMarcacoesDigitadas ──► aba "Marcações Digitadas"
HRCL006 Histórico de Afastamentos      ──► FormatarAfastamentos   ──► aba "Afastamentos"
aba PURO (exportação de ocorrências)   ──► PURO_Para_EDITADO      ──► aba EDITADO_MESAA ─► (Dimitri trata) ─► "Tratamento - Internos" / "Tratamento - Externos"
                                                         todas as abas ──► GerarAbaCSV ──► aba "CSV" ──► ExportarCSVPorGestor ──► dados_<gestor>.csv + dados_TODOS.csv
                                                                                                                                  └──► index.html (abre no navegador)
```

## 2. Como falar com o Dimitri

- **Sempre em português (pt-BR).**
- **Códigos VBA são entregues SEMPRE em `.txt`** (pedido explícito dele). Ele cola no editor do
  VBA. Mantenha também o `.bas` no repositório.
- Ele quer entender o **porquê**: explique o raciocínio, não só entregue o arquivo.
- Valoriza diagnóstico honesto: quando algo entregue estava errado, diga direto.
- Pensa junto e muda de ideia: confirme decisões de negócio antes de construir.

## 3. Restrições (TI da empresa)

- Gestores **não podem instalar nada**; não há servidor, banco nem hospedagem.
- O painel é **um único HTML autocontido**: sem CDN, sem bibliotecas; gráficos em SVG feitos à mão.
  O CSV é lido com `<input type="file">` + `FileReader` (`fetch` não funciona em `file://`).
- Gestores usam Chrome/Edge, no desktop e no celular (o layout precisa ser responsivo).
- Dimitri tem Excel 365.

## 4. Arquivos desta pasta

| Caminho | O que é |
|---|---|
| `CartaoPonto/ConsolidarCartaoPonto.bas` | HRCP102 → aba `Cartao_Consolidado` (1 linha por colaborador/dia) |
| `PausasTermicas/FormatarPausasTermicas.bas` | HRES114 → a mesma aba, só com a tabela de colaboradores e a de total |
| `PuroParaEditado/PURO_Para_EDITADO.bas` | aba `PURO` → `EDITADO_MESAA` (reordena, X→SIM, tira espaços) |
| `MarcacoesDigitadas/FormatarMarcacoesDigitadas.bas` | relatório "Marcações digitadas" → a mesma aba como tabela (renomeada "Marcações Digitadas") |
| `Afastamentos/FormatarAfastamentos.bas` | relatório HRCL006 "Histórico de Afastamentos" → a mesma aba como tabela (renomeada "Afastamentos") |
| `GerarCSVPonto/GerarCSVPonto.bas` | `GerarAbaCSV` (monta a aba CSV) e `ExportarCSVPorGestor` (grava os .csv) |
| `Dashboard/index.html` | O painel. Contém o logo em base64 (linha enorme, ~60 KB); não leia o arquivo inteiro de uma vez |
| `txt/*.txt` | As mesmas macros em `.txt`, CRLF, prontas para colar |
| `ferramentas/` | Scripts para conferir CSV e painel sem Excel, formatar os relatórios em Python, inserir abas na planilha e checar colisões de nome no VBA (ver seção 10) |
| `README.md` | Guia de uso e histórico de mudanças para humanos |

Os `.bas` estão em **UTF-8**. **Arquivo > Importar** do VBA lê ANSI e estraga os acentos (e com
isso quebra a busca de cabeçalhos como "Matrícula" e "Situação"). Instrução para o Dimitri: colar o
conteúdo num módulo.

## 5. As macros em detalhe

### 5.1 ConsolidarCartaoPonto (aba ativa = HRCP102 bruto)
- O relatório é um "dump" de impressão: um bloco por colaborador com as linhas `Empregado:` (B =
  matrícula, D = nome), `Cargo:` (B), `Localização:` (B = código, **D = setor**) e `Horários:` (B =
  código, C = descrição; **pode ter várias linhas**, as seguintes com a coluna A vazia), depois a
  linha `DT` e os dias até `Horas Normais:`.
- Linha de dia: A = dia (data serial com formato `dd`), B = Sem (SEG..DOM), **C = Hor (código do
  horário do dia)**, D = marcações + texto ("08:00 12:00 13:12 18:00    Trabalhando"),
  G..O = Trabalho, BH, Just., Injus., 50%, 60%, 100%, 120%, Ad. Not.
- Saída: cabeçalho na **linha 2** (a linha 1 tem os rótulos de grupo). Colunas, posições definidas
  em constantes `COL_*`:
  `A Matrícula · B Nome · C Cargo · D Setor · E DT · F Sem · G Cód. Horário · H Horário · I..N Ponto 1..6 ·
  O Descrição Marcação · P Trabalho · Q BH · R Just. · S Injus. · T 50% · U 60% · V 100% · W 120% ·
  X Ad. Not. · Y Banco de horas e extra no mesmo dia · Z Ponto faltando · AA Mais de 4 pontos ·
  AB Tolerância < 15min · AC Batidas < 30min`
- "Horário" é a descrição **do código daquele dia** (lista "Horários:" do colaborador). Os códigos
  9999 (DSR), 9998 (Compensado), 9997 (Feriado) e 9996 não estão na lista e ficam em branco.
  No HRCP102 de 28/08 a 29/09, 76 dos 202 colaboradores têm mais de um horário no período.
- A data é resolvida pelo dia + dia da semana dentro do período (`ResolverDataPorDiaSemana`).
- ⚠️ As abas "Cartão Julho/Agosto/atual" das planilhas de agosto e setembro ainda estão no
  **layout antigo de 28 colunas** (A Matrícula, B Nome, C Cargo, D Horário, E Setor, F DT, G Sem,
  H..M Pontos...). O GerarCSV lê por **nome de cabeçalho**, então funciona com os dois layouts.

### 5.2 FormatarPausasTermicas (aba ativa = HRES114 bruto)
- Reescreve a aba deixando **só**: a tabela de colaboradores (cabeçalho na linha 1: `Matrícula,
  Colaborador, Cargo, Pausa corretas, Pausa menor 0:20, Pausa maior 0:20, Trabalho correto 1:40,
  Trabalho maior 1:40, Trabalho menor 1:40, Marcações Ímpares`), **uma linha em branco** e a
  tabela "Total Geral". Sem título, legenda nem "Soma dos itens" (pedido do Dimitri).
- Não recalcula nada: copia o bloco "Total do Colaborador" (valores 2 linhas abaixo do rótulo, nas
  colunas A, B, D, F, H, J, K) e o "Total Geral".
- A linha em branco é o que marca o fim da tabela para o GerarCSV.

### 5.3 PURO_Para_EDITADO
- Mapeia a aba `PURO` por nome de cabeçalho → `EDITADO_<MES><AA>` com as colunas `Matricula, Nome,
  Gestor, Situação, Tratamento, Pontonet, Dt.Apuraç., Marcações, Extras, Faltas, Extra 100%, Qtd.
  Horas Pagas, Intra, Inter, Marcações Impares, Horas Excedentes, Extras e Faltas, Refeição fora de Escala`.
- `LimparTexto`: tira os espaços das pontas de **todo texto** (o sistema completa com espaços até
  uma largura fixa); célula só com espaços vira vazia; datas, horas e números ficam intactos.

### 5.3b FormatarMarcacoesDigitadas (aba ativa = relatório bruto)
- Dump de impressão: cabeçalho de página repetido ("Crachá | Colaborador | Cargo | Local | Ori. |
  Data | Hora | Coletor | Função | Justificativa") com os valores **deslocados**: nome = coluna de
  "Colaborador" + 1; cargo = código (col. "Cargo") + descrição (+1); motivo = col. "Justificativa";
  texto livre = col. "Justificativa" + 2. Crachá = matrícula. Data e hora em colunas separadas.
- Motivo longo quebra em duas linhas: a linha seguinte só tem o resto na coluna do motivo
  ("ponto"); a macro junta. Lê com `.Value2` (com `.Value`, data vira Date e `IsNumeric` dá False).
- Saída (linha 1): `Matrícula, Colaborador, Cód. Cargo, Cargo, Cód. Local, Local, Origem, Data,
  Hora, Dia da semana, Coletor, Função, Motivo, Justificativa`.
- Relatório de 28/08 a 27/09: 434 batidas, 137 pessoas; motivos: atividade externa 167,
  dispositivo 157, esquecimento 109, exclusão por erro 1; Origem D 422 / E 12. A macro mantém
  todas na aba; **o GerarCSV ignora Origem "E"** (decisão do Dimitri, 06/10) e o painel também
  descarta "E" ao ler um CSV antigo. O relatório **não diz quem digitou** (colaborador × RH) — não
  há como separar.

### 5.3c FormatarAfastamentos (aba ativa = HRCL006 bruto)
- Dump de impressão: por colaborador, cabeçalho "Tipo | Colaborador | Admissão" e a linha da pessoa
  (tipo **como texto** "1", matrícula na col. de "Colaborador", nome +1, admissão); depois o
  cabeçalho "Afastamento | Situação | Dem. | Dias | Faltas | Dt.Término | Prev Término | Exame" e
  uma linha por afastamento com valores **deslocados**: início = col. "Afastamento", hora início
  +1; código = "Situação"+1, descrição +2; término = "Dt.Término", hora término +1; prev. término
  = "Prev Término"+1; exame = "Exame"+1. "Dem./Dias/Faltas" vêm sempre 00 (não usados).
- Saída: `Matrícula, Colaborador, Admissão, Cód. Situação, Situação, Início, Hora início, Término,
  Hora término, Dias (corridos), Horas (só afastamento de poucas horas no mesmo dia), Prev.
  Término, Exame`.
- Relatório de 07/10 (inícios de 30/06 a 26/09): 283 afastamentos, 117 pessoas, 16 tipos (Viagem
  a Serviço 80, Férias 55, Atestado 33, Reunião 26, Saída Médico/Empresa 26 por horas, Curso 25,
  Horas Faltas 10, Treinamento 9, Dispensa Remunerada 8, Lic. Falecimento 4, Acidente 2, outros 1).
  Atestados: 33 registros, 105 dias, 21 pessoas (o Cartão dava 31 "episódios": o sistema registra
  atestados emendados separados; alguns do Cartão são de desligados ou começam antes de 30/06).
  **O relatório do sistema é a fonte oficial** (decisão do Dimitri, 07/10).

### 5.4 GerarCSVPonto
**Entradas (achadas por nome/cabeçalho, nunca por posição):**
- Abas cujo nome começa com **"Tratamento"** e que têm a coluna `Matricula` na linha 1:
  `Tratamento - Internos` / `Tratamento - Externos` (planilha nova) ou `Tratamento` (antiga).
  Cabeçalhos: `Check, Matricula, Nome, Cargo, Gestor, Situação, Status, Pontonet, Data, Marcações,
  Extras, Faltas, Extra 100%, Qtd. Horas Pagas, Intra, Inter, Marcações Impares, Horas Excedentes,
  Extras e Faltas, Refeição fora de Escala` (Gestor e Cargo são fórmulas XLOOKUP e podem dar `#N/A`).
- **"RE ..."** (ex.: `RE 08.09`, `RE 30.09`): o cadastro (Matrícula → `Nome Unidade` = setor, `Cargo`).
- Abas começando com **"Cartão"** com `Matrícula` e `DT` na linha 2: até as **3 mais recentes**
  (pela maior data). A mais recente é o "mês atual".
- `PontoNet` (opcional; se faltar, abre um seletor de arquivo): `Data falta, Matrícula, Situação
  atual, Justificativa, Avaliado em:`.
- A tabela de pausas, achada em qualquer aba pela linha que tem `Matrícula` + `Pausa corretas`.
- **Marcações Digitadas**: aba com `Matrícula, Data, Hora, Motivo, Justificativa` na linha 1.
- **Afastamentos**: aba com `Matrícula, Situação, Início, Término` na linha 1.
- **ORG** (relação colaborador × gestor): aba que começa com "ORG" com `Matricula`, `COLABORADOR`,
  `Descrição`, `Cargo`, `Unidade`, `GESTOR` na linha 1 (busca de cabeçalho sem acento/maiúscula,
  `ColunaSemAcento`). A coluna Matricula do ORG é fórmula e às vezes dá `#N/A`.

**Regras de negócio (decididas com o Dimitri):**
1. **Check = "S" → a linha da Tratamento é ignorada por inteiro**, inclusive a hora extra 100%.
2. Situação vazia ou **"Sem alteração"** → não é ocorrência.
3. **"Problema horário"** (com ou sem acento) → não entra (`SITUACOES_IGNORAR`).
4. Cada linha válida da Tratamento gera **exatamente 1 linha de ocorrência** (tipo = Situação;
   duração = −minutos de falta, ou 0 se não houver falta). Se houver hora extra, gera **também**
   uma linha "Hora Extra" (e "Hora Extra 100%" se o Cartão tiver 100%).
5. As colunas Extras/Faltas da Tratamento só decidem **se** há extra/falta; a **quantidade** vem do
   Cartão do mês atual (BH para falta, 50% para extra), com a Tratamento como reserva.
6. Situação final = justificativa do PontoNet quando existir; `data_tratativa_pontonet` = "Avaliado
   em:" só se a situação for `INTEGRADO`.
7. Status: `Tratado` → "Regularizado"; o resto → "Pendente".
8. Linhas de resumo geradas depois do laço:
   - "Auditoria Extra e Falta (3 Meses)": por matrícula, conta os "Sim" de "Banco de horas e extra
     no mesmo dia" em cada um dos 3 Cartões; o painel alerta a partir de **10** na soma.
   - "Falta < 15min" / "Extra < 15min": direto da coluna "Tolerância < 15min" do Cartão atual
     ("Extra e Falta" gera 2 linhas). **Não dependem do Check da Tratamento**: está em aberto com o
     Dimitri se um dia com Check "S" deve sair dessa lista.
   - "Resumo Pausas Térmicas": uma por colaborador da tabela de pausas.
   - "Resumo Horas Cartão" (06/10): uma por matrícula do Cartão atual com o total do mês de
     hora extra (50% + 100%) e de banco de horas (BH). Sem Tratamento, sem Check, sem período.
     No Cartão, BH é sempre positivo e é **débito** (dias de "Falta (Banco Horas)" e minutos de
     atraso em dias "Trabalhando").
   - "Marcação Digitada": uma por batida da aba Marcações Digitadas **com Origem "D"** (as "E"
     são ignoradas); data = dia + hora (minuto);
     situacao = motivo; colunas 26/27 = justificativa / origem. Sem Tratamento nem Check.
   - "Afastamento" (07/10): um por registro da aba Afastamentos (todos os tipos); data = início
     (+ hora início); situacao = Situação; duracao_minutos = Horas (afastamento de poucas horas);
     dias_ausencia = Dias; emenda_folga = `CalcularEmendaFolga` com os dias do Cartão
     (`MontarDiasCartao`). Com a aba, **substitui** as linhas "Absenteísmo".
   - "Absenteísmo" (06/10, agora só sem a aba Afastamentos): um por ATESTADO nos até 3 Cartões (dias seguidos com descrição que
     contém `ABSENTEISMO_CONTEM` = "atest" = 1 atestado; dias corridos, o Cartão marca sábado e
     domingo). data = 1º dia; situacao = descrição; setor = coluna Setor do Cartão; colunas 28/29 =
     dias_ausencia / emenda_folga ("antes (DSR)", "depois (Feriado)"... se o dia antes/depois é
     `FOLGA_CONTEM`). Só atestados por decisão do Dimitri; a lista completa de absenteísmo dele
     (print da Jussara): suspensão, acidente trab. 15, acidente trabalho, atest. odontológico,
     atest. médico até 15, faltas justificadas, faltas injustificadas.
   - "Cadastro ORG": uma por pessoa do ORG (matrícula repetida = 1ª linha; matrícula `#N/A` entra
     sem matrícula, salvo se o nome já existe com matrícula válida); situacao = área do ORG.
   Gestor e grupo dessas linhas vêm do primeiro gestor/grupo visto para a matrícula na Tratamento;
   quem não aparece na Tratamento pega o gestor do ORG.
9. **Antes de ler, tira os filtros de todas as abas** (`LimparFiltros`) e acha a última linha com
   `UltimaLinhaPreenchida` (não usa `End(xlUp)`).
10. `TextoLimpo`: erro do Excel (`#N/A`) vira "", quebras de linha viram espaço.

**Saída: contrato do CSV (29 colunas, nesta ordem):**
```
data, colaborador, matricula, gestor, setor, cargo, tipo_ocorrencia, situacao, status,
duracao_minutos, destino_horas_extra, data_tratativa_pontonet, horas_excedentes,
ocorrencias_extra_falta_mes_atual, ocorrencias_extra_falta_3_meses, pausas_corretas,
pausas_menor_20min, pausas_maior_20min, trabalho_correto_140, trabalho_maior_140,
trabalho_menor_140, pausas_marcacoes_impares, grupo,
minutos_hora_extra_cartao, minutos_banco_horas_cartao, justificativa_marcacao, origem_marcacao,
dias_ausencia, emenda_folga
```
(24/25 só na "Resumo Horas Cartão"; 26/27 só na "Marcação Digitada"; 28/29 na "Absenteísmo" e na "Afastamento")
- Separador **vírgula**; campo com vírgula/aspas/quebra vai entre aspas (RFC 4180); UTF-8 com BOM
  (`ADODB.Stream`); datas `dd/mm/aaaa` (`dd/mm/aaaa hh:mm` quando tem hora); duração em minutos
  inteiros (negativo = falta); `grupo` = Interno/Externo pelo nome da aba (vazio na planilha antiga).
- `destino_horas_extra` sai sempre vazio: a Tratamento não separa banco de horas × pagamento (o
  painel trata vazio como "Pagamento"). REGRA ainda não confirmada.
- `ExportarCSVPorGestor`: um `dados_<gestor>.csv` por gestor + `dados_TODOS.csv`.

> ⚠️ O documento de contexto antigo (planejamento de 30/09, com `Base_Ocorrencias`/`Base_Excecoes`,
> separador `;`, datas ISO e flags 1/0) descrevia um **plano que não foi seguido**. O que vale é o
> contrato acima, que é o que o GerarCSV e o index usam de fato.

## 6. O painel (Dashboard/index.html)

- JS puro numa IIFE, seções numeradas 0–9 no comentário do topo. `state.records` é a única fonte
  de dados; cada mudança de filtro refaz tudo (`renderDashboard`).
- Parser próprio: aceita `,` ou `;`, BOM, datas `dd/mm/aaaa` ou ISO, cabeçalhos com sinônimos
  (`HEADER_ALIASES`).
- **Classificação de linhas:** `TIPOS_RESUMO` (auditoria, pausas, "Resumo Horas Cartão", "Cadastro
  ORG"), `TIPOS_CURTAS`, "Marcação Digitada" (`isTipoDigitada`) e "Absenteísmo"/"Afastamento"
  (`isTipoAbsenteismo`) ficam fora das vistas gerais (`isTipoNaoOcorrencia`) e têm vistas próprias; `TIPOS_HORA_EXTRA` entram só nas horas do gráfico de banco
  de horas, **não** na contagem de ocorrências. `isTipoIgnorado` descarta "Problema horário" na
  leitura (proteção para CSV antigo); `textoCampo` transforma "Erro 2042"/`#N/D` em vazio.
- `TIPOS_RESUMO` = auditoria, pausas e "Resumo Horas Cartão". O gráfico de banco de horas usa
  esse resumo (verde = extra, vermelho = banco, saldo à direita, "Mostrar todos"); sem ele (CSV
  antigo) cai para as linhas da Tratamento e avisa no subtítulo.
- Topo (faixa preta): **período** em destaque = valores do filtro de período; o período padrão
  (`periodoCompleto`) vai da 1ª data das ocorrências/curtas até a data das linhas-resumo do
  Cartão (fim do ciclo). Botão **"Entenda o painel"** liga `body.modo-explicacao`, que mostra
  as caixas `.card-fonte` (texto de "de onde vem" de cada card, escrito a partir do guia do
  Dimitri). **Só aparece com arquivo carregado** (pedido do Dimitri, 08/10); sem arquivo, a
  tela inicial (`#emptyState`) traz o passo a passo "Como usar o painel".
- Listas com "Mostrar todos" (`limitarLista`, `state.mostrarTodos`, botões `.btn-todos`): curtas
  e top de digitadas mostram 20 (`LIMITE_LISTA`).
- Filtro de status: opções só das ocorrências (as linhas-resumo saem do macro como "Pendente");
  com um só status o grupo (`#fStatusGroup`) some.
- Demora no PontoNet: `feriasPorPessoa` (linhas "Afastamento" com "férias") + `feriasDoNome`
  (gestor casado pelo nome, ou nome único que começa por ele) → aviso amarelo (`row-ferias`)
  se as férias cruzam a janela 1ª ocorrência → última tratativa. Sem aba Afastamentos no CSV,
  avisa no card. A demora em si não muda.
- Lista do ORG: seletor só de cadastro (sem matrícula/gestor/cargo/área). Não dá para listar
  "fora do Cartão": o "Resumo Horas Cartão" só traz quem teve hora > 0 (163 de 202).
- `limparControlesFiltro()` zera filtros, busca e ordenação: usada em "Limpar filtros", ao tirar
  o arquivo e ao carregar um novo.
- **Card do colaborador**: qualquer elemento com `data-colab` (nomes nos gráficos de top e de
  banco de horas, `colabLink` nas tabelas) abre `abrirCardColaborador` por delegação. As
  ocorrências do card respeitam só o período do filtro; os resumos são do mês/3 meses.
- **Lista de colaboradores (ORG)** (`renderListaORG`): linhas "Cadastro ORG" casadas por
  matrícula (`chavePessoa`) com ocorrências do filtro, digitadas e horas do Cartão. O aviso em
  vermelho de "fora do ORG" foi **retirado a pedido do Dimitri** (06/10); o card da pessoa ainda
  mostra "Não está no ORG".
- Botão **"Esconder filtros"** (`setLateralOculta`, `body.sem-lateral`, guardado em
  localStorage `painelPonto.lateralOculta`); redesenha os gráficos.
- Card do colaborador: `.modal-body.colab-modal-body` é `display:block` (rola inteiro); tabelas
  internas usam `.colab-tab` (máx. 340px só no desktop). Não usar `max-height` inline nelas: no
  celular a tabela vira cartões e o conteúdo vazava por cima.
- **Marcações digitadas** (`renderDigitadas`, filtro `getDigitadasRecords`: gestor/grupo/cargo/
  colaborador/período, sem chips): KPIs, motivo, dia da semana, data, top 10, gestor, hora, esquecimento
  repetido, justificativas (agrupadas por `chaveTexto`), lista com busca. "No dia de uma ocorrência" =
  mesma matrícula + dia de uma ocorrência da Tratamento. `isTipoDigitada` entra em
  `isTipoNaoOcorrencia` e é excluído de `getFilteredRecords`.
- **Seções**: `.pagina[data-pagina]` (geral, destaque, pausas, digitadas, afastamentos, demora, org,
  historico); menu ☰ (`#menuSecoes`, `irParaPagina`, localStorage `painelPonto.pagina`, "todas"
  mostra tudo); KPIs só em "geral". Trocar de seção redesenha (largura dos SVG). O extrator de
  conferência clica em "todas" antes de ler.
- **Pausas Térmicas**: tabela na própria seção (o modal saiu), `tabela-compacta` com cabeçalho em
  2 linhas (`c.html` em `renderTableSimple`), cabe sem rolagem lateral.
- **Absenteísmo** (`renderAbsenteismo`, base `getAtestados()` = linhas "Afastamento" cuja situação
  contém `ATESTADO_CONTEM` ou, sem elas, as "Absenteísmo" do Cartão; sem período):
  seletor de ciclo 28→27 (`cicloDe`); cada atestado expandido em dias (`diasDoAtestado`). Vistas
  escolhidas pelo Dimitri (2, 4, 5, 6, 7, 8): evolução por ciclo, por gestor/setor/cargo,
  encostados em folga + dia da semana do início, ranking 3 meses (reincidente = 2+ ciclos ou 3+
  atestados), mapa colaborador × dia, atestados no card. Sem cards de KPI (visão 1 recusada) e
  sem a separação longo × curto (visão 3 recusada). Fica fora de `periodoCompleto`. Com ciclo
  escolhido, a evolução mostra as semanas do ciclo e destaca a linha (`row-sel`). O card do
  colaborador lista todos os afastamentos (atestados em vermelho).
- Cards (KPI): Colaboradores envolvidos · Ocorrências no período · Demora média no PontoNet.
  (**Os cards de total de horas extra e de horas falta foram retirados a pedido do Dimitri.**)
- Gráficos/tabelas: por tipo (rosca), por dia da semana, por data (linha), por gestor (top 12),
  por cargo (top 12), top 10 colaboradores, extra e falta no mesmo dia (mês/3 meses, alerta 10+),
  ocorrências curtas (Faltas/Extras/Total), banco de horas e hora extra (verde/vermelho + saldo, top
  10 / "Mostrar todos"), pausas térmicas (tabela na seção), marcações digitadas, afastamentos
  (atestados), demora por gestor e por colaborador, lista de colaboradores (ORG), histórico paginado.
- Filtros: gestor, **grupo** (só aparece se o CSV tiver a coluna `grupo` preenchida), cargo,
  colaborador, tipo (chips), status (chips), período. Linhas de resumo ignoram o período.
- Paleta: laranja `#F7931E`, vermelho `#E31E24`, amarelo `#FFC72C`, preto `#1A1A1A`, cinzas.

## 7. Conferência feita (01/10/2026)

Planilha `Tratamento Ponto 28.08 até 27.09` (versão _3). Todo número do painel foi recalculado
direto das abas de origem e comparado com o que o painel mostra no Chromium: **15/15 conferem**
(141 ocorrências = linhas válidas da Tratamento; 74 colaboradores; demora média de 5,7 dias com
56 tratativas; 7 colaboradores com 10+ na auditoria; 51 nas pausas).
Bugs achados e corrigidos nessa revisão:
- Filtro na aba "Cartão atual" + `End(xlUp)` → **36 linhas perdidas** no CSV.
- Ocorrência que só tinha hora extra virava apenas "Hora Extra" e **sumia** das contagens
  (3 "Falta ponto intervalo" e 2 "Problema feriado").
- `#N/A` do gestor ia para o CSV como "Erro 2042".
- O GerarCSV antigo procurava abas "Cartão ponto…" e não achava "Cartão Agosto/atual": a
  auditoria e as curtas saíam vazias.

Conferência de 06/10 (planilha versão _2, 138 ocorrências): 15/15 de novo, com o item de banco
de horas agora comparando extra, banco e saldo de cada colaborador com o Cartão atual somado.

Achados no ORG da versão _2 (06/10): nenhum gestor em branco; 22 matrículas repetidas; 9 linhas
com matrícula `#N/A` (2 duplicatas de quem já está com matrícula; o resto é desligado ou nome
diferente da RE); 10 pessoas da RE fora do ORG. Lista nominal só na conversa com o Dimitri
(dados pessoais), nunca no repositório. Conferência: 16/16.

Marcações digitadas (06/10): aba inserida na planilha _2 com `ferramentas/adicionar_aba_xlsx.py`
(sem openpyxl; só índices do pacote + estilos no fim mudam). Conferência: 17/17.

Absenteísmo (06/10): 31 atestados, 107 dias, 23 pessoas; 19 encostam em folga; 6 reincidentes.
Conferência: 18/18.
Afastamentos (07/10, relatório HRCL006): 33 atestados, 105 dias, 21 pessoas; ciclos jul 28 / ago 30
/ set 47 dias; 17 encostam em folga; 7 reincidentes. Conferência: 18/18.

## 8. Armadilhas técnicas — não repetir

1. **VBA não diferencia maiúsculas/minúsculas nos nomes**: uma função `UltimaLinha` colide com
   uma variável `ultimaLinha` (erro de compilação). Por isso o nome é `UltimaLinhaPreenchida`.
   Sempre conferir colisão entre função/constante e variável/parâmetro.
2. **`End(xlUp)` com filtro ativo** para na última linha visível. Usar `UltimaLinhaPreenchida`
   e limpar os filtros.
3. Passar `colIdx("X")` (Variant) para um parâmetro `ByRef ... As Long` dá "ByRef argument type
   mismatch": declare como `ByVal`.
4. `CStr` de uma célula com erro gera o texto "Erro 2042": use `TextoLimpo`.
5. `Scripting.Dictionary`: ler uma chave que não existe **cria** a chave com Empty; o erro só
   aparece depois. Valide os cabeçalhos obrigatórios antes (`ColunasFaltando`).
6. Cabeçalhos "50%"/"60%" escritos sem `NumberFormat = "@"` viram 0,5/0,6.
7. O período do HRCP102 vem como data serial numa coluna estreita (`.Text` = "####"): leia `.Value`.
8. Os `.xlsx` exportados pelo Senior (HRES114, HRCP102) têm caminhos com `\` dentro do zip e
   estilos inválidos: o openpyxl não abre direto. Leia o XML da planilha (`xl/sheet1.xml`) na mão
   ou reescreva o zip trocando `\` por `/` e removendo `styles.xml`.
9. **Nunca salvar as pastas de Tratamento com openpyxl**: elas têm validações x14, Power Query
   (`customXml`, `connections.xml`), comentários encadeados e tabelas. Só leitura (`data_only=True`).
10. A tabela da aba ORG costuma estar filtrada (122 linhas ocultas) e o "Cartão atual" também já
    esteve: considere que qualquer aba pode ter filtro.
11. As planilhas têm dados pessoais reais (nome, matrícula, salário na RE): não publique nem
    envie esses dados para fora.

## 9. Pendências e próximos passos

- **Rodar o GerarCSV novo no Excel de verdade.** As macros foram validadas por simulação em
  Python e por checagem estrutural, mas **nunca foram executadas no Excel** nesta sessão.
- Decidir com o Dimitri: um dia com Check "S" deve sair das "ocorrências curtas" e da auditoria
  (que vêm do Cartão)?
- A planilha de 28/08 a 27/09 ainda está no formato antigo (aba única "Tratamento"). Separar em
  Internos/Externos depende de um **critério** que o Dimitri ainda não deu.
- Regra de `destino_horas_extra` (banco × pagamento) não confirmada.
- Absenteísmo: hoje só atestados. Quando o Dimitri liberar, incluir os outros eventos da lista
  (suspensão, acidente, faltas justificadas/injustificadas) em `ABSENTEISMO_CONTEM`. Atestado
  de poucas horas (dia "Trabalhando" com horas em "Just.") não entra: o Cartão não diz o motivo.
- Sugestões para o painel ainda não aprovadas: comparativo com o mês anterior; pendências mais
  antigas ("há X dias"); % tratado por gestor; reincidentes (3+ do mesmo tipo); ocorrências por
  colaborador da equipe; mapa de calor colaborador × dia (a sexta concentrou 51 de 141); exportar
  o histórico filtrado; nomes inteiros nos gráficos (hoje cortados em 15 caracteres).

## 10. Como conferir sem Excel (ferramentas/)

```bash
pip install openpyxl                      # Playwright já vem no ambiente web do Claude Code
# (opcional) formatar relatórios do sistema e inserir as abas na planilha, sem openpyxl:
PYTHONPATH=ferramentas python3 ferramentas/formatar_marcacoes.py "Marcações digitadas.xlsx" marcacoes.pkl
PYTHONPATH=ferramentas python3 ferramentas/formatar_afastamentos.py "Afastamentos.xlsx" afastamentos.pkl
python3 ferramentas/adicionar_aba_xlsx.py entrada.xlsx saida.xlsx marcacoes.pkl marcacoes
python3 ferramentas/adicionar_aba_xlsx.py saida.xlsx saida2.xlsx afastamentos.pkl afastamentos
python3 ferramentas/carregar_planilha.py "Tratamento Ponto ....xlsx"     # -> planilha.pkl
python3 ferramentas/simular_gerarcsv.py                                  # -> dados_TODOS.csv (regras da seção 5.4)
NODE_PATH=$(npm root -g) node ferramentas/extrair_painel.js "$PWD/Dashboard/index.html" "$PWD/dados_TODOS.csv" painel.json
PYTHONPATH=ferramentas python3 ferramentas/conferir_painel.py painel.json   # OK/FALHA item a item (hoje 18 itens)
python3 ferramentas/checar_colisoes_vba.py */*.bas                         # antes de entregar VBA
```
Os scripts foram escritos para a planilha de 28/08 a 27/09 (nomes de aba fixos: "Tratamento",
"RE 08.09", "Pontonet", "Pausas Térmicas", "Cartão atual/Agosto/Julho"). Ajuste os nomes para
outras planilhas. Para comparar com uma aba CSV gerada pelo Excel, leia a aba "CSV" do
`planilha.pkl` e compare linha a linha (data, matrícula, tipo, duração).

## 11. Histórico do repositório

- Antes: macros em uso, sem versionamento (commit "Adiciona macros de ponto (versões atuais…)").
- 30/09: horário do dia no Cartão; pausas só com as tabelas; PURO sem espaços; GerarCSV com
  Internos/Externos, RE e Cartão achados por prefixo.
- 01/10: index versionado; filtros, Check S, Problema horário, ocorrência só com extra, Erro 2042;
  cards de horas retirados; filtro de grupo; conferência 15/15.
- 06/10: resumo de horas do Cartão (CSV com 25 colunas) e gráfico verde/vermelho com saldo;
  período no topo; filtros limpos ao tirar/carregar arquivo; card do colaborador; modo
  "Entenda o painel". Depois: lista do ORG no CSV ("Cadastro ORG") e seção "Lista de
  colaboradores (ORG)" com aviso de quem está fora do ORG; gestor das linhas-resumo pelo ORG.
  Depois: macro FormatarMarcacoesDigitadas, linhas "Marcação Digitada" no CSV (27 colunas) e seção
  "Marcações digitadas" no painel. Depois: Origem "E" ignorada; absenteísmo (atestados) no CSV
  (29 colunas) e seção "Absenteísmo (atestados)" no painel. Depois: botão "Esconder filtros",
  card do colaborador sem tabelas espremidas, aviso "fora do ORG" retirado da lista.
- 07/10: macro FormatarAfastamentos e linhas "Afastamento" no CSV (relatório HRCL006 substitui o
  Cartão como fonte dos atestados); menu de seções; Pausas na própria seção com colunas estreitas;
  evolução de atestados reage ao ciclo.
- 08/10: revisão do painel com o Dimitri (README, "Revisão do painel com o Dimitri"): guia na tela
  inicial, "Entenda o painel" só com arquivo, curtas/top digitadas com 20 + "Mostrar todos", 2
  cards de digitadas retirados, folga encostada com data, mapa com todos os atestados de 3 meses,
  aviso de férias na demora, seletor do ORG de cadastro, filtro de status sem o falso "Pendente".
  Planilha nova dele (ORG corrigido, 210 pessoas) **sem a aba Afastamentos**: conferência 18/18
  com o Absenteísmo do Cartão; o aviso de férias foi testado com linhas de férias simuladas.
  Depois ele mandou a versão "com Marcações e Afastamentos" (a de 07/10, ORG antigo, 215): a aba
  Afastamentos dela foi inserida na planilha de ORG corrigido (`adicionar_aba_xlsx.py`).
  Resultado: 210 no ORG, 283 afastamentos, 33 atestados; conferência 18/18. Aviso de férias: só
  1 gestor. **O Tyrone não está no HRCL006 nem no Cartão** (só na RE): o relatório de
  afastamentos precisa ser exportado incluindo os gestores para o aviso dele aparecer.

## 12. Estado atual e como continuar (08/10/2026)

**Onde está o código.** Repositório `dimitrimateus/dashboard_ponto`, branch `claude/new-session-69o070`.
Os commits de 06/10 a 08/10 foram feitos, mas o **push para o GitHub foi bloqueado** pelas permissões
do ambiente anterior. Se o seu repositório não tiver os arquivos `Afastamentos/`,
`MarcacoesDigitadas/` e `ferramentas/formatar_afastamentos.py`, o código mais novo veio no zip
`Dashboard_Ponto (codigo).zip` junto com este arquivo: use-o como base. Não tente "contornar" um
bloqueio de push; peça para o Dimitri liberar ou empurrar ele mesmo.

**O que o Dimitri tem para testar** (pacote `Teste Ponto 07-10.zip`, entregue a ele): as 6 macros em
`.txt`, o `index.html`, a planilha `Tratamento Ponto 28.08 até 27.09_2 - com Marcações e
Afastamentos.xlsx` (abas inseridas com `adicionar_aba_xlsx.py`) e um CSV de referência simulado.
Números esperados quando ele rodar o GerarCSV nessa planilha: 138 ocorrências, 72 colaboradores,
demora média 5,7 dias, 422 batidas digitadas (12 de Origem "E" ignoradas), 283 afastamentos (33
atestados, 105 dias), 215 pessoas no ORG, 163 com horas no Cartão. **As macros nunca rodaram no
Excel**: o próximo retorno dele provavelmente traz erros de compilação/execução (peça print da
mensagem e da linha amarela).

**Decisões já tomadas (não reabrir sem ele pedir):**
- Check "S" ignora a linha inteira; "Problema horário" não entra; Origem "E" das digitadas não entra.
- Cards de horas extra/falta foram retirados. Aviso vermelho "fora do ORG" na lista foi retirado.
- Absenteísmo = **só atestados** por enquanto; fonte oficial = relatório HRCL006 (não o Cartão).
- Visões de absenteísmo escolhidas: 2, 4, 5, 6, 7, 8 (sem cards de KPI e sem longo × curto).
- O relatório de digitadas não diz quem digitou (colaborador × RH): não há como separar.
- Painel em seções com menu ☰; Pausas Térmicas na própria seção, sem rolagem lateral.
- (08/10) Sem os cards "No dia de uma ocorrência" e "Horário redondo" nas digitadas; sem
  "com/sem ocorrência no filtro" na lista do ORG; "Entenda o painel" só com arquivo carregado.

**Aguardando o Dimitri escolher** (sugestões feitas em 07/10, com base no HRCL006):
1. Absenteísmo completo: somar acidente de trabalho e faltas injustificadas (e justificadas e
   suspensão quando aparecerem) + **taxa de absenteísmo** (dias ausentes ÷ dias previstos no Cartão).
2. Ausências por horas: "Saída Médico/Empresa" e "Horas Faltas" em horas por pessoa (uma pessoa
   concentra 14 das 26 saídas médicas; ~65h no total).
3. "Quem está fora agora": afastados no fim do ciclo com previsão de retorno.
4. Mapa da equipe com todos os tipos de afastamento (uma cor por tipo).
5. Cruzamento digitadas × afastamentos: das 167 digitadas por "atividade externa", só 2 caem em dia
   com "Viagem a Serviço" registrada (auditoria).
6. Férias por gestor/semana (55 períodos, 1.457 dias).
7. Tempo de casa (admissão) × atestados.
Também em aberto: Check "S" nas curtas/auditoria; critério Internos × Externos;
`destino_horas_extra`; sugestões antigas da seção 9.

**Correções de cadastro que o Dimitri vai fazer no ORG** (lista nominal só com ele): nome com letra
trocada que quebra o PROCX, pessoas da RE fora do ORG, desligados ainda no ORG. Ele perguntou se
corrigir direto no Excel resolve: sim, o GerarCSV lê o ORG a cada execução.

**Como trabalhar com ele (o que funcionou):**
- Toda resposta em pt-BR, explicando o porquê; dizer direto quando algo entregue estava errado.
- Entregar VBA em `.txt` (CRLF): `sed 's/$/\r/' X/X.bas > txt/X.txt`. Manter `.bas` e `.txt` iguais.
- Antes de entregar: `checar_colisoes_vba.py`, simular com `simular_gerarcsv.py`, abrir o painel no
  Chromium (Playwright) e rodar `conferir_painel.py` (todos OK). Olhar screenshots desktop e 390px.
- Relatórios novos do sistema (Senior) chegam como "dump" de impressão: cabeçalho repetido, valores
  deslocados em relação aos títulos, textos quebrados na linha de baixo. Padrão adotado: uma macro
  `FormatarX` que reescreve a aba ativa como tabela (cabeçalho na linha 1), achando colunas pelo
  título; o GerarCSV acha a aba pelo cabeçalho; o painel ganha um `tipo_ocorrencia` novo, excluído
  das vistas gerais e do `periodoCompleto`, com seção própria e texto de "De onde vem".
- Quando ele manda uma planilha nova, quer as **abas inseridas na planilha de Tratamento** dele: use
  `adicionar_aba_xlsx.py` (nunca salve com openpyxl). O LibreOffice deste tipo de ambiente não
  abre nem planilhas simples: valide pela estrutura (zip/XML/contagens) e peça para ele abrir no Excel.
- Ele manda pacotes de teste em zip quando pede "todos os arquivos": macros, painel, planilha e CSV de
  referência, em pastas.
- Dados pessoais (nomes, matrículas, salários): nunca no repositório nem em commits; só na conversa.

