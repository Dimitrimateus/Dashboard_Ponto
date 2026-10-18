# CONTEXTO DO PROJETO — Tratamento de Ponto e Painel de Ocorrências (AuroraCoop)

> Arquivo de contexto para o Claude Code. Estado em **13/10/2026**. Leia inteiro antes de propor
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
Intrajornada (HRES108), Interjornada,
Interjornada semanal, Horas excedentes ──► FormatarJornada (1 macro, acha o tipo) ──► abas "Intrajornada",
                                           "Interjornada", "Interjornada Semanal", "Horas Excedentes"
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
| `Jornada/FormatarJornada.bas` | os 4 relatórios de limite de jornada (intra, inter, inter semanal, horas excedentes) → a mesma aba como tabela, renomeada pelo tipo |
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
- (18/10) Lê o fim do período (linha "Período: ... até ...") antes de apagar e **renomeia a aba**
  para "Pausas Térmicas <Mês>" (`NomeDoMes`, `AbaExiste`; se o nome já existe, avisa e mantém).
  Versão Python: `ferramentas/formatar_pausas.py` (lê o numFmt do styles.xml para separar linha
  de pessoa da linha de dia, como a macro faz com `NumberFormat`).
- (18/10, depois) **Terceira tabela: as batidas** (`GuardarBatidasDoDia`, `EscreverTabelaBatidas`):
  duas linhas em branco depois do Total Geral, título "Marcações do período (uma linha por
  batida)" e uma Tabela do Excel (ListObject, filtro próprio): Matrícula, Colaborador, Data, Dia da
  semana, Hora, Nº da batida no dia, Tipo da batida, Letra no relatório, Exceções do dia. Linha de dia
  do relatório = data em A + pares horário/letra (B/C, D/E...) + exceções "(1) (3)" no fim (col. AH).
  **Letras: P = entrada/saída da jornada, G = refeição, F = pausa térmica** (dedução conferida:
  pausas F–F de 20 min e blocos > 1h40 recontados batem com o "Total do Colaborador"; "trabalho
  menor" difere em 1–2 em quem tem dia ímpar). Batida depois da meia-noite fica na data do turno.
  A GerarCSV continua lendo só a tabela 1 (para na linha em branco). `formatar_pausas.batidas()`
  é a versão Python. Jul/ago/set: 9.545 / 9.992 / 9.009 batidas.

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

### 5.3d FormatarJornada (aba ativa = um dos 4 relatórios de jornada)
- Uma macro só: `TipoDoRelatorio` olha os títulos ("Registros de Ponto"+"Intervalo" → intra;
  "Horas Descansadas" → inter; "Horas DSR"/"Marcação Antes DSR" → semanal; "Carga Hor."+"Extras" →
  excedentes) e chama a rotina `Montar*` do tipo. Versão Python: `ferramentas/formatar_jornada.py`.
- **Intrajornada** (HRES108, intervalo menor que o mínimo): batidas começam NA coluna de "Registros
  de Ponto" (8 colunas); depois carga e até 3 intervalos. "Dia" = "Sex (985)" (dia + cód. horário).
  Saída: `Matrícula, Colaborador, Local, Admissão, Cargo, Data, Dia, Cód. Horário, Marcações, Carga
  horária, Intervalo 1, Intervalo 2, Intervalo 3, Intervalo total`.
- **Interjornada** (< 11h entre dias): valores = título+1 (marcações: data em +1, hora em +2).
  Saída: `Matrícula, Colaborador, Cód. Local, Local, Cargo, C.C., Data apuração, Marcação anterior,
  Marcação atual, Horas descansadas, Ocorrência`.
- **Interjornada semanal** (trabalhou no DSR/feriado): texto "Trabalhou no Dia do DSR/Feriado
  (dd/mm/aaaa)" numa célula mesclada (H ou I); linha sem matrícula = mesma pessoa. Saída:
  `Matrícula, Colaborador, Cód. Local, Local, Data DSR/Feriado, Ocorrência`.
- **Horas excedentes** (> 10h no dia): matrícula na col. de "Colaborador/Nome" (nome +1); carga =
  título+1; batidas = 8 colunas a partir de "Marcações"+1, depois extras, horas, conv. noturna, total.
  Saída: `Matrícula, Colaborador, Cargo, C.C., Filial, Cód. Local, Local, Data, Carga horária,
  Marcações, Extras, Horas trabalhadas, Conv. noturna, Total`.
- Arquivos de 09/10 (28/06 a 27/09): intra 21 ocorrências/18 pessoas; inter 1; semanal 11/8;
  excedentes 72/34. Horas arredondadas ao minuto (o sistema grava 10:05:00,96).

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
   - "Resumo Pausas Térmicas": uma por colaborador da tabela de pausas, **de cada aba de pausas**
     (13/10, T9): a aba sem mês no nome é o mês atual (data = fim do Cartão); as de meses
     anteriores têm o mês por extenso no nome ("Pausas Térmicas Agosto") e saem com a data 27
     daquele mês (`CabecalhoPausasNaAba`, `MesNoNomeDaAba`, `DataDoMesDaAba`); mês repetido é
     ignorado com aviso. O painel usa só o mês mais recente, menos na evolução e no card.
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
   - "Intrajornada" / "Interjornada" / "Interjornada Semanal" / "Horas Excedentes" (10/10):
     `EncontrarAbaJornada` (pelos títulos) + `GerarLinhasJornada`, uma por ocorrência. situacao =
     "Sem intervalo"/"Intervalo de 0:55" (intra), a Ocorrência (inter e semanal), "Total de 12:40
     no dia" (excedentes); duracao_minutos = extras (só excedentes); colunas 30/31/32 =
     minutos_trabalhados / minutos_descanso / marcacoes_dia. Sem Tratamento nem Check.
   - "Cadastro ORG": uma por pessoa do ORG (matrícula repetida = 1ª linha; matrícula `#N/A` entra
     sem matrícula, salvo se o nome já existe com matrícula válida); situacao = área do ORG.
   Gestor e grupo dessas linhas vêm do primeiro gestor/grupo visto para a matrícula na Tratamento;
   quem não aparece na Tratamento pega o gestor do ORG.
9. **Antes de ler, tira os filtros de todas as abas** (`LimparFiltros`) e acha a última linha com
   `UltimaLinhaPreenchida` (não usa `End(xlUp)`).
10. `TextoLimpo`: erro do Excel (`#N/A`) vira "", quebras de linha viram espaço.

**Saída: contrato do CSV (32 colunas, nesta ordem):**
```
data, colaborador, matricula, gestor, setor, cargo, tipo_ocorrencia, situacao, status,
duracao_minutos, destino_horas_extra, data_tratativa_pontonet, horas_excedentes,
ocorrencias_extra_falta_mes_atual, ocorrencias_extra_falta_3_meses, pausas_corretas,
pausas_menor_20min, pausas_maior_20min, trabalho_correto_140, trabalho_maior_140,
trabalho_menor_140, pausas_marcacoes_impares, grupo,
minutos_hora_extra_cartao, minutos_banco_horas_cartao, justificativa_marcacao, origem_marcacao,
dias_ausencia, emenda_folga, minutos_trabalhados, minutos_descanso, marcacoes_dia
```
(24/25 só na "Resumo Horas Cartão"; 26/27 só na "Marcação Digitada"; 28/29 na "Absenteísmo" e na
"Afastamento"; 30/31/32 nas 4 de jornada)
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
  (`isTipoAbsenteismo`) e os 4 de jornada (`isTipoJornada`, `TIPOS_JORNADA`; seção "Jornada"
  desde 13/10) ficam fora das vistas gerais (`isTipoNaoOcorrencia`) e têm vistas próprias; `TIPOS_HORA_EXTRA` entram só nas horas do gráfico de banco
  de horas, **não** na contagem de ocorrências. `isTipoIgnorado` descarta "Problema horário" na
  leitura (proteção para CSV antigo); `textoCampo` transforma "Erro 2042"/`#N/D` em vazio.
- `TIPOS_RESUMO` = auditoria, pausas e "Resumo Horas Cartão". O gráfico de banco de horas usa
  esse resumo (verde = extra, vermelho = banco, saldo à direita, lista inteira em tela cheia); sem ele (CSV
  antigo) cai para as linhas da Tratamento e avisa no subtítulo.
- Topo (faixa preta): **período** em destaque = valores do filtro de período; o período padrão
  (`periodoCompleto`) vai da 1ª data das ocorrências/curtas até a data das linhas-resumo do
  Cartão (fim do ciclo). Botão **"Entenda o painel"** liga `body.modo-explicacao`, que mostra
  as caixas `.card-fonte` (texto de "de onde vem" de cada card, escrito a partir do guia do
  Dimitri). **Só aparece com arquivo carregado** (pedido do Dimitri, 08/10); sem arquivo, a
  tela inicial (`#emptyState`) traz o passo a passo "Como usar o painel".
- **Filtros de marcar vários** (10/10, pedido do Dimitri): gestor, grupo, cargo, colaborador, tipo e
  status são o componente `.ms` (`msCriar`, `msValores`, `msSetOf`, `msLimpar`): botão que abre a
  lista com caixinhas (busca em gestor/cargo/colaborador); nada marcado = todos. Os filtros são lidos
  por `filtroPessoas`/`passaPessoa`/`periodoFiltro` (getFilteredRecords, getResumoRecords,
  getDigitadasRecords, getCurtasRecords). O extrator lê as opções em `#fTipo .ms-op`.
- **Barras empilhadas por tipo** (`empilhar` + `legendaMini`, `it.segs` em renderBarH/renderBarV;
  pedaços com classe `seg`, ignorados pelo extrator): Visão geral (dia da semana, gestor, cargo,
  top colaboradores) por tipo de ocorrência (`corSerie("tipo")`, mesmas cores da rosca);
  digitadas (dia, top, gestor, hora) por motivo (`corSerie("motivo")`); afastamentos por grupo.
  Cada gráfico tem filtro próprio na legenda, tela cheia (⛶), selo 📅 de período e clique
  "de onde vem" (ver histórico 13/10 na seção 11).
- **Afastamentos com todos os tipos** (10/10): seletor `#absMostrar` (todos × só atestados);
  `getAfastamentosBase`, `GRUPOS_AFAST` (7 grupos com cor fixa, atestado vermelho), `grupoAfast`,
  `corAfast`; `isPorHoras` (duração > 0 e não atestado: saída médica, horas faltas, reunião curta)
  fica fora da conta de dias e aparece no resumo em horas. Ciclos cortados no fim do arquivo
  (`periodoCompleto`), para férias já programadas não criarem um mês incompleto. Encostados em
  folga e ranking continuam só de atestados. Extrator escolhe "Só atestados" antes de ler.
- Listas longas (`limitarLista`, botões `.btn-todos` e `#btnBancoTodos`): curtas e top de
  digitadas mostram 20 (`LIMITE_LISTA`), banco de horas 10; a lista inteira só em tela cheia.
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
  mostra "Não está na lista de equipes".
- Botão **"Esconder filtros"** (`setLateralOculta`, `body.sem-lateral`); redesenha os gráficos.
  Não é guardado: ao abrir, os filtros sempre aparecem (13/10).
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
- **Jornada** (13/10, `renderJornada`, `getJornadaRecords`: filtros de pessoa, sem período, seletor
  `#jorCiclo`): J1 4 números + barras por ciclo por tipo; J2 mais de 10h por pessoa (barra = minutos
  acima de 10h, texto "N dias · Xh", clique = dias com batidas); J7 por gestor (gráfico + tabela com
  % da equipe via `tamanhoEquipes`) + lista. Nomes leigos em `JORNADA_INFO` ("Mais de 10h no dia",
  "Intervalo de almoço curto", "Trabalho na folga ou feriado", "Menos de 11h de descanso").
  **Horas do dia = coluna "Total"** do relatório (relógio + conversão da hora noturna), a que o
  sistema usa para o limite; "Horas trabalhadas" (só relógio) daria 28h51 em vez de 66h acima.
- **Sinais de atestado** (`renderSinaisAtestado`, 3 meses): P1 2+ atestados começando no mesmo dia da
  semana; P2 atestado de 1 dia na seg/sex (resumo, por pessoa, por equipe); P3 atestado a até 3 dias
  das férias (`DIST_FERIAS_DIAS`, antes/depois/dentro). Sem selo no título (ver decisão de 18/10).
- **Esquecimento crônico** (P10, `renderEsquecimentoCronico`): por pessoa e ciclo, sem período,
  tendência penúltimo → último ciclo, vermelho com `LIMITE_ESQUECIMENTO` (5)+ no último.
- **Pausas** (`renderPausas`): T1 4 números; T3 acima de 1h40 por pessoa (vermelho > 30%,
  `LIMITE_ACIMA_PCT`) e por gestor; T4 certas/curtas/longas por gestor ou cargo (`#pauAgrupar`);
  T9 evolução por mês (`getPausasTodosMeses`); T8 no card (`pausasDoCard`, pessoa × média da equipe).
  `getPausasTermicasRecords` = só o mês mais recente (`mesPausasAtual`).
- **M3** `renderPorGestor` (seletor `#gestorMedida`: total × por pessoa da equipe); **M4** botões
  "Filtrar o painel: ..." na janela de origem (`opts.filtro`/`filtroSeg` → `marcarOrigem` →
  `aplicarFiltroDoGrafico`; tipos "Outros" não filtram; ponto da linha filtra o dia); **M7** % do
  total (`opts.total`, `pctTxt`, `.legend-pct` na rosca, fora do que o extrator lê).
  `it.texto` (texto à direita da barra) e `it.dica` (balão) sobrescrevem o padrão em renderBarH.
- **Tendências** (18/10, pedido do Dimitri): `INDICADORES` (dig, esq, atest, he, intra, folga,
  inter, jor [só nas seções], pauCertas e pauAcima [%, `num/den`], pauImpar), cada um com `itens()`
  → {r, c: ciclo, num, den} e `pior` ("sobe"/"desce"). `ciclosTendencia()` = os 3 ciclos que
  terminam no ciclo do fim de `periodoCompleto`. `agregarTendencia`, `valorCiclo`, `tendenciaDe`
  (penúltimo → último; vermelho = piorou), `miniBarras` (SVG 3 barras), `renderTabelaTendencia`,
  `linhasTendencia`, `variacoes` (% de pessoa só com base ≥ 10 nos dois ciclos). Seção
  `data-pagina="tendencias"` (`renderTendencias`, seletor `#tendIndicador`, nome do assunto
  clicável `data-tend-ind`); cards por seção via `renderTendenciaSecao`; card do colaborador
  `tendenciaDoCard`; digitadas ciclo a ciclo (`chartDigCiclo`); coluna de tendência na "Extra e
  falta" (mês atual × média dos 2 anteriores = (3 meses − mês) ÷ 2). Filtros de pessoa sim,
  período não. **`periodoCompleto` não conta as digitadas** (com 3 meses de digitadas o painel
  continuaria abrindo no mês atual).
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
PYTHONPATH=ferramentas python3 ferramentas/formatar_jornada.py "intrajornada.xlsx" intra.pkl   # idem p/ os outros 3
python3 ferramentas/adicionar_aba_xlsx.py saida2.xlsx saida3.xlsx intra.pkl intrajornada        # interjornada|intersemanal|excedentes
python3 ferramentas/carregar_planilha.py "Tratamento Ponto ....xlsx"     # -> planilha.pkl
python3 ferramentas/simular_gerarcsv.py                                  # -> dados_TODOS.csv (regras da seção 5.4)
NODE_PATH=$(npm root -g) node ferramentas/extrair_painel.js "$PWD/Dashboard/index.html" "$PWD/dados_TODOS.csv" painel.json
python3 ferramentas/formatar_pausas.py "pausa_termica.xlsx" pausas.pkl       # -> (fim do período, linhas)
python3 ferramentas/adicionar_aba_xlsx.py in.xlsx out.xlsx dig3.pkl marcacoes "Marcações Digitadas" --substituir  # troca o conteúdo
PYTHONPATH=ferramentas python3 ferramentas/conferir_painel.py painel.json   # OK/FALHA item a item (hoje 27 itens)
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

- 10/10: macro FormatarJornada (intrajornada, interjornada, interjornada semanal, horas
  excedentes), abas inseridas na planilha de ORG corrigido, GerarCSV com 32 colunas e 4 tipos
  novos (105 linhas). Painel: só exclui os tipos novos das vistas gerais (conferência 18/18);
  seção própria aguardando o Dimitri escolher as vistas.

- 10/10 (depois): filtros de marcar vários, barras empilhadas por tipo com legenda, afastamentos com
  todos os tipos. Ideias para o Dimitri escolher numa página com banco de dados (artifact
  https://claude.ai/artifact/SZEwQzRntr66SCDMms4YBG, doc `escolhas/atual`: marcadas[], observacao);
  códigos J1–J8 (jornada), A1–A7 (afastamentos), P1–P12 (pegar no pulo), M1–M12 (melhorias).
  Gestores também veem as análises de "pegar no pulo" (decisão do Dimitri).
  **Escolhas salvas do Dimitri (10/10)**, a implementar: J1 resumo de jornada por ciclo, J2 horas
  acima de 10h por pessoa, J7 risco trabalhista por gestor, P1 atestado no mesmo dia da semana, P2
  atestado de 1 dia na segunda/sexta, P3 atestado colado em férias, P10 esquecimento crônico, M3
  ocorrências por pessoa da equipe, M4 clicar no gráfico filtra o painel, M7 percentuais nos
  gráficos, M10 legenda clicável. Depois ele pediu ideias para Pausas Térmicas: tema T (T1–T12)
  na mesma página; marcou (11/10) **T1** 4 números no topo, **T3** trabalho acima de 1h40 por
  pessoa/gestor, **T4** pausas por gestor e cargo, **T8** pausas no card do colaborador, **T9**
  evolução mês a mês (exige manter a aba de pausas de cada mês).
- 11/10: subtítulo "de X ..." do card "Colaboradores envolvidos" e o % de pessoas das digitadas
  usam `equipeDoFiltro` (tamanho da equipe do recorte gestor/grupo/cargo/colaborador, pelas linhas
  "Cadastro ORG" + quem tem ocorrência e não está no ORG; sem período/tipo/status). Ex.: "de 22 na
  equipe de <gestor>", "de 34 nas equipes dos 2 gestores", "de 211 pessoas no ORG" (210 no ORG +
  1 com ocorrência fora dele). Antes era sempre o total do arquivo.
- 13/10: pedidos do Dimitri no painel. (1) Botão ⛶ em todo card com gráfico/tabela: tela cheia
  (`abrirTelaCheia`, `.card-cheia`; `containerWidth` desenha a 2/3 para tudo ficar 1,5x maior; Esc
  fecha). (2) **Filtro dentro do gráfico**: legenda clicável (`legendaSerie`, `filtrarSerie`,
  `state.serieSel` por dimensão: `tipo` na Visão geral/destaque, `motivo` nas digitadas, `afast` nos
  afastamentos, incluindo o mapa); clicar = só aquele, outros somam, "Mostrar todos" volta; não
  muda KPIs, resumos, tabelas nem filtros da lateral. (3) Selo 📅 em cada card com o período que
  ele considera (`atualizarPeriodosCards`, `PERIODO_DOS_CARDS`: filtro, mês do Cartão, pausas,
  ciclo/3 meses dos afastamentos via `state.absPeriodo`, cadastro). (4) Mapa colaborador × dia:
  balão do painel (`data-tooltip`) com pessoa, dia e tipo. (5) Filtros da lateral sempre visíveis ao
  abrir (não guarda mais `painelPonto.lateralOculta`). (6) Arquivo carregado abre na Visão geral.
  (7) Cores: `CORES_SERIE` (5 cores validadas para daltonismo, sem vermelho, que é alerta) +
  cinza "Outros" a partir do 6º tipo; cor fixa por tipo, decidida ao carregar
  (`prepararCoresSeries`). `GRUPOS_AFAST` recolorido (atestado continua vermelho; antes havia
  laranja×marrom e vermelho×vinho parecidos). Conferência 18/18.
- 13/10 (depois): o filtro da legenda passou a ser **por gráfico** (`state.serieSel[idDoGráfico]`;
  cada gráfico chama `legendaSerie(id, …)`/`filtrarSerie(id, …)`). Clique na legenda em tela cheia
  não fecha mais a tela cheia (o clique-fora ignora alvo desconectado, `e.target.isConnected`).
  Listas longas (curtas, top de digitadas, banco de horas) só aparecem inteiras em **tela cheia**:
  o botão virou "Ver todos (N) em tela cheia" (`limitarLista`/`emTelaCheia`); ao fechar volta aos
  primeiros (`state.mostrarTodos`/`bancoMostrarTodos` não são mais usados). Mapa colaborador × dia:
  legenda em cima, botão ⛶ (cards com `.heat-wrap` também ganham), e com filtro na legenda quem
  fica sem nenhum dia no ciclo some. Conferência 18/18.
- 13/10 (depois 2): **"De onde vem"**: clicar em barra, pedaço de barra, fatia, ponto da linha ou
  quadrado do mapa abre a janela `#origemOverlay` (`abrirOrigem`) com os registros por trás do
  número (ocorrências, batidas digitadas ou afastamentos, + resumo por pessoa; nomes abrem o card).
  Os registros vão em `it.recs`/`seg.recs` (`empilhar`, `anexarOrigem`, `legendaSerie`,
  `ocorrenciasPorData`), o elemento leva `data-drill="barra[:pedaço]"` e o contêiner `marcarOrigem`.
  Nos gráficos de colaborador, o nome abre o card e a barra abre a origem. Esc fecha só a camada de
  cima (`e._tratado`). Textos do "Entenda o painel" (34 `.card-fonte`) **reescritos para leigos**
  (pedido do Dimitri): sem nomes de aba, coluna, macro, código de relatório, "Check", "ORG",
  "Tratamento", "DSR"; explicam o que é cada coisa (espelho de ponto, ciclo 28→27, batida
  digitada, pausa térmica, afastamento). Card da pessoa: "Não está na lista de equipes".
  Ao escrever texto novo para a tela, manter esse tom.
- 13/10 (depois 3): na janela "de onde vem", os nomes de cima viraram **filtros** da tabela
  (`desenharOrigem`, `state.origem.sel`; o card da pessoa abre pelo nome na tabela) e a janela rola
  inteira (`.origem-body` com `min-height:0; overflow-y:auto`). Varredura completa no Chromium
  (todas as seções, cliques em todas as legendas, tela cheia de todo card, 3 cliques de origem por
  gráfico, filtros, período vazio, ciclos/seletores dos afastamentos, lista de equipes, recarregar
  arquivo, CSV antigo de 23 colunas, celular 390 e 360 px): sem erro de JavaScript nem texto
  "NaN/undefined". Corrigidos: topo passava da tela no celular (rolagem lateral de 6 px em 390 e
  4 px em 360); "Esconder filtros" ficava laranja no modo explicação; termos técnicos nos
  subtítulos e selos ("Cartão Ponto", "DSR", "ORG", "#N/A") trocados por linguagem simples.
- 13/10 (depois 4): a janela "de onde vem" **não rolava** (relato do Dimitri): a regra
  `.origem-body{overflow-y:auto}` perdia para `.modal-body{overflow:hidden}`, de mesma
  especificidade e escrita depois no arquivo. Agora é `.modal-body.origem-body`. Testado no
  Chromium com a roda do mouse: desktop, 390 px e com o gráfico em tela cheia. Lição: regra
  específica de um modal precisa vencer a de `.modal-body` pela especificidade, não pela ordem.
- 13/10 (depois 5): na janela "de onde vem", o cabeçalho da tabela grudava 14 px abaixo do topo
  (padding da janela) e as linhas apareciam nessa faixa. Agora a janela não tem padding em cima
  (fica nos nomes, `.origem-pessoas{padding-top}`) e o `th` gruda em `top:0` com linha embaixo.
  No celular a tabela vira cartões, sem cabeçalho (cada cartão tem os rótulos).

- 13/10 (depois 6): **construídas as 15 escolhas pendentes** da página de ideias (J1, J2, J7, P1,
  P2, P3, P10, T1, T3, T4, T8, T9, M3, M4, M7; ver seção 6). GerarCSV lê todas as abas de pausas
  (T9); simulador igual. M4 virou botão dentro do "de onde vem" (o clique na barra continua abrindo
  a origem, decisão de 13/10). Conferência 26/26 (8 itens novos); varredura no Chromium (todas as
  seções, 1400/390/360 px, cliques de origem, legendas, tela cheia, modo explicação; CSV atual,
  antigo de 23 colunas e um de teste com 2 meses de pausas): sem erro de JS, NaN ou rolagem lateral.
  Também sugeridas (não aprovadas) ideias para reunião R1–R10: página da equipe por gestor, pontos
  de atenção em frases, pendências com idade, quadro comparando gestores, seta vs ciclo anterior,
  modo reunião (com nomes escondidos), PDF por equipe, combinados (aba nova), custo da hora extra
  (salário da RE, só por equipe), atividade externa sem viagem. Perguntas em aberto: reunião é por
  gestor ou com todos? Pode usar salário (só total da equipe)?

- 18/10: texto "pegar no pulo" retirado da tela (decisão). Dimitri mandou digitadas de 3 meses
  (28/06 a 27/09: 1.664 batidas, 1.642 "D"; o ciclo de setembro tem 428 contra 434 da exportação
  anterior: o sistema mudou 8/2 batidas) e as pausas de jul/ago/set (52/53/51 pessoas; setembro
  difere da aba anterior em 4 pessoas). Planilha montada: aba "Marcações Digitadas" e "Pausas
  Térmicas" com o conteúdo trocado (`--substituir`) + "Pausas Térmicas Julho" e "Agosto"; CSV
  simulado com 3.082 linhas. Seção "Tendências" + tendências em cada seção e no card (ver seção 6).
  FormatarPausasTermicas renomeia a aba pelo mês. Conferência 27/27 nas duas planilhas.

- 18/10 (depois): FormatarPausasTermicas ganhou a tabela de batidas (ver 5.2); prévia do resultado
  dos 3 meses enviada ao Dimitri (arquivo à parte, feito com openpyxl, não é a planilha dele).
  Análise das batidas (3 meses, 3.002 dias-pessoa) para propor vistas B1–B14 (ver conversa):
  13h tem 22% dos períodos acima de 1h40; turno do dia 7,7% × noite 4,2%; 229 dias com mais de
  1h40 sem nenhuma pausa; 98% das pausas com exatos 20 min; batida faltando sempre de pausa (F);
  nenhuma batida de pausa digitada (659 entradas/saídas digitadas). **Nada disso está no CSV nem
  no painel ainda**: depende de ele escolher as vistas (exigirá a GerarCSV ler a tabela de batidas).

## 12. Estado atual e como continuar (13/10/2026)

**Onde está o código.** Repositório `dimitrimateus/dashboard_ponto`, branch
`ccr-c6360952-zmkobs` (continuação da `claude/fervent-hamilton-281bul`; push funcionando). O
`Dashboard/index.html` é o painel; as macros ficam em `<Pasta>/<Macro>.bas` e `txt/<Macro>.txt`.

**O que o Dimitri tem em mãos.**
- Pacote `Teste Ponto 11-10.zip` (pastas Macros / Planilhas / Site + LEIA-ME): as 7 macros em
  `.txt`, a planilha `Tratamento Ponto 28.08 até 27.09_2 - completa (com Jornada).xlsx` (ORG
  corrigido + abas Marcações Digitadas, Afastamentos, Intrajornada, Interjornada, Interjornada
  Semanal, Horas Excedentes) e o CSV de referência simulado.
- Depois disso recebeu só o `index.html` atualizado (13/10): tela cheia, filtro na legenda por
  gráfico, selo de período, clique "de onde vem" com nomes que filtram, textos para leigos,
  correções da revisão geral; depois, a rolagem e o cabeçalho fixo da janela "de onde vem".
- 13/10: pacote com as 15 escolhas construídas: `index.html` + `GerarCSVPonto.txt` (lê todas as
  abas de pausas).
- 18/10: pacote `Teste Ponto 18-10.zip`: `index.html` com as tendências, `FormatarPausasTermicas.txt`
  (renomeia a aba pelo mês), `GerarCSVPonto.txt` (igual ao de 13/10) e a planilha
  `... - com 3 meses.xlsx` (digitadas de 3 meses + pausas de jul/ago/set). Números esperados nela:
  os mesmos de antes nas ocorrências; CSV com 3.082 linhas; 1.642 digitadas (416 no ciclo de
  setembro); 156 linhas de pausas (52/53/51).
- Números esperados ao rodar o GerarAbaCSV nessa planilha: 138 ocorrências, 72 colaboradores,
  demora média 5,7 dias (56 tratadas), 422 digitadas, 283 afastamentos (33 atestados, 105 dias),
  210 pessoas no ORG, 163 com horas no Cartão, Intrajornada 21, Interjornada 1, Interjornada
  Semanal 11, Horas Excedentes 72; CSV com 32 colunas e 1.757 linhas.
- **As macros nunca rodaram no Excel dele.** O próximo retorno pode trazer erro de compilação ou
  execução: peça print da mensagem e da linha amarela.

**Escolhas da página de ideias** (https://claude.ai/artifact/SZEwQzRntr66SCDMms4YBG, doc
`escolhas/atual`, campo `marcadas`): **todas construídas** (13/10). Em aberto: as sugestões para
reunião R1–R10 (seção 11, 13/10 depois 6), que esperam ele responder se a reunião é por gestor ou
com todos e se pode usar salário (só o total da equipe).

**Decisões já tomadas (não reabrir sem ele pedir):**
- Check "S" ignora a linha inteira; "Problema horário" não entra; Origem "E" das digitadas não entra.
- Cards de horas extra/falta retirados; aviso vermelho "fora do ORG" na lista retirado.
- Absenteísmo = só atestados nas vistas de atestado; fonte oficial = relatório HRCL006.
  Afastamentos mostram todos os tipos por cor, com opção "Só atestados".
- O relatório de digitadas não diz quem digitou (colaborador × RH).
- Painel em seções com menu ☰; Pausas Térmicas na própria seção.
- Filtros da lateral sempre visíveis ao abrir; arquivo carregado abre na Visão geral.
- Filtro da legenda vale só para o gráfico clicado e não mexe nos números do topo.
- Listas longas (curtas, top de digitadas, banco de horas) inteiras só em tela cheia.
- Clique na barra/fatia/ponto/quadrado abre "de onde vem"; os nomes de cima dessa janela são
  filtros; o resumo da pessoa abre pelo nome na tabela.
- Textos da tela para leigos: sem aba, coluna, macro, código de relatório, Check, ORG,
  Tratamento, DSR. Cores: 5 cores de série + cinza "Outros"; vermelho só para alerta/atestado.
- Gestores também veem as análises de "pegar no pulo".
- Relatórios do sistema com 3 meses (digitadas, afastamentos, jornada, Cartão); pausas uma aba por
  mês; planilha de tratamento e PontoNet só o mês atual (conversa de 14/10 e envio de 18/10).
- **O texto "pegar no pulo" não aparece na tela** (pedido do Dimitri, 18/10): os cards de sinais
  de atenção ficam só com o título. A expressão pode existir em comentários do código.

**Ainda em aberto com ele:** Check "S" nas curtas/auditoria; critério Internos × Externos;
`destino_horas_extra`; absenteísmo completo (acidente, faltas, taxa); sugestões de 07/10 (horas de
saída médica, quem está fora agora, férias por gestor, tempo de casa × atestados); o Tyrone só
aparece no aviso de férias se o HRCL006 for exportado incluindo os gestores.

**Como trabalhar com ele (o que funcionou):**
- Toda resposta em pt-BR, explicando o porquê; dizer direto quando algo entregue estava errado.
- Entregar VBA em `.txt` (CRLF): `sed 's/$/\r/' X/X.bas > txt/X.txt`. Manter `.bas` e `.txt` iguais.
- Antes de entregar: `checar_colisoes_vba.py`, `simular_gerarcsv.py`, abrir o painel no Chromium
  (Playwright), `conferir_painel.py` (18/18), screenshots desktop, 390 e 360 px, procurar erro de
  JavaScript e texto "NaN/undefined". O extrator (`extrair_painel.js`) lê balões `data-tooltip`
  sem a classe `seg` e a legenda `.legend-row` da rosca: não mude esses formatos sem ajustá-lo.
- Relatórios novos do sistema (Senior) chegam como "dump" de impressão: macro `FormatarX` que
  reescreve a aba ativa como tabela; o GerarCSV acha a aba pelo cabeçalho; o painel ganha um
  `tipo_ocorrencia` novo, fora das vistas gerais e do `periodoCompleto`, com seção própria.
- Planilha nova dele: inserir abas com `adicionar_aba_xlsx.py` (nunca salvar com openpyxl).
- Pacotes de teste em zip (Macros / Planilhas / Site + LEIA-ME) quando ele pede "tudo".
- Dados pessoais (nomes, matrículas, salários): nunca no repositório nem em commits; só na conversa.
- Código comentado para leigos; ao mudar código, atualizar os mapas no topo do index.html
  ("O QUE APARECE NA TELA") e do GerarCSVPonto ("MAPA DO MÓDULO") e o histórico (seção 11).
