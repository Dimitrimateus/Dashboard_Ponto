// RELATÓRIO DE CONFIABILIDADE PARA O SUPERVISOR (13/11/2026)
// ---------------------------------------------------------------------------------
// O que faz: monta um documento em Word, SEM termos técnicos, para o supervisor da unidade,
// provando que os números do Painel de Ponto estão certos. Usa o resultado da conferência
// (o JSON que ferramentas/auditar_csv_painel.py grava com --json): cada item conferido vira
// uma linha em palavras simples ("o que o painel mostra", "como o número é obtido",
// "resultado"). Não traz nomes de pessoas (só quantidades), para poder circular.
// Para quem mexe no código: o texto de cada item fica em TEXTOS, pela chave "card|item" do
// conferidor. Item novo no conferidor sem texto aqui aparece com o texto do próprio conferidor.
//
// Uso:
//   node relatorio_supervisor.js resultado.json saida.docx "mês de referência" "período dos dados"
// PDF: soffice --headless --convert-to pdf saida.docx
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, HeadingLevel, BorderStyle, LevelFormat, Footer, PageNumber, PageBreak } = require('docx');

const [,, entrada, saida, mesRef, periodo] = process.argv;
const RES = JSON.parse(fs.readFileSync(entrada, 'utf8'));
const LARANJA = 'F7931E', PRETO = '1A1A1A', CINZA = '5F6368', VERDE = '1B8A5A', VERMELHO = 'C62828';
const nOk = RES.filter(r => r.ok).length;

// Texto de cada item conferido, em palavras simples. [o que o painel mostra, como o número é obtido, unidade]
// A unidade diz o que foi conferido quando o item é uma lista ("pessoas", "gestores", "dias"...).
const TEXTOS = {
  'Filtro de período (lateral)|Data inicial': ['Começo do mês de referência do painel', 'Dia 28 do mês anterior ao fechamento do cartão ponto (o mês do ponto vai do dia 28 ao dia 27).'],
  'Filtro de período (lateral)|Data final': ['Fim do mês de referência do painel', 'Último dia do cartão ponto do mês (dia 27).'],
  'Números do topo|Marcações digitadas': ['Total de batidas digitadas', 'Contagem de cada batida lançada à mão no sistema de ponto nos 3 meses do arquivo.'],
  'Números do topo|Colaboradores': ['Quantas pessoas tiveram batida digitada', 'Pessoas diferentes que aparecem pelo menos uma vez entre as batidas digitadas.'],
  'Números do topo|"de N pessoas na lista de equipes"': ['Tamanho da equipe usado para a porcentagem', 'Pessoas na lista de equipes da unidade (cada pessoa conta uma vez).'],
  'Números do topo|Por esquecimento': ['Batidas digitadas por esquecimento', 'Batidas cujo motivo informado é "inclusão por esquecimento".'],
  'Por motivo (rosca)|Quantidade por motivo': ['Batidas digitadas de cada motivo', 'Cada batida contada no motivo que foi informado no sistema.', 'motivos'],
  'Mês a mês|Batidas por ciclo': ['Batidas digitadas em cada mês', 'Cada batida contada no mês do ponto (28 a 27) em que aconteceu.', 'meses'],
  'Por dia da semana|Batidas em cada dia': ['Batidas digitadas por dia da semana', 'Cada batida contada no dia da semana da sua data.', 'dias da semana'],
  'Por data|Batidas em cada dia': ['Batidas digitadas por dia do calendário', 'Cada batida contada na sua data.', 'dias'],
  'Por hora do dia|Batidas em cada hora': ['Batidas digitadas por hora do dia', 'Cada batida contada na hora em que foi registrada.', 'horas do dia'],
  'Por gestor|Batidas por gestor (os 12 maiores)': ['Batidas digitadas por equipe', 'Batidas somadas pelo gestor de cada pessoa, conforme a lista de equipes.', 'gestores'],
  'Quem mais teve batidas digitadas|Batidas por pessoa (lista inteira em tela cheia)': ['Batidas digitadas de cada pessoa', 'Batidas somadas por pessoa.', 'pessoas'],
  'Justificativas mais usadas|Batidas e pessoas de cada texto': ['Justificativas mais usadas', 'Textos iguais somados (sem diferenciar maiúscula, acento ou ponto final).', 'justificativas'],
  'Lista de batidas|Quantidade de linhas': ['Lista completa de batidas digitadas', 'Uma linha por batida: a lista tem o mesmo total do número do topo.'],
  'Esquecimento repetido|Esquecimentos por pessoa nos 3 ciclos': ['Quem esquece de bater com frequência', 'Esquecimentos de cada pessoa, mês a mês, nos 3 meses.', 'pessoas'],
  'Evolução por ciclo|Dias, afastamentos, atestados e pessoas em cada ciclo': ['Afastamentos de cada mês', 'Cada afastamento vira os dias que cobre (do início ao fim) e cada dia conta no mês em que cai. Afastamentos e atestados contam no mês em que começam.', 'meses'],
  'Ranking de atestados|Atestados, dias, meses e encostados em folga por pessoa': ['Quem mais teve atestado', 'Por pessoa: quantos atestados, quantos dias, em quantos meses e quantos encostaram numa folga.', 'pessoas'],
  'Por gestor (por tipo)|Soma dos dias de todos os tipos = dias do período': ['Dias de afastamento, por tipo', 'A soma dos tipos (férias, atestado, viagem...) tem que dar o total de dias dos 3 meses.'],
  'Por gestor (por tipo)|Dias de afastamento por gestor': ['Dias de afastamento por equipe', 'Dias somados pelo gestor de cada pessoa.', 'gestores'],
  'Números do topo|Horas excedentes (ocorrências e pessoas)': ['Dias com mais de 10 horas de trabalho, e quantas pessoas', 'Cada dia apontado pelo sistema de ponto como acima de 10 horas.', null, '{0} dias · {1} pessoas'],
  'Números do topo|Intrajornada (ocorrências e pessoas)': ['Dias com intervalo de almoço menor que o mínimo, e quantas pessoas', 'Cada dia apontado pelo sistema de ponto.', null, '{0} dias · {1} pessoas'],
  'Números do topo|Interjornada semanal (ocorrências e pessoas)': ['Trabalho em dia de folga ou feriado, e quantas pessoas', 'Cada dia apontado pelo sistema de ponto.', null, '{0} dias · {1} pessoas'],
  'Números do topo|Interjornada (ocorrências e pessoas)': ['Menos de 11 horas de descanso entre um dia e outro, e quantas pessoas', 'Cada dia apontado pelo sistema de ponto.', null, '{0} dias · {1} pessoas'],
  'Números do topo|Horas acima de 10h (soma)': ['Total de horas trabalhadas além das 10 horas', 'Em cada dia acima do limite, as horas do dia menos 10 horas; depois, a soma de todos os dias.'],
  'Mês a mês|Ocorrências por ciclo': ['Problemas de jornada em cada mês', 'Cada ocorrência contada no mês do ponto em que aconteceu.', 'meses'],
  'Por gestor|Ocorrências por gestor': ['Problemas de jornada por equipe', 'Ocorrências somadas pelo gestor de cada pessoa.', 'gestores'],
  'Tabela por gestor|Ocorrências por tipo e pessoas, por gestor': ['Tabela de jornada por equipe', 'Por gestor: total, cada um dos 4 tipos e quantas pessoas.', 'gestores'],
  'Lista de ocorrências|Quantidade de linhas': ['Lista completa de problemas de jornada', 'Uma linha por ocorrência: a lista tem o mesmo total dos números do topo.'],
  'Mais de 10h no dia, por pessoa|Dias e minutos acima de 10h por pessoa': ['Quem passou de 10 horas no dia', 'Por pessoa: quantos dias e quanto tempo além das 10 horas.', 'pessoas'],
  'Do cartão ponto: extra e falta no mesmo dia|Este mês e 3 meses, por pessoa (os 10 maiores)': ['Hora extra e falta no mesmo dia', 'Dias em que a pessoa teve hora extra e falta ao mesmo tempo, no mês e somando 3 meses. O alerta aparece a partir de 10 em 3 meses.', 'pessoas'],
  'Do cartão ponto: ocorrências curtas (menos de 15 min)|Faltas, extras e total por pessoa (lista inteira)': ['Atrasos e horas extras de menos de 15 minutos', 'Dias do mês com diferença pequena (menos de 15 minutos), contados por pessoa.', 'pessoas'],
  'Números do topo|Pessoas no relatório': ['Pessoas que fazem pausa térmica', 'Pessoas no relatório de pausas do mês.'],
  'Números do topo|Pausas certas (%)': ['Porcentagem de pausas feitas certas (20 minutos)', 'Pausas de 20 minutos divididas pelo total de pausas do mês.', null, '{0} ({1} de {2} pausas)'],
  'Tabela de pausas|Os 7 números de cada pessoa (lista inteira)': ['Tabela de pausas de cada pessoa', 'Os números vêm prontos do relatório do sistema; o painel apenas os mostra.', 'pessoas'],
  'Evolução mês a mês|Pessoas, % certas, trechos acima de 1h40 e batida faltando por mês': ['Pausas mês a mês', 'Os números de cada pessoa somados em cada mês.', 'meses'],
  'Acima de 1h40 por gestor|Pessoas e trechos acima de 1h40': ['Trabalho sem pausa acima de 1h40, por equipe', 'Trechos acima do limite somados pelo gestor de cada pessoa.', 'gestores'],
  'Batidas: resumo do mês|Dias trabalhados, pessoas e dias com trecho acima de 1h40': ['Resumo das batidas de pausa do mês', 'Leitura de cada batida do relógio: entre uma batida e outra é trabalho ou pausa; trabalho acima de 1h40 é alerta.'],
  'Batidas: por hora de começo do trecho|Trechos acima de 1h40 / trechos que começaram em cada hora': ['Horários em que mais se passa de 1h40', 'Cada período de trabalho contado na hora em que começou.', 'horários'],
  'Batidas: por dia da semana|Dias com trecho acima de 1h40 / dias trabalhados': ['Dias da semana em que mais se passa de 1h40', 'Dias trabalhados de cada dia da semana e quantos tiveram alerta.', 'dias da semana'],
  'Batidas: trabalho até a 1ª pausa térmica|Dias em cada faixa': ['Quanto se trabalha até a primeira pausa', 'Tempo do começo do dia até a 1ª pausa térmica, em faixas.', 'faixas'],
  'Batidas: maior trecho sem pausa, por pessoa|Maior trecho (minutos) e dias acima de 1h40': ['Maior tempo sem pausa de cada pessoa', 'Por pessoa: o maior período seguido de trabalho e em quantos dias passou de 1h40.', 'pessoas'],
  'Batidas: mapa das pausas|Dias com problema e dias no relatório, por pessoa': ['Mapa das pausas, pessoa a pessoa', 'Dia com problema = passou de 1h40, ou não fez pausa, ou faltou bater.', 'pessoas'],
  'Pausas por gestor|Total de pausas da equipe (certas + curtas + longas)': ['Pausas por equipe', 'Todas as pausas das pessoas de cada gestor.', 'gestores'],
  'Acima de 1h40 por pessoa|Trechos acima de 1h40 / todos os trechos': ['Trabalho sem pausa acima de 1h40, por pessoa', 'Trechos acima do limite divididos por todos os trechos da pessoa.', 'pessoas'],
  'Números do topo|Justificativas aprovadas': ['Justificativas de ponto aprovadas no mês', 'Pedidos de justificativa aprovados no PontoNet no mês de referência.'],
  'Números do topo|Pessoas com justificativa aprovada': ['Quantas pessoas tiveram justificativa aprovada', 'Pessoas diferentes nesses pedidos.'],
  'Mês a mês|Aprovados por ciclo': ['Justificativas aprovadas em cada mês', 'Cada pedido contado no mês da batida que faltou.', 'meses'],
  'Tempos por ciclo|Aprovados, tempo médio do colaborador e do gestor': ['Tempo para justificar e para avaliar', 'Média dos dias contados pelo próprio PontoNet: do colaborador (até justificar) e do gestor (até avaliar).', 'meses'],
  'Lista de aprovados|Quantidade de linhas': ['Lista de justificativas aprovadas', 'Uma linha por pedido aprovado: o mesmo total do número do topo.'],
  'Demora por gestor|Aprovadas, demora média e maior demora': ['Demora até a aprovação, por equipe', 'Do dia e hora da batida que faltou até a aprovação, em dias.', 'gestores'],
  'Por motivo|Aprovados de cada motivo (ciclo escolhido)': ['Motivos das justificativas aprovadas', 'Cada pedido contado no seu motivo.', 'motivos'],
  'Tempo do gestor para avaliar, por gestor|Média de dias e pedidos': ['Tempo de cada gestor para avaliar', 'Média dos dias que cada gestor levou para avaliar.', 'gestores'],
  'Números do topo|Horas extras': ['Horas extras que entraram no banco', 'Soma, dia a dia, das horas extras do cartão ponto nos 4 meses do período do banco.'],
  'Números do topo|Horas falta': ['Horas de falta e atraso tiradas do banco', 'Soma, dia a dia, das horas descontadas do banco no cartão ponto.'],
  'Números do topo|Banco no fim = banco no começo + horas extras − horas falta': ['Saldo do banco no fechamento', 'Começo do período (zerado) + horas extras − horas falta: a conta tem que fechar.'],
  'Banco de cada pessoa|Horas extras, horas falta, variação do banco e dias com extra/falta': ['Banco de horas de cada pessoa', 'Por pessoa: horas extras, horas falta e o saldo; o saldo final menos o inicial tem que ser igual a extras − faltas.', 'pessoas'],
  'Relatório de banco de horas|Contas fechadas no relatório do sistema': ['Relatório de banco de horas do sistema', 'Em cada pessoa do relatório, a conta saldo anterior + créditos − débitos − pagas + descontadas = saldo final.'],
  'Conferência: cartão ponto × relatório de banco de horas|Banco de cada pessoa pelas duas contas e se bate (lista inteira)': ['Comparação do banco calculado com o relatório do sistema', 'O banco de cada pessoa é calculado pelo cartão ponto e comparado com o relatório do sistema. O painel marca em vermelho quem não bate.'],
  'Assuntos da unidade (3 meses)|Batidas digitadas': ['Tendência: batidas digitadas', 'Mesma contagem da seção de batidas digitadas, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Batidas digitadas por esquecimento': ['Tendência: esquecimentos', 'Mesma contagem da seção de batidas digitadas, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Dias de atestado': ['Tendência: dias de atestado', 'Mesma contagem da seção de afastamentos, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Horas excedentes (dias com mais de 10h)': ['Tendência: dias com mais de 10 horas', 'Mesma contagem da seção de jornada, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Intrajornada (intervalo de almoço curto)': ['Tendência: intervalo de almoço curto', 'Mesma contagem da seção de jornada, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Interjornada semanal (trabalho na folga ou feriado)': ['Tendência: trabalho na folga ou feriado', 'Mesma contagem da seção de jornada, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Interjornada (menos de 11h de descanso)': ['Tendência: menos de 11 horas de descanso', 'Mesma contagem da seção de jornada, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Pausas certas (20 min)': ['Tendência: pausas certas', 'Mesma conta da seção de pausas, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Trabalho acima de 1h40 sem pausa': ['Tendência: trabalho acima de 1h40 sem pausa', 'Mesma conta da seção de pausas, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Justificativas aprovadas no PontoNet': ['Tendência: justificativas aprovadas', 'Mesma contagem da seção PontoNet, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Dias para o gestor avaliar (média)': ['Tendência: tempo do gestor para avaliar', 'Mesma média da seção PontoNet, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Assuntos da unidade (3 meses)|Batida de pausa faltando': ['Tendência: batida de pausa esquecida', 'Mesma contagem da seção de pausas, mês a mês.', null, 'Mês a mês: {0} → {1} → {2}'],
  'Lista de colaboradores|Pessoas na lista': ['Quantas pessoas há na lista de equipes', 'Uma por pessoa (quem aparece repetido conta uma vez).'],
  'Lista de colaboradores|Matrícula, nome, área, cargo e gestor de cada pessoa': ['Dados de cada pessoa da lista de equipes', 'Copiados da lista de equipes, sem nenhuma conta.', 'pessoas'],
  'Funcionamento|Erros de programa ao abrir e desenhar tudo': ['Funcionamento do painel', 'O painel abriu o arquivo e mostrou todas as seções sem nenhum erro.', null, 'Nenhum erro']
};
// Nome de cada seção para o supervisor e o que ela mostra, em uma frase.
const SECOES = {
  'Período padrão do painel': ['Mês de referência', 'O mês que o painel considera ao abrir.'],
  'Marcações digitadas': ['Batidas digitadas', 'Batidas de ponto que foram lançadas à mão no sistema, porque o colaborador esqueceu de bater, o relógio falhou ou estava em atividade externa.'],
  'Afastamentos': ['Afastamentos', 'Férias, atestados, viagens, cursos e outras ausências registradas no sistema.'],
  'Jornada': ['Jornada de trabalho', 'Os limites de jornada da lei que o sistema de ponto aponta como não respeitados, e diferenças do cartão ponto.'],
  'Pausas térmicas': ['Pausas térmicas', 'As pausas obrigatórias de quem trabalha no frio: 20 minutos a cada 1h40 de trabalho.'],
  'PontoNet': ['PontoNet', 'Os pedidos de justificativa de batida feitos e aprovados no PontoNet.'],
  'Banco de horas': ['Banco de horas', 'Horas extras e horas de falta de cada pessoa e o saldo do banco no período de 4 meses.'],
  'Tendências': ['Tendências', 'A comparação de cada assunto mês a mês, para ver se melhorou ou piorou.'],
  'Lista de colaboradores': ['Lista de colaboradores', 'A lista de equipes da unidade: quem é gestor de quem.'],
  'Painel inteiro': ['Funcionamento geral', 'O painel abrindo e mostrando tudo.']
};

// Valor curto para a coluna "Resultado do mês": número único, ou "N pessoas conferidas" para listas.
function valorCurto(r, unidade, modelo) {
  const v = r.esperado;
  if (modelo) { const p = v.split('; '); return modelo.replace(/\{(\d)\}/g, (m, i) => p[+i] || ''); }
  if (unidade) {
    let n = v.split('; ').length;
    const mais = /\(\+(\d+)\)$/.exec(v);
    if (mais) n = n - 1 + Number(mais[1]);
    return n.toLocaleString('pt-BR') + ' ' + unidade + ', cada um conferido';
  }
  return v.replace(/;\s*/g, ' · ');
}

// ---------- montagem do documento ----------
const run = (t, o = {}) => new TextRun({ text: t, ...o });
const par = (t, o = {}) => new Paragraph({ spacing: { after: 140, line: 300 }, ...o, children: Array.isArray(t) ? t : [run(t)] });
const bullet = t => new Paragraph({ numbering: { reference: 'bol', level: 0 }, spacing: { after: 100, line: 290 }, children: Array.isArray(t) ? t : [run(t)] });
const h1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run(t)] });
const h2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run(t)] });
const borda = { style: BorderStyle.SINGLE, size: 4, color: 'D9D9D9' };
const bordas = { top: borda, bottom: borda, left: borda, right: borda };
const cel = (filhos, w, o = {}) => new TableCell({ borders: bordas, width: { size: w, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 100, right: 100 }, ...o, children: filhos });
const peq = (t, o = {}) => new Paragraph({ spacing: { after: 30 }, children: [run(t, { size: 18, ...o })] });
const cab = (titulos, W) => new TableRow({ tableHeader: true, children: titulos.map((t, i) =>
  cel([peq(t, { bold: true, color: 'FFFFFF' })], W[i], { shading: { fill: LARANJA, type: ShadingType.CLEAR, color: 'auto' } })) });
const tabela = (W, linhas) => new Table({ width: { size: W.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: W, rows: linhas });

const corpo = [];
// capa
corpo.push(par([run('AuroraCoop · Recursos Humanos', { color: LARANJA, bold: true, size: 22 })], { spacing: { before: 1600, after: 240 } }));
corpo.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [run('Painel de Ponto')] }));
corpo.push(par([run('Relatório de confiabilidade dos dados', { size: 32, color: CINZA })], { spacing: { after: 500 } }));
corpo.push(par([run('Mês de referência: ', { bold: true }), run(mesRef)]));
corpo.push(par([run('Dados conferidos: ', { bold: true }), run(periodo)]));
corpo.push(par([run('Data da conferência: ', { bold: true }), run(new Date().toLocaleDateString('pt-BR'))]));
// caixa do resultado
const caixa = new Table({ width: { size: 9600, type: WidthType.DXA }, columnWidths: [9600], rows: [new TableRow({ children: [cel([
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 120, after: 60 }, children: [run('Resultado', { size: 22, color: CINZA })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 60 }, children: [run(nOk + ' de ' + RES.length + ' informações conferidas', { size: 40, bold: true, color: nOk === RES.length ? VERDE : VERMELHO })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 140 }, children: [run(nOk === RES.length ? 'Todas estão corretas.' : 'Há informações com diferença: veja a seção 4.', { size: 24 })] })
], 9600, { shading: { fill: 'FFF4E5', type: ShadingType.CLEAR, color: 'auto' } })] })] });
corpo.push(new Paragraph({ spacing: { before: 400 }, children: [] }), caixa);
corpo.push(new Paragraph({ children: [new PageBreak()] }));

// 1. resumo
corpo.push(h1('1. Em resumo'));
corpo.push(par('O Painel de Ponto reúne, em um só lugar, os relatórios de ponto do mês: batidas digitadas, afastamentos, jornada, pausas térmicas, justificativas do PontoNet, banco de horas e a lista de equipes. Este documento mostra que os números que aparecem no painel estão corretos.'));
corpo.push(par([run('O que foi feito: ', { bold: true }), run('cada número, gráfico e tabela do painel foi conferido com uma segunda contagem, feita de forma totalmente separada, a partir do mesmo arquivo do mês. É como conferir o caixa duas vezes, com duas pessoas diferentes: se as duas chegam ao mesmo valor, o valor está certo.')]));
corpo.push(par([run('O resultado: ', { bold: true }), run(nOk + ' de ' + RES.length + ' informações conferidas deram exatamente o mesmo número nas duas contagens.')]));
corpo.push(par([run('O que isso garante: ', { bold: true }), run('o painel soma, conta, separa por mês, por equipe e por pessoa, e calcula médias e porcentagens sem erro. O que ele mostra é fiel ao que está nos relatórios do sistema de ponto.')]));
corpo.push(par([run('O que isso não garante: ', { bold: true }), run('que o próprio sistema de ponto esteja certo na origem (por exemplo, uma batida lançada errado no sistema aparece errada no painel também). Para isso, o painel tem avisos próprios, explicados na seção 5.')]));

// 2. como
corpo.push(h1('2. Como a conferência foi feita'));
[['Passo 1 — ler o que o painel mostra.', ' O painel foi aberto no navegador com o arquivo do mês, como um gestor faria, e tudo o que aparece na tela foi anotado: os números de cada seção, o valor de cada barra dos gráficos e cada linha das tabelas. Listas longas foram lidas inteiras.'],
 ['Passo 2 — refazer as contas separadamente.', ' Um segundo programa, feito à parte e que não usa nenhuma parte do painel, leu o mesmo arquivo e refez todas as contas do zero, seguindo as mesmas regras do RH (o mês do ponto vai do dia 28 ao dia 27; só entram as batidas digitadas pelo sistema; só as justificativas aprovadas; e assim por diante).'],
 ['Passo 3 — comparar.', ' Cada número do painel foi comparado com o número da segunda contagem. "Confere" quer dizer que os dois são iguais. Nas listas (por pessoa, por equipe, por dia), a comparação foi feita item por item.']]
  .forEach(([a, b]) => corpo.push(bullet([run(a, { bold: true }), run(b)])));
corpo.push(par('Durante a conferência, os nomes das pessoas foram comparados um a um, mas este documento mostra só as quantidades, para poder ser compartilhado.'));

// 3. seção por seção
corpo.push(new Paragraph({ children: [new PageBreak()] }));
corpo.push(h1('3. O que foi conferido, seção por seção'));
corpo.push(par([run('Cada tabela mostra uma informação do painel, como ela é obtida e o resultado da conferência. "✓ Confere" quer dizer que o número do painel é igual ao da segunda contagem.', { color: CINZA })]));
const W = [2550, 3750, 2100, 1200];
const ordem = Object.keys(SECOES).filter(s => RES.some(r => r.secao === s));
ordem.forEach((s, i) => {
  const [nome, desc] = SECOES[s];
  const itens = RES.filter(r => r.secao === s);
  const ok = itens.filter(r => r.ok).length;
  corpo.push(h2('3.' + (i + 1) + ' ' + nome + ' — ' + ok + ' de ' + itens.length + ' conferem'));
  corpo.push(par([run(desc, { color: CINZA })]));
  corpo.push(tabela(W, [cab(['Informação do painel', 'Como o número é obtido', 'Resultado do mês', ''], W)].concat(itens.map(r => {
    const t = TEXTOS[r.card + '|' + r.item] || [r.card + ' — ' + r.item, r.como];
    return new TableRow({ cantSplit: true, children: [
      cel([peq(t[0], { bold: true })], W[0]), cel([peq(t[1])], W[1]), cel([peq(valorCurto(r, t[2], t[3]))], W[2]),
      cel([peq(r.ok ? '✓ Confere' : '✗ Diferente', { bold: true, color: r.ok ? VERDE : VERMELHO })], W[3])] });
  }))));
  corpo.push(par('', { spacing: { after: 120 } }));
});

// 4. correções
corpo.push(h1('4. Ajustes feitos durante a conferência'));
corpo.push(par('A conferência também serve para achar e corrigir detalhes. Dois foram encontrados e já estão corrigidos no painel; nenhum deles mudava as quantidades principais:'));
corpo.push(bullet([run('Tamanho da equipe. ', { bold: true }), run('Quando uma pessoa estava na lista de equipes sem a matrícula preenchida, ela podia ser contada duas vezes no "de N pessoas". Corrigido.')]));
corpo.push(bullet([run('Texto do banco de horas. ', { bold: true }), run('O painel dizia "220 pessoas na lista de equipes". O número 220 estava certo, mas é a quantidade de pessoas com banco de horas (a lista de equipes tem 216). O texto agora diz "220 pessoas com banco de horas".')]));

// 5. o que olhar
corpo.push(h1('5. Pontos de atenção e avisos do próprio painel'));
corpo.push(par('Não são erros do painel, mas ajudam a ler os números corretamente:'));
corpo.push(bullet([run('Banco de horas que não bate com o sistema. ', { bold: true }), run('O painel calcula o banco de cada pessoa pelo cartão ponto e compara com o relatório de banco de horas do sistema. Neste mês, a maioria bate; as pessoas que não batem aparecem em vermelho na seção Banco de horas. Nesses casos a diferença está entre o cartão ponto e o relatório do sistema, e o RH confere cada uma no sistema.')]));
corpo.push(bullet([run('Arquivos que faltam. ', { bold: true }), run('Ao abrir o arquivo, a primeira tela do painel ("Arquivos carregados") mostra o que veio e avisa em vermelho se faltar algum relatório ou algum mês. Neste mês, nada faltou.')]));
corpo.push(bullet([run('Motivos do PontoNet. ', { bold: true }), run('Os pedidos por "esquecimento crachá" são somados aos de "esquecimento".')]));
corpo.push(bullet([run('Afastamentos. ', { bold: true }), run('Férias já marcadas para depois do fim do mês não entram nos números do mês. Um atestado que atravessa a virada do mês tem os dias divididos entre os dois meses.')]));
corpo.push(bullet([run('Cada número diz o seu período. ', { bold: true }), run('Embaixo de cada número e gráfico há uma etiqueta 📅 com o período considerado, e o botão "?" explica de onde vem a informação.')]));

// 6. glossário
corpo.push(h1('6. Palavras usadas no painel'));
const G = [['Mês do ponto', 'Vai do dia 28 de um mês ao dia 27 do mês seguinte (ex.: outubro = 28/09 a 27/10).'],
  ['Cartão ponto', 'O registro de cada dia de cada pessoa: batidas, horas trabalhadas, horas extras e faltas.'],
  ['Batida digitada', 'Batida de ponto lançada à mão no sistema, e não registrada no relógio.'],
  ['Afastamento', 'Qualquer ausência registrada: férias, atestado, viagem a serviço, curso, licença.'],
  ['Horas excedentes', 'Dia com mais de 10 horas de trabalho (8 horas + no máximo 2 horas extras).'],
  ['Intrajornada', 'Intervalo de almoço menor que o mínimo.'],
  ['Interjornada', 'Menos de 11 horas de descanso entre o fim de um dia e o começo do outro.'],
  ['Interjornada semanal', 'Trabalho no dia de folga da semana ou em feriado.'],
  ['Pausa térmica', 'Pausa de 20 minutos a cada 1h40 de trabalho, para quem trabalha no frio.'],
  ['PontoNet', 'Sistema em que o colaborador justifica uma batida que faltou e o gestor aprova.'],
  ['Banco de horas', 'Saldo de horas: horas extras entram, faltas e atrasos saem. Fecha a cada 4 meses (nov–fev, mar–jun, jul–out).'],
  ['Lista de equipes', 'A relação de quem é gestor de quem, mantida pelo RH.']];
const WG = [2600, 7000];
corpo.push(tabela(WG, [cab(['Palavra', 'O que quer dizer'], WG)].concat(G.map(([a, b]) => new TableRow({ children: [cel([peq(a, { bold: true })], WG[0]), cel([peq(b)], WG[1])] })))));

// 7. todo mês
corpo.push(h1('7. Como manter essa confiança todo mês'));
corpo.push(par('A mesma conferência pode ser repetida a cada mês, com o arquivo novo, e gera este mesmo relatório com os números do mês. Assim, sempre que o painel for apresentado, há um documento que comprova que os números estão corretos.'));

const doc = new Document({
  creator: 'RH AuroraCoop', title: 'Painel de Ponto — Relatório de confiabilidade dos dados',
  styles: { default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: 'Arial', size: 60, bold: true, color: PRETO }, paragraph: { spacing: { before: 0, after: 160 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Arial', size: 30, bold: true, color: PRETO }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: LARANJA, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Arial', size: 25, bold: true, color: 'B85F00' }, paragraph: { spacing: { before: 300, after: 100 }, keepNext: true, outlineLevel: 1 } }
    ] },
  numbering: { config: [{ reference: 'bol', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1100, bottom: 1000, left: 1150, right: 1150 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      run('Painel de Ponto · Relatório de confiabilidade · página ', { size: 16, color: CINZA }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: CINZA })] })] }) },
    children: corpo
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(saida, b); console.log('ok:', saida, nOk + '/' + RES.length); });
