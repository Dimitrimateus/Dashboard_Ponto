// RELATÓRIO DA CONFERÊNCIA DO PAINEL, EM WORD (11/11/2026)
// ---------------------------------------------------------------------------------
// O que faz: pega o resultado da conferência (o JSON que ferramentas/auditar_csv_painel.py
// grava com --json) e monta um relatório em Word (.docx), seção por seção do painel,
// card por card: o que foi conferido, como o número é calculado (em palavras simples),
// o valor e se o painel bateu com a conta feita à parte.
// O PDF sai do próprio .docx (LibreOffice: soffice --convert-to pdf).
//
// Uso:
//   node relatorio_conferencia.js resultado.json saida.docx "nome do(s) arquivo(s) conferido(s)" "linhas lidas" [teste]
// O 5º argumento "teste" põe na capa o aviso de que a rodada foi com um arquivo inventado.
const fs = require('fs');
const d = require('docx');
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, HeadingLevel, BorderStyle, LevelFormat, Footer, PageNumber, PageBreak } = d;

const [,, entrada, saida, arquivo, linhas, modo] = process.argv;
const RES = JSON.parse(fs.readFileSync(entrada, 'utf8'));
const TESTE = modo === 'teste';
const LARANJA = 'F7931E', PRETO = '1A1A1A', CINZA = '666666', VERDE = '1B8A5A', VERMELHO = 'E31E24';
const hoje = new Date().toLocaleDateString('pt-BR');
const nOk = RES.filter(r => r.ok).length;

// Texto de abertura de cada seção do painel: o que ela mostra, para quem nunca viu.
const INTRO = {
  'Período padrão do painel': 'Ao abrir o arquivo, o painel escolhe sozinho o período do filtro: o ciclo (28 a 27) do cartão ponto mais recente. Desde 11/11 ele vale para as ocorrências curtas do cartão ponto; as outras seções têm o seu próprio seletor de mês.',
  'Marcações digitadas': 'Batidas lançadas à mão no sistema de ponto (relatório "Marcações digitadas"), só as de origem "D". A seção mostra todos os ciclos do arquivo.',
  'Afastamentos': 'Afastamentos do relatório "Histórico de Afastamentos" do sistema (férias, atestados, viagens, cursos...). Cada afastamento vira os dias corridos que cobre.',
  'Jornada': 'Os quatro relatórios de limite de jornada do sistema: horas excedentes (mais de 10h no dia), intrajornada (intervalo curto), interjornada (menos de 11h de descanso) e interjornada semanal (trabalho na folga ou feriado). No fim da seção, a parte "Do cartão ponto": extra e falta no mesmo dia (3 meses) e ocorrências curtas (menos de 15 minutos no mês).',
  'Pausas térmicas': 'Pausas de quem trabalha no frio: o resumo por pessoa do relatório do sistema e, dia a dia, as batidas do relógio (trabalho e pausa alternando; trecho de trabalho acima de 1h40 é o alerta).',
  'PontoNet': 'Pedidos de justificativa de batida feitos no PontoNet. O painel usa só os aprovados (situação INTEGRADO) e só de gestores da lista de equipes.',
  'Banco de horas': 'Horas extras (coluna 50% do cartão ponto) e horas falta (coluna BH do cartão ponto), dia a dia, nos meses escolhidos; e o relatório de banco de horas do sistema para conferência.',
  'Tendências': 'Os mesmos assuntos das outras seções, contados ciclo a ciclo nos 3 últimos meses, para ver se melhorou ou piorou.',
  'Lista de colaboradores': 'A lista de equipes do jeito que está no cadastro: matrícula, colaborador, área, cargo e gestor.',
  'Painel inteiro': 'Funcionamento geral: o painel não pode ter nenhum erro de programa ao abrir o arquivo e desenhar todas as seções.'
};

const p = (t, o = {}) => new Paragraph({ spacing: { after: 120 }, ...o, children: (Array.isArray(t) ? t : [new TextRun(t)]) });
const run = (t, o = {}) => new TextRun({ text: t, ...o });
const bullet = t => new Paragraph({ numbering: { reference: 'bol', level: 0 }, spacing: { after: 80 }, children: Array.isArray(t) ? t : [run(t)] });
const W = [2300, 3700, 2700, 1200];         // larguras das colunas da tabela (somam 9900)
const borda = { style: BorderStyle.SINGLE, size: 4, color: 'D0D0D0' };
const bordas = { top: borda, bottom: borda, left: borda, right: borda };
const cel = (filhos, w, o = {}) => new TableCell({ borders: bordas, width: { size: w, type: WidthType.DXA },
  margins: { top: 60, bottom: 60, left: 90, right: 90 }, ...o, children: filhos });
const cab = ['O que foi conferido', 'Como o número é calculado', 'Valor conferido', 'Resultado'].map((t, i) =>
  cel([p([run(t, { bold: true, color: 'FFFFFF', size: 18 })], { spacing: { after: 0 } })], W[i],
      { shading: { fill: LARANJA, type: ShadingType.CLEAR, color: 'auto' } }));
const pequeno = (t, o = {}) => p([run(t, { size: 17, ...o })], { spacing: { after: 40 } });

function linha(r) {
  const valor = [pequeno(r.esperado.length > 420 ? r.esperado.slice(0, 420) + ' ...' : r.esperado)];
  if (!r.ok) valor.push(pequeno('Painel mostrou: ' + r.painel.slice(0, 300), { color: VERMELHO }));
  if (r.obs) valor.push(pequeno(r.obs, { italics: true, color: CINZA }));
  return new TableRow({ cantSplit: true, children: [
    cel([pequeno(r.card, { bold: true }), pequeno(r.item)], W[0]),
    cel([pequeno(r.como)], W[1]),
    cel(valor, W[2]),
    cel([pequeno(r.ok ? '✓ Confere' : '✗ Diferente', { bold: true, color: r.ok ? VERDE : VERMELHO })], W[3])
  ] });
}

const corpo = [];
// ---- capa
corpo.push(p([run('AuroraCoop · RH · Tratamento de ponto', { color: LARANJA, bold: true, size: 22 })], { spacing: { before: 1800, after: 200 } }));
corpo.push(new Paragraph({ heading: HeadingLevel.TITLE, children: [run('Conferência do Painel de Ponto')] }));
corpo.push(p([run('Relatório de verificação de todos os cálculos, seção por seção e gráfico por gráfico', { size: 26, color: CINZA })], { spacing: { after: 600 } }));
corpo.push(p([run('Data: ', { bold: true }), run(hoje)]));
corpo.push(p([run('Arquivo conferido: ', { bold: true }), run(arquivo)]));
corpo.push(p([run('Linhas lidas: ', { bold: true }), run(linhas)]));
corpo.push(p([run('Resultado: ', { bold: true }), run(nOk + ' de ' + RES.length + ' itens conferem', { bold: true, color: nOk === RES.length ? VERDE : VERMELHO })]));
if (TESTE) corpo.push(p([run('Atenção: esta rodada usou um arquivo de TESTE, com pessoas e números inventados (PESSOA 01, GESTOR ALFA...). Ela prova que as contas do painel estão certas; a versão para arquivo da unidade sai com o arquivo real do mês, rodando as mesmas ferramentas.', { color: VERMELHO, bold: true })], { spacing: { before: 300 } }));
corpo.push(new Paragraph({ children: [new PageBreak()] }));

// ---- método
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('1. Como a conferência foi feita')] }));
corpo.push(p('A pergunta era: os números que o painel mostra estão certos? Para responder sem depender da palavra de quem fez o painel, a conferência usa duas peças separadas, que não compartilham nenhuma linha de código:'));
corpo.push(bullet([run('Leitura da tela. ', { bold: true }), run('Um programa abre o painel no navegador (o mesmo Chrome dos gestores), carrega o arquivo do mês e anota tudo o que aparece: os números do topo de cada seção, o número de cada barra e fatia dos gráficos, as legendas e todas as linhas das tabelas. Listas que no card mostram só os primeiros nomes são lidas em tela cheia, onde aparecem inteiras.')]));
corpo.push(bullet([run('Calculadora independente. ', { bold: true }), run('Outro programa, escrito à parte em outra linguagem (Python), lê o mesmo arquivo e refaz cada conta do zero, seguindo as regras combinadas com o RH (o que entra, o que não entra, como se conta cada coisa).')]));
corpo.push(p('Cada item abaixo compara as duas. "Confere" quer dizer que o painel mostrou exatamente o mesmo número que a calculadora independente chegou. Nos gráficos com muitos nomes, a comparação vale para cada barra mostrada, e a calculadora também verifica que ninguém que deveria aparecer ficou de fora.'));
corpo.push(p('O que esta conferência garante: que o painel conta, soma, agrupa e calcula médias e porcentagens corretamente a partir do arquivo. O que ela não pode garantir: que os relatórios exportados do sistema de ponto estejam corretos na origem. Para isso existem as checagens da seção "Arquivos carregados" do painel (avisa se falta algum relatório ou mês) e a conferência do banco de horas (o saldo do cartão ponto tem que bater com o relatório de banco de horas do sistema).'));

// ---- resumo
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('2. Resumo do resultado')] }));
const secoes = [...new Set(RES.map(r => r.secao))];
const W2 = [5000, 2400, 2500];
corpo.push(new Table({ width: { size: 9900, type: WidthType.DXA }, columnWidths: W2, rows: [
  new TableRow({ tableHeader: true, children: ['Seção do painel', 'Itens conferidos', 'Resultado'].map((t, i) =>
    cel([p([run(t, { bold: true, color: 'FFFFFF', size: 18 })], { spacing: { after: 0 } })], W2[i], { shading: { fill: LARANJA, type: ShadingType.CLEAR, color: 'auto' } })) }),
  ...secoes.map(s => { const l = RES.filter(r => r.secao === s), ok = l.filter(r => r.ok).length;
    return new TableRow({ children: [cel([pequeno(s)], W2[0]), cel([pequeno(String(l.length))], W2[1]),
      cel([pequeno(ok === l.length ? '✓ todos conferem' : (l.length - ok) + ' diferente(s)', { bold: true, color: ok === l.length ? VERDE : VERMELHO })], W2[2])] }); })
] }));
corpo.push(p('', { spacing: { after: 200 } }));
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run('Correções feitas durante a conferência')] }));
corpo.push(p('A primeira rodada achou uma diferença real no painel, já corrigida: no tamanho da equipe ("de N pessoas", usado nos números do topo e na faixa "Mostrando" de cada seção), uma pessoa que está na lista de equipes SEM matrícula (a lista às vezes traz "#N/A" na matrícula) e que tem ocorrência COM matrícula era contada duas vezes, aumentando o tamanho da equipe em 1. Agora o painel reconhece que é a mesma pessoa pelo nome. Não mudava nenhuma contagem, só o "de N pessoas".'));
corpo.push(p('Com o arquivo real apareceu uma segunda, de texto, também corrigida: na seção Banco de horas, os números do topo diziam "220 pessoas na lista de equipes". O 220 está certo, mas é o número de pessoas com banco de horas (216 do relatório do sistema + 4 que só aparecem no cartão ponto); a lista de equipes tem 216. Agora o texto diz "220 pessoas com banco de horas".'));
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run('Pontos de atenção (não são erros, mas é bom saber como lê)')] }));
[
  'Afastamentos, "Ranking de atestados": a coluna Dias soma os dias de ausência informados pelo sistema. No gráfico de evolução e no mapa, todo afastamento conta no mínimo 1 dia (um registro com 0 dias aparece como 1 dia). Por isso, se o sistema mandar um atestado com 0 dias, o ranking e a evolução podem diferir em 1.',
  'Afastamentos, seletor "Últimos 3 meses": mostra todos os ciclos do arquivo com algum dia de afastamento até o fim do cartão ponto. Com o relatório de 3 meses, um afastamento que começou antes (ex.: férias iniciadas no mês anterior) pode criar um 4º ciclo na tabela.',
  'PontoNet, "Tempo do gestor para avaliar, por gestor": um gestor cujos pedidos aprovados não trazem a contagem de dias do gestor aparece com 0 dias, e não "sem dado".',
  'Afastamentos e atestados da tabela de evolução contam no ciclo em que COMEÇAM; os dias contam no ciclo em que cada dia cai. Um atestado de 25/10 a 30/10 conta como 1 atestado em outubro, com 3 dias em outubro e 3 em novembro.',
  'PontoNet, "Por motivo": o painel escreve os motivos em palavras simples e junta os que falam de esquecimento, inclusive "Esquecimento crachá", em "Esquecimento". Se o RH preferir o crachá separado, é uma mudança simples.',
  'Banco de horas, conferência com o relatório do sistema: nas pessoas em que o banco calculado pelo cartão ponto não bate com o relatório de banco de horas, a diferença está nos dados de origem (o cartão ponto e o relatório não dizem a mesma coisa), não numa conta do painel. O painel mostra essas pessoas em vermelho para o RH conferir no sistema.',
  'Banco de horas: o painel só parte do zero no começo do período (nov–fev, mar–jun, jul–out) quando o arquivo tem o cartão ponto de todos os meses do período. Faltando um mês, ele se ancora no relatório de banco de horas do sistema e avisa em vermelho.'
].forEach(t => corpo.push(bullet(t)));

// ---- seção a seção
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('3. Conferência seção por seção')] }));
secoes.forEach((s, i) => {
  corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [run('3.' + (i + 1) + ' ' + s)] }));
  if (INTRO[s]) corpo.push(p([run(INTRO[s], { color: CINZA })]));
  corpo.push(new Table({ width: { size: 9900, type: WidthType.DXA }, columnWidths: W,
    rows: [new TableRow({ tableHeader: true, children: cab }), ...RES.filter(r => r.secao === s).map(linha)] }));
  corpo.push(p('', { spacing: { after: 160 } }));
});

// ---- código comentado e como repetir
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('4. O código do painel e das macros')] }));
corpo.push(p('Além da conferência dos números, todo o código foi revisado e comentado em português simples, para que tanto uma pessoa leiga quanto alguém que venha a mexer no painel entenda o que cada parte faz:'));
corpo.push(bullet('Painel (index.html): no começo do código há um "Guia rápido" com os limites que o RH pode querer mudar (alerta de extra e falta a partir de 10, esquecimento a partir de 5, trecho de trabalho de 1h40, 3 dias para o gestor avaliar, 10 horas de jornada, 20 nomes por lista...) e onde mudar cada um. Toda função tem um comentário dizendo o que faz e por quê.'));
corpo.push(bullet('Macros do Excel (arquivos .txt e .bas): cada rotina tem um comentário no topo explicando o que lê, o que escreve e em que ordem.'));
corpo.push(bullet('Ferramentas de conferência: as duas peças descritas no item 1 ficam guardadas junto com o painel e podem ser rodadas de novo a cada mês.'));
corpo.push(new Paragraph({ heading: HeadingLevel.HEADING_1, children: [run('5. Como repetir esta conferência')] }));
corpo.push(p('Qualquer mês, com o arquivo novo, a conferência pode ser refeita do mesmo jeito e gera este mesmo relatório com os números do mês:'));
['Ler a tela: ferramentas/extrair_tudo_painel.js (abre o painel e anota tudo o que aparece).',
 'Refazer as contas: ferramentas/auditar_csv_painel.py (calculadora independente; mostra OK ou FALHA em cada item).',
 'Montar o relatório: ferramentas/relatorio_conferencia.js (este documento, em Word; o PDF sai do Word).'].forEach(t => corpo.push(bullet(t)));

const doc = new Document({
  creator: 'RH AuroraCoop', title: 'Conferência do Painel de Ponto',
  styles: { default: { document: { run: { font: 'Arial', size: 21 } } },
    paragraphStyles: [
      { id: 'Title', name: 'Title', basedOn: 'Normal', run: { font: 'Arial', size: 52, bold: true, color: PRETO }, paragraph: { spacing: { before: 0, after: 200 }, keepNext: true } },
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Arial', size: 30, bold: true, color: PRETO }, paragraph: { spacing: { before: 360, after: 160 }, keepNext: true, outlineLevel: 0, border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: LARANJA, space: 4 } } } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: 'Arial', size: 25, bold: true, color: 'C46A00' }, paragraph: { spacing: { before: 280, after: 120 }, keepNext: true, outlineLevel: 1 } }
    ] },
  numbering: { config: [{ reference: 'bol', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] }] },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [
      run('Conferência do Painel de Ponto · página ', { size: 16, color: CINZA }), new TextRun({ children: [PageNumber.CURRENT], size: 16, color: CINZA })] })] }) },
    children: corpo
  }]
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(saida, b); console.log('ok:', saida, nOk + '/' + RES.length); });
