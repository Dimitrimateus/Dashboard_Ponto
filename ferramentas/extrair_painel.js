// Abre o index.html no Chromium (Playwright), carrega um CSV e grava em JSON o que cada
// card/gráfico/tabela mostra. Uso:
//   NODE_PATH=$(npm root -g) node extrair_painel.js /caminho/index.html /caminho/dados.csv painel.json
// Onde se encaixa: passo de conferência (CLAUDE.md, seção 10). O JSON gerado aqui é lido
// pelo conferir_painel.py, que compara cada número com a planilha. Se um id de elemento do
// painel mudar, a leitura correspondente abaixo precisa mudar junto.
// errs guarda erros de JavaScript e alertas da página: tem que sair vazio.
const { chromium } = require('playwright');
(async () => {
  const [,, html, csv, out] = process.argv;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 1000 } });
  const errs=[]; p.on('pageerror', e=>errs.push(String(e))); p.on('dialog', d=>{errs.push('dialog:'+d.message()); d.dismiss();});
  await p.goto('file://' + html);
  await p.setInputFiles('#fileInput', csv);
  await p.waitForTimeout(800);
  // o painel mostra uma seção por vez: abre "Todas as seções" para ler tudo
  await p.evaluate(() => { const b = document.querySelector('#menuSecoes button[data-pagina="todas"]'); if (b) b.click(); });
  await p.waitForTimeout(500);
  // a seção de afastamentos abre com todos os tipos; a conferência compara os atestados
  await p.evaluate(() => { const s = document.getElementById('absMostrar'); if (s){ s.value = 'atestados'; s.dispatchEvent(new Event('change')); } });
  await p.waitForTimeout(300);
  // Dentro da página: lê o texto de cada card, os balões (data-tooltip) dos gráficos e as
  // linhas das tabelas, pelo id de cada elemento do index.html.
  const data = await p.evaluate(() => {
    const $ = id => document.getElementById(id);
    const txt = id => $(id) ? $(id).textContent.trim() : null;
    const tips = id => [...document.querySelectorAll('#'+id+' [data-tooltip]:not(.seg)')].map(e=>e.getAttribute('data-tooltip'));
    const table = id => [...document.querySelectorAll('#'+id+' tbody tr')].map(tr=>[...tr.children].map(td=>td.textContent.trim()));
    const kpis = [...document.querySelectorAll('#kpiRow .kpi-tile')].map(t=>[t.querySelector('.kpi-label').textContent.trim(), t.querySelector('.kpi-value').textContent.trim(), t.querySelector('.kpi-sub').textContent.trim()]);
    // a demora média saiu do card próprio (29/10) e virou a linha "Média de todos" da seção PontoNet
    kpis.push(['Demora média no PontoNet', txt('kpiDemora'), txt('kpiDemoraSub').replace(/^\(|\)$/g,'')]);
    return {
      kpis,
      tipo: [...document.querySelectorAll('#chartTipo .legend-row')].map(r=>[r.querySelector('.legend-label').textContent, r.querySelector('.legend-value').textContent]),
      tipoTotal: document.querySelector('#chartTipo .donut-total') && document.querySelector('#chartTipo .donut-total').textContent,
      diaSemana: tips('chartDiaSemana'), porData: tips('chartPorData'), gestor: tips('chartPorGestor'), cargo: tips('chartPorCargo'),
      top: tips('chartTopColaboradores'), banco: tips('chartBancoHoras'),
      extraFalta: table('tableExtraFalta'), curtas: table('tableCurtas'), pausas: table('tablePausasTermicas'),
      demoraG: table('tableDemoraGestor'), demoraC: table('tableDemoraColaborador'),
      histCount: txt('historyCount'), fileStatus: txt('fileStatus'),
      chipsTipo: [...document.querySelectorAll('#fTipo .ms-op')].map(c=>c.textContent),   // opções do filtro de tipo
      digKpis: $('digKpis') ? $('digKpis').innerText : null,
      digMotivo: [...document.querySelectorAll('#chartDigMotivo .legend-row')].map(r=>[r.querySelector('.legend-label').textContent, r.querySelector('.legend-value').textContent]),
      digTop: tips('chartDigTop'), digLinhas: document.querySelectorAll('#tableDigLista tbody tr').length,
      absEvol: table('tableAbsEvolucao'), absEmenda: txt('absEmendaResumo'), absRankN: document.querySelectorAll('#tableAbsRanking tbody tr').length,
      absReinc: [...document.querySelectorAll('#tableAbsRanking tr.row-alert')].map(tr=>tr.children[0].textContent.trim()),
      orgResumo: $('orgResumo') ? $('orgResumo').innerText : null,
      orgLinhas: document.querySelectorAll('#tableOrg tbody tr').length,
      grupoOpts: [...document.querySelectorAll('#fGrupo .ms-op input')].map(o=>o.value),
      // itens novos (out/2026): jornada, sinais de atestado, esquecimento crônico, pausas
      jorKpis: $('jorKpis') ? [...document.querySelectorAll('#jorKpis .colab-kpi')].map(k=>[k.querySelector('.ck-label').textContent, k.querySelector('.ck-value').textContent, k.querySelector('.ck-sub').textContent]) : null,
      jorDez: [...document.querySelectorAll('#chartJorDezHoras .chart-bar[data-tooltip]')].map(e=>e.getAttribute('data-tooltip')),
      jorGestor: table('tableJorGestor'),
      absMesmoDia: table('tableAbsMesmoDia'), absSegSex: txt('absSegSexResumo'), absSegSexPessoas: table('tableAbsSegSex'),
      absFerias: table('tableAbsFerias'),
      digEsq: table('tableDigEsquecimento'), digEsqAlerta: [...document.querySelectorAll('#tableDigEsquecimento tr.row-alert')].map(tr=>tr.children[0].textContent.trim()),
      tendUnidade: table('tableTendUnidade'),
      pauKpis: [...document.querySelectorAll('#pauKpis .colab-kpi')].map(k=>[k.querySelector('.ck-label').textContent, k.querySelector('.ck-value').textContent, k.querySelector('.ck-sub').textContent]),
      // PontoNet (27/10): pedidos e tempos por ciclo
      pnTempos: $('tablePnTempos') ? table('tablePnTempos') : null,
      // batidas de cada dia das pausas (B1, B2, B5, B7, B9, B12), do mês mais recente
      pauBat: $('pauHeat') ? {
        resumo: txt('pauBatResumo'),
        hora: [...document.querySelectorAll('#chartPauHora .chart-bar[data-tooltip]')].map(e=>e.getAttribute('data-tooltip')),
        dia: [...document.querySelectorAll('#chartPauDiaSemana .chart-bar[data-tooltip]')].map(e=>e.getAttribute('data-tooltip')),
        primeira: [...document.querySelectorAll('#chartPauPrimeira .chart-bar[data-tooltip]')].map(e=>e.getAttribute('data-tooltip')),
        semPausa: txt('pauSemPausaHint'), falta: txt('pauFaltaHint'),
        faltaTipo: [...document.querySelectorAll('#chartPauFaltaTipo .chart-bar[data-tooltip]')].map(e=>e.getAttribute('data-tooltip')),
        mapa: [...document.querySelectorAll('#pauHeat td.aus')].reduce((m, td)=>{ const c = td.style.background; m[c] = (m[c]||0) + 1; return m; }, {}),
      } : null,
    };
  });
  // ranking de mais de 10h e esquecimento crônico: lista inteira (só em tela cheia)
  for (const [id, chave, sel] of [['chartJorDezHoras','jorDez','.chart-bar[data-tooltip]'], ['tableDigEsquecimento','digEsqAlerta','tr.row-alert']]) {
    data[chave] = await p.evaluate(([id, sel]) => {
      const card = document.getElementById(id).closest('.card');
      card.querySelector('.btn-cheia').click();
      const r = [...document.querySelectorAll('#'+id+' '+sel)].map(e => sel.startsWith('tr') ? e.children[0].textContent.trim() : e.getAttribute('data-tooltip'));
      card.querySelector('.btn-cheia').click();
      return r;
    }, [id, sel]);
  }
  data.errors = errs;
  require('fs').writeFileSync(out, JSON.stringify(data, null, 1));
  await b.close();
})();
