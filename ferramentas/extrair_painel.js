// Abre o index.html no Chromium (Playwright), carrega um CSV e grava em JSON o que cada
// card/gráfico/tabela mostra. Uso:
//   NODE_PATH=$(npm root -g) node extrair_painel.js /caminho/index.html /caminho/dados.csv painel.json
const { chromium } = require('playwright');
(async () => {
  const [,, html, csv, out] = process.argv;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1400, height: 1000 } });
  const errs=[]; p.on('pageerror', e=>errs.push(String(e))); p.on('dialog', d=>{errs.push('dialog:'+d.message()); d.dismiss();});
  await p.goto('file://' + html);
  await p.setInputFiles('#fileInput', csv);
  await p.waitForTimeout(800);
  const data = await p.evaluate(() => {
    const $ = id => document.getElementById(id);
    const txt = id => $(id) ? $(id).textContent.trim() : null;
    const tips = id => [...document.querySelectorAll('#'+id+' [data-tooltip]')].map(e=>e.getAttribute('data-tooltip'));
    const table = id => [...document.querySelectorAll('#'+id+' tbody tr')].map(tr=>[...tr.children].map(td=>td.textContent.trim()));
    const kpis = [...document.querySelectorAll('#kpiRow .kpi-tile')].map(t=>[t.querySelector('.kpi-label').textContent.trim(), t.querySelector('.kpi-value').textContent.trim(), t.querySelector('.kpi-sub').textContent.trim()]);
    return {
      kpis,
      tipo: [...document.querySelectorAll('#chartTipo .legend-row')].map(r=>[r.querySelector('.legend-label').textContent, r.querySelector('.legend-value').textContent]),
      tipoTotal: document.querySelector('#chartTipo .donut-total') && document.querySelector('#chartTipo .donut-total').textContent,
      diaSemana: tips('chartDiaSemana'), porData: tips('chartPorData'), gestor: tips('chartPorGestor'), cargo: tips('chartPorCargo'),
      top: tips('chartTopColaboradores'), banco: tips('chartBancoHoras'),
      extraFalta: table('tableExtraFalta'), curtas: table('tableCurtas'), pausas: table('tablePausasTermicas'),
      demoraG: table('tableDemoraGestor'), demoraC: table('tableDemoraColaborador'),
      histCount: txt('historyCount'), fileStatus: txt('fileStatus'),
      chipsTipo: [...document.querySelectorAll('#fTipo .chip')].map(c=>c.textContent),
      digKpis: $('digKpis') ? $('digKpis').innerText : null,
      digMotivo: [...document.querySelectorAll('#chartDigMotivo .legend-row')].map(r=>[r.querySelector('.legend-label').textContent, r.querySelector('.legend-value').textContent]),
      digTop: tips('chartDigTop'), digLinhas: document.querySelectorAll('#tableDigLista tbody tr').length,
      absEvol: table('tableAbsEvolucao'), absEmenda: txt('absEmendaResumo'), absRankN: document.querySelectorAll('#tableAbsRanking tbody tr').length,
      absReinc: [...document.querySelectorAll('#tableAbsRanking tr.row-alert')].map(tr=>tr.children[0].textContent.trim()),
      orgResumo: $('orgResumo') ? $('orgResumo').innerText : null,
      orgFora: [...document.querySelectorAll('#orgForaBox li')].map(li=>li.querySelector('.colab-link').textContent),
      orgLinhas: document.querySelectorAll('#tableOrg tbody tr').length,
      grupoOpts: $('fGrupo') ? [...$('fGrupo').options].map(o=>o.value) : null,
    };
  });
  data.errors = errs;
  require('fs').writeFileSync(out, JSON.stringify(data, null, 1));
  await b.close();
})();
