// LEITOR DA TELA DO PAINEL (para a conferência completa, 11/11/2026)
// ---------------------------------------------------------------------------------
// O que faz: abre o index.html no Chromium (Playwright), carrega um ou mais CSVs e
// grava num JSON TUDO o que a tela mostra, card por card: os números do topo de cada
// seção, os balões dos gráficos ("Nome: 12"), as legendas, as tabelas (cabeçalho e
// linhas) e os textos de resumo. Listas que no card mostram só os primeiros nomes são
// lidas em TELA CHEIA, onde aparecem inteiras.
// Quem usa o JSON: ferramentas/auditar_csv_painel.py, que refaz cada conta direto do CSV
// (sem usar o código do painel) e compara com o que foi lido aqui.
//
// Uso:
//   NODE_PATH=$(npm root -g) node extrair_tudo_painel.js /caminho/index.html saida.json arquivo1.csv [arquivo2.csv ...]
//
// O painel é lido no estado em que abre (sem filtros, seletores no padrão de cada
// seção). Erros de JavaScript da página vão para "erros" no JSON: tem que sair vazio.
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const [,, html, saida, ...csvs] = process.argv;
  if (!html || !saida || !csvs.length) { console.error('uso: node extrair_tudo_painel.js index.html saida.json arquivo.csv [...]'); process.exit(2); }
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1500, height: 1000 } });
  const erros = [];
  p.on('pageerror', e => erros.push(String(e)));
  p.on('dialog', d => { erros.push('janela: ' + d.message()); d.dismiss(); });
  await p.goto('file://' + html);
  await p.setInputFiles('#fileInput', csvs[0]);
  await p.waitForTimeout(1200);
  if (csvs.length > 1) { await p.setInputFiles('#fileInputAdd', csvs.slice(1)); await p.waitForTimeout(1200); }
  // (11/11) fecha o aviso do botão "?" que aparece ao abrir o arquivo
  await p.evaluate(() => { const o = document.getElementById('avisoAjudaOverlay'); if (o) o.hidden = true; });
  // "Todas as seções numa página só": tudo desenhado com largura real
  await p.evaluate(() => document.querySelector('#menuSecoes button[data-pagina="todas"]').click());
  await p.waitForTimeout(800);

  // Lê o conteúdo de um elemento (por id) num formato simples.
  const lerTudo = async (somenteCard) => p.evaluate((somenteCard) => {
    const txt = el => (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
    function info(el) {
      const o = {};
      const tips = [...el.querySelectorAll('[data-tooltip]:not(.seg)')]
        .filter(x => !x.closest('svg.mini-barras') && !x.closest('table.heat'))
        .map(x => x.getAttribute('data-tooltip'));
      if (tips.length) o.tips = tips;
      const segs = [...el.querySelectorAll('.seg[data-tooltip]')].map(x => x.getAttribute('data-tooltip'));
      if (segs.length) o.segs = segs;
      const leg = [...el.querySelectorAll('.legend-row')].map(r => [txt(r.querySelector('.legend-label') || r), txt(r.querySelector('.legend-value') || r)]);
      if (leg.length) o.legenda = leg;
      const legm = [...el.querySelectorAll('.legenda-mini > *')].map(x => txt(x)).filter(Boolean);
      if (legm.length) o.legendaMini = legm;
      const tabs = [...el.querySelectorAll('table')].map(t => {
        const head = [...t.querySelectorAll('thead th')].map(txt);
        const rows = [...t.querySelectorAll('tbody tr')].map(tr => [...tr.children].map(td => txt(td)));
        const r = { head, rows };
        if (t.classList.contains('heat')) {
          // mapas: em cada linha, quantos quadrados de cada cor
          r.cores = [...t.querySelectorAll('tbody tr')].map(tr => {
            const c = {};
            tr.querySelectorAll('td[style*="background"]').forEach(td => { const k = td.style.background; c[k] = (c[k] || 0) + 1; });
            return c;
          });
          r.balaoQuadrados = [...t.querySelectorAll('td[data-tooltip]')].map(td => td.getAttribute('data-tooltip'));
        }
        return r;
      });
      if (tabs.length) o.tabelas = tabs;
      const tiles = [...el.querySelectorAll('.colab-kpi, .kpi-tile')].map(t => {
        const g = s => { const x = t.querySelector(s); return x ? txt(x) : ''; };
        return [g('.ck-label') || g('.kpi-label'), g('.ck-value') || g('.kpi-value'), g('.ck-sub') || g('.kpi-sub')];
      });
      if (tiles.length) o.numeros = tiles;
      o.texto = txt(el).slice(0, 6000);
      o.escondido = !!el.closest('[hidden]') || el.offsetParent === null;
      return o;
    }
    const out = {};
    const raiz = somenteCard ? document.querySelector('.card-cheia') : document;
    if (!raiz) return out;
    raiz.querySelectorAll('#dashboardBody [id]').forEach(el => {
      if (somenteCard && !raiz.contains(el)) return;
      if (/^(btn|f[A-Z]|chartTooltip)/.test(el.id)) return;
      out[el.id] = info(el);
    });
    return out;
  }, somenteCard);

  const dados = { normal: await lerTudo(false), telaCheia: {} };

  // Cards com lista cortada ("Ver todos (N) em tela cheia"): abre cada um em tela
  // cheia e lê de novo, para pegar a lista inteira.
  const nCards = await p.evaluate(() => {
    const cards = [...document.querySelectorAll('.pagina .card')].filter(c =>
      [...c.querySelectorAll('.btn-todos, #btnBancoTodos')].some(x => !x.hidden));
    cards.forEach((c, i) => c.setAttribute('data-aud-card', i));
    return cards.length;
  });
  for (let i = 0; i < nCards; i++) {
    await p.evaluate(i => { const c = document.querySelector('[data-aud-card="' + i + '"]'); const bt = c.querySelector('.btn-todos:not([hidden]), #btnBancoTodos:not([hidden])'); if (bt) bt.click(); }, i);
    await p.waitForTimeout(250);
    const lido = await lerTudo(true);
    Object.assign(dados.telaCheia, lido);
    await p.keyboard.press('Escape');
    await p.waitForTimeout(150);
  }

  // Período do filtro em que o painel abriu e os seletores padrão de cada seção
  dados.estado = await p.evaluate(() => {
    const v = id => (document.getElementById(id) || {}).value;
    const t = id => { const s = document.getElementById(id); return s && s.selectedIndex > -1 ? s.options[s.selectedIndex].text : ''; };
    return { periodoIni: v('fDataIni'), periodoFim: v('fDataFim'), digCiclo: t('digCiclo'), absCiclo: t('absCiclo'), absMostrar: t('absMostrar'),
             jorCiclo: t('jorCiclo'), pauBatCiclo: t('pauBatCiclo'), pnCiclo: t('pnCiclo'), tendIndicador: t('tendIndicador'),
             bhMeses: [...document.querySelectorAll('#bhMeses input')].map(i => [i.parentElement.textContent.trim(), i.checked]) };
  });
  // Etiquetas 📅 de cada quadro (o período que ele considera), pelo id do gráfico/tabela de dentro
  dados.selos = await p.evaluate(() => {
    const out = {};
    document.querySelectorAll('.card-periodo').forEach(s => {
      const card = s.closest('.card'); if (!card || s.hidden || !s.textContent.trim()) return;
      card.querySelectorAll('[id]').forEach(el => { if (/^(chart|table)/.test(el.id) || el.id === 'absHeat' || el.id === 'pauHeat' || el.id === 'orgSection') out[el.id] = s.textContent.trim(); });
      if (card.id) out[card.id] = s.textContent.trim();
    });
    document.querySelectorAll('.kpi-tile .kpi-periodo, .colab-kpi .ck-periodo').forEach(p => {
      const sec = p.closest('[id$="Kpis"]'); if (sec) (out[sec.id] = out[sec.id] || []).push(p.textContent.trim());
    });
    return out;
  });
  dados.erros = erros;
  fs.writeFileSync(saida, JSON.stringify(dados));
  console.log('ok:', Object.keys(dados.normal).length, 'elementos lidos;', nCards, 'listas lidas em tela cheia; erros de JS:', erros.length);
  await b.close();
})();
