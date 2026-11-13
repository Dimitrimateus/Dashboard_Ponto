// LEITOR DA TELA, PARTE 2 (13/11/2026): o que só aparece depois de um clique
// ---------------------------------------------------------------------------------
// Complemento de extrair_tudo_painel.js (que lê o painel como ele abre). Aqui o programa
// mexe no painel como um usuário e anota o que aparece:
//   cards    — o resumo de CADA pessoa do arquivo (a janela que abre ao clicar no nome);
//   gestores — o painel inteiro com o filtro de CADA gestor (um de cada vez);
//   estados  — o painel com cada opção de cada seletor de mês/tipo das seções;
//   origem   — as janelas "de onde vem" (clique numa barra, pedaço, fatia, ponto ou quadrado).
// Quem usa o JSON: ferramentas/auditar_extra.py, que refaz as contas e compara.
//
// Uso:
//   NODE_PATH=$(npm root -g) node extrair_extra_painel.js /caminho/index.html saida.json arquivo.csv [...]
const { chromium } = require('playwright');
const fs = require('fs');

(async () => {
  const [,, html, saida, ...csvs] = process.argv;
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1500, height: 1000 } });
  const erros = [];
  p.on('pageerror', e => erros.push(String(e)));
  p.on('dialog', d => { erros.push('janela: ' + d.message()); d.dismiss(); });
  await p.goto('file://' + html);
  await p.setInputFiles('#fileInput', csvs[0]);
  await p.waitForTimeout(1200);
  if (csvs.length > 1) { await p.setInputFiles('#fileInputAdd', csvs.slice(1)); await p.waitForTimeout(1200); }
  await p.evaluate(() => { document.getElementById('avisoAjudaOverlay').hidden = true; });
  await p.evaluate(() => document.querySelector('#menuSecoes button[data-pagina="todas"]').click());
  await p.waitForTimeout(600);

  // Lê, de cada elemento com id dentro do painel, números, balões, legendas e tabelas.
  const ler = (somenteCard) => p.evaluate((somenteCard) => {
    const txt = el => (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
    const out = {};
    const raiz = somenteCard ? document.querySelector('.card-cheia') : document;
    if (!raiz) return out;
    raiz.querySelectorAll('#dashboardBody [id]').forEach(el => {
      if (/^(btn|f[A-Z]|chartTooltip)/.test(el.id)) return;
      const o = {};
      const tips = [...el.querySelectorAll('[data-tooltip]:not(.seg)')].filter(x => !x.closest('svg.mini-barras') && !x.closest('table.heat')).map(x => x.getAttribute('data-tooltip'));
      if (tips.length) o.tips = tips;
      const segs = [...el.querySelectorAll('.seg[data-tooltip]')].map(x => x.getAttribute('data-tooltip'));
      if (segs.length) o.segs = segs;
      const leg = [...el.querySelectorAll('.legend-row')].map(r => [txt(r.querySelector('.legend-label') || r), txt(r.querySelector('.legend-value') || r)]);
      if (leg.length) o.legenda = leg;
      const legm = [...el.querySelectorAll('.legenda-mini > *')].map(x => txt(x)).filter(Boolean);
      if (legm.length) o.legendaMini = legm;
      const tabs = [...el.querySelectorAll('table')].map(t => {
        const r = { head: [...t.querySelectorAll('thead th')].map(txt), rows: [...t.querySelectorAll('tbody tr')].map(tr => [...tr.children].map(td => txt(td))) };
        if (t.classList.contains('heat')) r.cores = [...t.querySelectorAll('tbody tr')].map(tr => { const c = {}; tr.querySelectorAll('td[style*="background"]').forEach(td => { const k = td.style.background; c[k] = (c[k] || 0) + 1; }); return c; });
        return r;
      });
      if (tabs.length) o.tabelas = tabs;
      const tiles = [...el.querySelectorAll('.colab-kpi, .kpi-tile')].map(t => { const g = s => { const x = t.querySelector(s); return x ? txt(x) : ''; }; return [g('.ck-label') || g('.kpi-label'), g('.ck-value') || g('.kpi-value'), g('.ck-sub') || g('.kpi-sub')]; });
      if (tiles.length) o.numeros = tiles;
      o.texto = txt(el).slice(0, 6000);
      o.escondido = !!el.closest('[hidden]') || el.offsetParent === null;
      out[el.id] = o;
    });
    return out;
  }, somenteCard);
  // Lê o painel inteiro + as listas longas em tela cheia.
  const lerComTelaCheia = async () => {
    const normal = await ler(false), telaCheia = {};
    const n = await p.evaluate(() => {
      document.querySelectorAll('[data-aud-card]').forEach(c => c.removeAttribute('data-aud-card'));
      const cards = [...document.querySelectorAll('.pagina .card')].filter(c => [...c.querySelectorAll('.btn-todos, #btnBancoTodos')].some(x => !x.hidden));
      cards.forEach((c, i) => c.setAttribute('data-aud-card', i));
      return cards.length;
    });
    for (let i = 0; i < n; i++) {
      await p.evaluate(i => { const c = document.querySelector('[data-aud-card="' + i + '"]'); const bt = c && c.querySelector('.btn-todos:not([hidden]), #btnBancoTodos:not([hidden])'); if (bt) bt.click(); }, i);
      await p.waitForTimeout(150);
      Object.assign(telaCheia, await ler(true));
      await p.keyboard.press('Escape');
      await p.waitForTimeout(80);
    }
    return { normal, telaCheia };
  };
  const estado = () => p.evaluate(() => {
    const v = id => (document.getElementById(id) || {}).value;
    const t = id => { const s = document.getElementById(id); return s && s.selectedIndex > -1 ? s.options[s.selectedIndex].text : ''; };
    return { periodoIni: v('fDataIni'), periodoFim: v('fDataFim'), digCiclo: t('digCiclo'), absCiclo: t('absCiclo'), absMostrar: t('absMostrar'),
             jorCiclo: t('jorCiclo'), pauBatCiclo: t('pauBatCiclo'), pnCiclo: t('pnCiclo'), tendIndicador: t('tendIndicador'),
             bhMeses: [...document.querySelectorAll('#bhMeses input')].map(i => [i.parentElement.textContent.trim(), i.checked]),
             gestor: [...document.querySelectorAll('#fGestor input:checked')].map(i => i.value),
             recorte: [...document.querySelectorAll('.recorte-secao')].map(x => x.textContent.trim())[0] || '' };
  });
  const dados = { cards: {}, gestores: [], estados: [], origem: [], erros };

  // PARTES=cards,gestores,estados,origem escolhe o que ler (padrão: tudo)
  const partes = (process.env.PARTES || 'cards,gestores,estados,origem').split(',');
  // 1) o resumo de cada pessoa
  const nomes = await p.evaluate(() => {
    const s = new Set();
    document.querySelectorAll('[data-colab]').forEach(e => s.add(e.getAttribute('data-colab')));
    return [...s];
  });
  // todos os nomes do arquivo: a lista de colaboradores tem os do cadastro; os de fora entram pelos gráficos/tabelas
  for (const nome of (partes.includes('cards') ? nomes : [])) {
    const c = await p.evaluate(nome => {
      const a = document.createElement('span'); a.setAttribute('data-colab', nome); document.body.appendChild(a); a.click(); a.remove();
      const txt = el => (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
      const ov = document.getElementById('colabModalOverlay');
      if (ov.hidden) return null;
      const corpo = document.getElementById('colabModalBody');
      const r = { meta: txt(document.getElementById('colabModalMeta')),
        tiles: [...corpo.querySelectorAll('.colab-kpi')].map(t => [txt(t.querySelector('.ck-label')), txt(t.querySelector('.ck-value')), t.querySelector('.ck-sub') ? txt(t.querySelector('.ck-sub')) : '']),
        secoes: [...corpo.querySelectorAll('h3.colab-section')].map(h => {
          let el = h.nextElementSibling; const tb = el && el.querySelector ? el.querySelector('table') : null;
          return { titulo: txt(h), head: tb ? [...tb.querySelectorAll('thead th')].map(txt) : [], rows: tb ? [...tb.querySelectorAll('tbody tr')].map(tr => [...tr.children].map(td => txt(td))) : [] };
        }),
        tags: [...corpo.querySelectorAll('.colab-tag')].map(txt) };
      document.getElementById('colabModalClose').click();
      return r;
    }, nome);
    if (c) dados.cards[nome] = c;
  }

  // 2) cada gestor no filtro
  const gestores = partes.includes('gestores') ? await p.evaluate(() => [...document.querySelectorAll('#fGestor .ms-op input')].map(i => i.value)) : [];
  for (const g of gestores) {
    await p.evaluate(g => { const i = [...document.querySelectorAll('#fGestor .ms-op input')].find(x => x.value === g); i.click(); }, g);
    await p.waitForTimeout(350);
    const lido = await lerComTelaCheia();
    dados.gestores.push({ gestor: g, estado: await estado(), normal: lido.normal, telaCheia: lido.telaCheia });
    console.error('gestor', g);
    await p.evaluate(g => { const i = [...document.querySelectorAll('#fGestor .ms-op input')].find(x => x.value === g); i.click(); }, g);
    await p.waitForTimeout(250);
  }

  // 3) cada opção dos seletores (os outros no padrão)
  const seletores = ['digCiclo', 'jorCiclo', 'pnCiclo', 'pauBatCiclo', 'absCiclo', 'absMostrar', 'tendIndicador'];
  for (const id of (partes.includes('estados') ? seletores : [])) {
    const ops = await p.evaluate(id => [...document.getElementById(id).options].map(o => o.value), id);
    const padrao = await p.evaluate(id => document.getElementById(id).value, id);
    for (const v of ops) {
      if (v === padrao) continue;
      await p.evaluate(([id, v]) => { const s = document.getElementById(id); s.value = v; s.dispatchEvent(new Event('change', { bubbles: true })); }, [id, v]);
      await p.waitForTimeout(400);
      const lido = await lerComTelaCheia();
      dados.estados.push({ seletor: id, valor: v, estado: await estado(), normal: lido.normal, telaCheia: lido.telaCheia });
    }
    await p.evaluate(([id, v]) => { const s = document.getElementById(id); s.value = v; s.dispatchEvent(new Event('change', { bubbles: true })); }, [id, padrao]);
    await p.waitForTimeout(300);
  }
  // meses do banco de horas: cada mês sozinho
  const meses = await p.evaluate(() => [...document.querySelectorAll('#bhMeses input')].map(i => i.getAttribute('data-bh-mes')));
  for (const m of (partes.includes('estados') ? meses : [])) {
    // desmarca todos e marca só este (o painel mantém os meses seguidos)
    await p.evaluate(m => {
      const ins = [...document.querySelectorAll('#bhMeses input')];
      ins.forEach(i => { if (i.checked && i.getAttribute('data-bh-mes') !== m) i.click(); });
      const alvo = document.querySelector('#bhMeses input[data-bh-mes="' + m + '"]'); if (!alvo.checked) alvo.click();
      [...document.querySelectorAll('#bhMeses input')].forEach(i => { if (i.checked && i.getAttribute('data-bh-mes') !== m) i.click(); });
    }, m);
    await p.waitForTimeout(400);
    const lido = await lerComTelaCheia();
    dados.estados.push({ seletor: 'bhMeses', valor: m, estado: await estado(), normal: lido.normal, telaCheia: lido.telaCheia });
  }

  // 4) janelas "de onde vem": cada barra/pedaço/fatia/ponto/quadrado clicável (até 40 por gráfico)
  await p.goto('file://' + html);
  await p.setInputFiles('#fileInput', csvs[0]);
  await p.waitForTimeout(1200);
  if (csvs.length > 1) { await p.setInputFiles('#fileInputAdd', csvs.slice(1)); await p.waitForTimeout(1200); }
  await p.evaluate(() => { document.getElementById('avisoAjudaOverlay').hidden = true; });
  await p.evaluate(() => document.querySelector('#menuSecoes button[data-pagina="todas"]').click());
  await p.waitForTimeout(600);
  const alvos = await p.evaluate(() => {
    const out = [];
    document.querySelectorAll('#dashboardBody [id]').forEach(cont => {
      const els = [...cont.querySelectorAll(':scope [data-drill]')].filter(e => e.closest('[id]') === cont || !e.closest('[id]') || e.closest('[id]').id === cont.id);
      els.slice(0, 40).forEach((e, i) => { e.setAttribute('data-aud-drill', cont.id + '#' + i); out.push(cont.id + '#' + i); });
    });
    return out;
  });
  for (const a of (partes.includes('origem') ? alvos : [])) {
    const r = await p.evaluate(a => {
      const e = document.querySelector('[data-aud-drill="' + a + '"]');
      if (!e) return null;
      const filho = e.querySelector ? e.querySelector('[data-tooltip]') : null;
      const dica = e.getAttribute('data-tooltip') || (filho ? filho.getAttribute('data-tooltip') : '') || '';
      e.dispatchEvent(new MouseEvent('click', { bubbles: true }));
      const ov = document.getElementById('origemOverlay');
      if (ov.hidden) return { alvo: a, dica, aberto: false };
      const txt = el => (el.innerText || el.textContent || '').replace(/\s+/g, ' ').trim();
      const o = { alvo: a, dica, aberto: true, titulo: txt(document.getElementById('origemTitulo')), sub: txt(document.getElementById('origemSub')),
        linhas: ov.querySelectorAll('.origem-body table tbody tr').length,
        head: [...ov.querySelectorAll('.origem-body table thead th')].map(txt) };
      document.getElementById('origemFechar').click();
      return o;
    }, a);
    if (r) dados.origem.push(r);
  }
  fs.writeFileSync(saida, JSON.stringify(dados));
  console.log('ok:', Object.keys(dados.cards).length, 'pessoas;', dados.gestores.length, 'gestores;', dados.estados.length, 'estados;', dados.origem.length, 'cliques "de onde vem"; erros de JS:', erros.length);
  await b.close();
})();
