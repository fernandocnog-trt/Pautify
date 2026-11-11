/* =========================================================
   Pauta da Sessão — lógica da página
   Etapa 1: importa .docx/.xlsx → interpreta colunas → painel de confirmação
   Etapa 2: layout otimizado com "Decisão e observações",
            salvamento automático (localStorage), impressão A4 e exportação .xlsx
   ========================================================= */
(function () {
  'use strict';

  const STORAGE_KEY = 'pautaSessao.v1';
  const LABELS_KEY = 'pautaSessao.relatores.v1';
  const DEFAULT_OBS = 'Sem divergência lançada até às 15h00.';
  const LIBS = {
    xlsx: 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js',
    mammoth: 'https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js'
  };
  const EXPORT_HEADERS = ['ORDEM', 'PAUTA', 'RELATOR', 'PROCESSO', 'ADVOGADO', 'MODALIDADE', 'PARTE', 'DECISÃO E OBSERVAÇÕES', 'OBS.'];
  const MOD_NAMES = { V: 'Virtual', P: 'Presencial' };
  const COLUMN_NAMES = { ordem: 'Ordem', pauta: 'Pauta', relator: 'Relator', processo: 'Processo', advogado: 'Advogado', modalidade: 'Modalidade', parte: 'Parte' };

  /* ---------------------------------------------------------
     Funções puras (interpretação da tabela)
     --------------------------------------------------------- */
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase().replace(/[^A-Z0-9]/g, '');
  const clean = s => String(s ?? '').replace(/ /g, ' ').replace(/[ \t]+/g, ' ').trim();
  const splitLines = s => String(s ?? '').split(/\r?\n/).map(clean).filter(Boolean);
  const oneLine = s => splitLines(s).join(' ');

  function headerKey(h) {
    const n = norm(h);
    if (!n) return null;
    if (n.startsWith('DESFECHO') || n.startsWith('DECISAO')) return 'desfecho';
    if (n.startsWith('OBS')) return 'obs';
    if (n.startsWith('ORDEM')) return 'ordem';
    if (n.startsWith('PAUTA')) return 'pauta';
    if (n.startsWith('RELATOR')) return 'relator';
    if (n.startsWith('PROCESSO')) return 'processo';
    if (n.startsWith('ADVOGAD')) return 'advogado';
    if (n.startsWith('MODALIDADE')) return 'modalidade';
    if (n.startsWith('PARTE')) return 'parte';
    return null;
  }

  function findHeader(rows) {
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const map = {};
      (rows[i] || []).forEach((cell, j) => {
        const k = headerKey(cell);
        if (k && !(k in map)) map[k] = j;
      });
      if ('processo' in map && Object.keys(map).length >= 3) return { index: i, map };
    }
    return null;
  }

  function fmtOrdem(v) {
    const s = oneLine(v);
    const m = s.match(/^(\d+)\s*[°ºo]?\.?$/i);
    return m ? m[1] + 'º' : s;
  }

  function fmtPauta(v) {
    const s = oneLine(v);
    return /^\d$/.test(s) ? '0' + s : s;
  }

  /* Pareia, de cima para baixo, cada advogado com a sua parte (1 linha = 1 sustentação) */
  function pairUp(partes, advogados) {
    const n = Math.max(partes.length, advogados.length);
    const out = [];
    for (let i = 0; i < n; i++) out.push({ parte: partes[i] || '', advogado: advogados[i] || '', modalidade: '' });
    return out;
  }

  /* "v", "Virtual" → "V" · "p", "Presencial" → "P" · outros valores são mantidos */
  function normMod(v) {
    const n = norm(v);
    if (!n) return '';
    if (n === 'V' || n.startsWith('VIRT')) return 'V';
    if (n === 'P' || n.startsWith('PRES')) return 'P';
    return clean(v).toUpperCase();
  }

  /* Célula de modalidade → lista na ordem vertical ("V\nV\nP" ou "V V P" → [V, V, P]) */
  function splitMods(s) {
    return splitLines(s)
      .flatMap(l => (/^[vp]([\s,;\/]+[vp])+$/i.test(l) ? l.split(/[\s,;\/]+/) : [l]))
      .map(normMod)
      .filter(Boolean);
  }

  /* Regra da pauta: uma única marcação vale para todas as sustentações do processo;
     havendo mais de uma, segue estritamente a ordem vertical dos advogados.
     Retorna false quando a quantidade não fecha (para gerar aviso). */
  function assignMods(linhas, mods) {
    if (!linhas.length && mods.length) linhas.push({ parte: '', advogado: '', modalidade: '' });
    linhas.forEach((l, i) => { l.modalidade = (mods.length === 1 ? mods[0] : mods[i]) || ''; });
    return mods.length === 1 || mods.length === linhas.length;
  }

  /* 0 = fonte normal … 3 = mais compacta (muitas sustentações ou linhas longas) */
  function partyDensity(partes) {
    const n = partes.length;
    const longest = Math.max(0, ...partes.map(p => (p.advogado || '').length + (p.parte || '').length));
    if (n >= 7) return 3;
    if (n >= 5 || (n >= 3 && longest > 70)) return 2;
    if (n >= 3 || longest > 70) return 1;
    return 0;
  }

  /* Converte sessões salvas no formato antigo (modalidade única por processo) */
  function migrate(session) {
    (session.items || []).forEach(it => {
      it.partes = it.partes || [];
      if ('modalidade' in it) {
        assignMods(it.partes, splitMods(it.modalidade));
        delete it.modalidade;
      }
      it.partes.forEach(p => { if (p.modalidade == null) p.modalidade = ''; });
    });
    return session;
  }

  function rowsToSession(rows, fallbackTitle) {
    const header = findHeader(rows);
    if (!header) {
      throw new Error('Não encontrei o cabeçalho da tabela (ORDEM, PAUTA, RELATOR, PROCESSO, ADVOGADO, MODALIDADE, PARTE). Verifique se o arquivo contém a tabela da pauta.');
    }
    const { index, map } = header;

    let title = '';
    for (let i = 0; i < index && !title; i++) {
      const t = (rows[i] || []).map(oneLine).filter(Boolean).join(' ');
      if (t) title = t;
    }

    const get = (r, k) => (k in map ? String(r[map[k]] ?? '') : '');
    const items = [];
    const warnings = [];

    for (let i = index + 1; i < rows.length; i++) {
      const r = rows[i] || [];
      const processo = oneLine(get(r, 'processo'));
      const ordem = fmtOrdem(get(r, 'ordem'));
      const pairs = pairUp(splitLines(get(r, 'parte')), splitLines(get(r, 'advogado')));
      const mods = splitMods(get(r, 'modalidade'));

      // Linha de continuação (células mescladas): só partes/advogados → anexa ao processo anterior
      if (!processo && !ordem) {
        if (items.length) {
          items[items.length - 1].partes.push(...pairs);
          items[items.length - 1]._mods.push(...mods);
        }
        continue;
      }

      const item = {
        id: 'i' + items.length + '-' + Date.now().toString(36),
        ordem,
        pauta: fmtPauta(get(r, 'pauta')),
        relator: oneLine(get(r, 'relator')),
        processo,
        partes: pairs,
        _mods: mods,
        desfecho: String(get(r, 'desfecho')).trim(),
        obs: oneLine(get(r, 'obs'))
      };
      if (!processo) warnings.push(`Ordem ${ordem}: sem número de processo.`);
      items.push(item);
    }

    items.forEach(it => {
      const ref = `Ordem ${it.ordem || '?'} (${it.processo})`;
      const mods = it._mods;
      delete it._mods;
      const ok = assignMods(it.partes, mods);
      if (mods.length && !ok) {
        warnings.push(`${ref}: ${mods.length} modalidade(s) para ${it.partes.length} sustentação(ões) — atribuídas de cima para baixo; confira.`);
      } else if (!mods.length && 'modalidade' in map && it.partes.length) {
        warnings.push(`${ref}: modalidade não informada.`);
      }
      const nP = it.partes.filter(p => p.parte).length;
      const nA = it.partes.filter(p => p.advogado).length;
      if (nP && nA && nP !== nA) {
        warnings.push(`${ref}: ${nP} parte(s) e ${nA} advogado(s) — pareados pela ordem das linhas.`);
      }
    });

    if (!items.length) throw new Error('A tabela foi encontrada, mas não há processos abaixo do cabeçalho.');

    return {
      version: 1,
      title: title || fallbackTitle || 'Sessão',
      items,
      warnings,
      columns: Object.keys(map)
    };
  }

  /* Preserva decisões já digitadas ao reimportar a MESMA sessão (mesmo título) */
  function mergePrevious(next, prev) {
    if (!prev || !prev.items || norm(prev.title) !== norm(next.title)) return 0;
    const byProc = new Map(prev.items.map(it => [norm(it.processo), it]));
    let kept = 0;
    next.items.forEach(it => {
      const old = byProc.get(norm(it.processo));
      if (!old) return;
      let touched = false;
      if (!it.desfecho && old.desfecho) { it.desfecho = old.desfecho; touched = true; }
      if (!it.obs && old.obs) { it.obs = old.obs; touched = true; }
      if (touched) kept++;
    });
    return kept;
  }

  const api = { norm, headerKey, findHeader, fmtOrdem, fmtPauta, pairUp, normMod, splitMods, assignMods, partyDensity, migrate, rowsToSession, mergePrevious };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  if (typeof document === 'undefined') return; // execução em Node (testes)

  /* ---------------------------------------------------------
     Leitura de arquivos
     --------------------------------------------------------- */
  const loaded = {};
  function loadScript(key) {
    if (!loaded[key]) {
      loaded[key] = new Promise((resolve, reject) => {
        const s = document.createElement('script');
        s.src = LIBS[key];
        s.onload = resolve;
        s.onerror = () => { delete loaded[key]; reject(new Error('Não foi possível carregar a biblioteca de leitura. Verifique a conexão com a internet.')); };
        document.head.appendChild(s);
      });
    }
    return loaded[key];
  }

  function cellText(td) {
    const ps = td.querySelectorAll('p, li');
    if (ps.length) return Array.from(ps, p => p.textContent).join('\n');
    return td.textContent;
  }

  function tableToRows(table) {
    const grid = [];
    Array.from(table.rows).forEach((tr, ri) => {
      grid[ri] = grid[ri] || [];
      let ci = 0;
      Array.from(tr.cells).forEach(td => {
        while (grid[ri][ci] !== undefined) ci++;
        const text = cellText(td);
        const rs = Math.max(1, td.rowSpan || 1);
        const cs = Math.max(1, td.colSpan || 1);
        for (let r = 0; r < rs; r++) {
          grid[ri + r] = grid[ri + r] || [];
          for (let c = 0; c < cs; c++) grid[ri + r][ci + c] = (r === 0 && c === 0) ? text : '';
        }
        ci += cs;
      });
    });
    return grid.map(row => Array.from(row, v => v ?? ''));
  }

  async function readDocx(buf) {
    await loadScript('mammoth');
    const { value } = await window.mammoth.convertToHtml({ arrayBuffer: buf });
    const doc = new DOMParser().parseFromString(value, 'text/html');
    const tables = Array.from(doc.querySelectorAll('table'));
    for (const table of tables) {
      const rows = tableToRows(table);
      if (!findHeader(rows)) continue;
      // título: parágrafo antes da tabela que mencione "Sessão" (ou o 1º não vazio)
      const before = [];
      for (const el of doc.body.children) {
        if (el === table || el.contains(table)) break;
        const t = el.textContent.trim();
        if (t) before.push(t);
      }
      const title = before.find(t => /sess[aã]o/i.test(t)) || before[0] || '';
      return title ? [[title], ...rows] : rows;
    }
    throw new Error('Não encontrei, no documento Word, uma tabela com as colunas da pauta.');
  }

  async function readSheet(buf) {
    await loadScript('xlsx');
    const wb = window.XLSX.read(buf, { type: 'array' });
    for (const name of wb.SheetNames) {
      const rows = window.XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: '', blankrows: false });
      if (findHeader(rows)) return rows;
    }
    throw new Error('Não encontrei, na planilha, uma aba com as colunas da pauta.');
  }

  /* ---------------------------------------------------------
     Estado e persistência
     --------------------------------------------------------- */
  const $ = (sel, el = document) => el.querySelector(sel);
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  function loadJSON(key) {
    try { return JSON.parse(localStorage.getItem(key)); } catch { return null; }
  }

  let state = null;     // sessão em edição (persistida)
  let pending = null;   // sessão recém-importada aguardando confirmação
  let relatorLabels = loadJSON(LABELS_KEY) || {};
  let saveTimer = null;

  const els = {
    fileInput: $('#fileInput'),
    dropzone: $('#dropzone'),
    panel: $('#confirmPanel'),
    stageUpload: $('#stageUpload'),
    stageEditor: $('#stageEditor'),
    title: $('#sessionTitle'),
    cards: $('#cards'),
    status: $('#saveStatus'),
    toast: $('#toast'),
    btnImport: $('#btnImport'),
    btnPrint: $('#btnPrint'),
    btnExport: $('#btnExport'),
    btnReset: $('#btnReset'),
    step1: $('#step1'),
    step2: $('#step2')
  };

  function hhmm(d = new Date()) {
    return d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
  }

  function progressHTML() {
    if (!state) return '';
    const total = state.items.length;
    const done = state.items.filter(it => it.desfecho.trim()).length;
    return `<span class="progress${done === total ? ' is-done' : ''}" title="Processos com desfecho preenchido">${done}/${total}</span>`;
  }

  function setStatus(kind, text) {
    if (!state) { els.status.innerHTML = ''; return; }
    const cls = kind === 'saving' ? ' is-saving' : kind === 'error' ? ' is-error' : '';
    els.status.innerHTML = `${progressHTML()}<span class="dot${cls}"></span><span>${esc(text)}</span>`;
  }

  function saveNow() {
    clearTimeout(saveTimer);
    saveTimer = null;
    if (!state) return;
    try {
      state.updatedAt = new Date().toISOString();
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
      setStatus('ok', 'Salvo às ' + hhmm());
    } catch (e) {
      setStatus('error', 'Falha ao salvar');
      toast('Não foi possível salvar no navegador. Exporte para Excel para não perder os dados.');
    }
  }

  function scheduleSave() {
    setStatus('saving', 'Salvando…');
    clearTimeout(saveTimer);
    saveTimer = setTimeout(saveNow, 400);
  }

  window.addEventListener('beforeunload', () => { if (saveTimer) saveNow(); });
  document.addEventListener('visibilitychange', () => { if (document.hidden && saveTimer) saveNow(); });
  if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => {});

  let toastTimer;
  function toast(msg) {
    els.toast.textContent = msg;
    els.toast.classList.add('is-visible');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => els.toast.classList.remove('is-visible'), 3200);
  }

  /* ---------------------------------------------------------
     Etapas
     --------------------------------------------------------- */
  function showStage(stage) {
    const editor = stage === 'editor';
    els.stageUpload.hidden = editor;
    els.stageEditor.hidden = !editor;
    els.btnPrint.disabled = !editor;
    els.btnExport.disabled = !editor;
    els.btnReset.disabled = !state;
    els.step1.className = 'step ' + (editor ? 'is-done' : 'is-active');
    els.step2.className = 'step ' + (editor ? 'is-active' : '');
    if (!editor) els.status.innerHTML = '';
    window.scrollTo(0, 0);
  }

  /* ---------- Etapa 1 ---------- */
  const ICON_OK = '<svg viewBox="0 0 24 24"><circle cx="12" cy="12" r="9"/><path d="m8 12 3 3 5-6"/></svg>';
  const ICON_ERR = '<svg viewBox="0 0 24 24"><path d="M12 9v4M12 17h.01"/><path d="M10.3 3.9 2 18a2 2 0 0 0 1.7 3h16.6a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0z"/></svg>';
  const ICON_ARROW = '<svg viewBox="0 0 24 24"><path d="M5 12h14M13 6l6 6-6 6"/></svg>';

  function panelBusy(name) {
    els.panel.hidden = false;
    els.panel.className = 'panel panel--busy';
    els.panel.innerHTML = `<div class="panel__head"><div class="panel__icon"><div class="spinner"></div></div>
      <div><h2 class="panel__title">Convertendo…</h2><p class="panel__sub">${esc(name)}</p></div></div>`;
  }

  function panelError(name, msg) {
    els.panel.hidden = false;
    els.panel.className = 'panel panel--error';
    els.panel.innerHTML = `<div class="panel__head"><div class="panel__icon">${ICON_ERR}</div>
      <div><h2 class="panel__title">Não foi possível converter</h2><p class="panel__sub">${esc(name)}</p></div></div>
      <ul class="notes notes--warn"><li>${esc(msg)}</li></ul>
      <div class="panel__actions">
        ${state ? '<button class="btn" data-act="back">Voltar à sessão atual</button>' : ''}
        <button class="btn btn--primary" data-act="pick">Escolher outro arquivo</button>
      </div>`;
  }

  function panelConfirm(session, kept) {
    const items = session.items;
    const relatores = new Set(items.map(i => i.relator).filter(Boolean));
    const partes = items.reduce((n, i) => n + i.partes.length, 0);
    const mods = {};
    items.forEach(i => i.partes.forEach(p => { const m = p.modalidade || '—'; mods[m] = (mods[m] || 0) + 1; }));
    const modTxt = Object.entries(mods).map(([m, n]) => `${MOD_NAMES[m] || m} × ${n}`).join(' · ');

    const chips = Object.keys(COLUMN_NAMES).map(k =>
      `<span class="chip${session.columns.includes(k) ? '' : ' chip--off'}">${COLUMN_NAMES[k]}</span>`).join('')
      + '<span class="chip chip--new">+ Decisão e observações</span>';

    const notes = [];
    if (kept) notes.push(`<ul class="notes notes--info"><li>${kept} decisão(ões) já preenchida(s) desta sessão foram preservadas.</li></ul>`);
    if (session.warnings.length) notes.push(`<ul class="notes notes--warn">${session.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul>`);
    const replacing = state && norm(state.title) !== norm(session.title)
      ? `<ul class="notes notes--warn"><li>Ao continuar, a sessão atual (“${esc(state.title)}”) será substituída. Exporte-a para Excel antes, se precisar.</li></ul>` : '';

    els.panel.hidden = false;
    els.panel.className = 'panel';
    els.panel.innerHTML = `
      <div class="panel__head"><div class="panel__icon">${ICON_OK}</div>
        <div><h2 class="panel__title">Conversão concluída</h2><p class="panel__sub">${esc(session.sourceName)}</p></div></div>
      <p class="panel__session">${esc(session.title)}</p>
      <div class="stats">
        <div class="stat"><b>${items.length}</b><span>processos</span></div>
        <div class="stat"><b>${relatores.size}</b><span>relatores</span></div>
        <div class="stat"><b>${partes}</b><span>sustentações</span></div>
        <div class="stat"><b style="font-size:16px;padding:3px 0">${esc(modTxt)}</b><span>modalidade</span></div>
      </div>
      <p class="panel__label">Colunas reconhecidas</p>
      <div class="chips">${chips}</div>
      ${notes.join('')}${replacing}
      <div class="panel__actions">
        <button class="btn" data-act="${state ? 'back' : 'pick'}">${state ? 'Cancelar' : 'Escolher outro arquivo'}</button>
        <button class="btn btn--primary" data-act="open">Abrir layout otimizado ${ICON_ARROW}</button>
      </div>`;
    els.panel.querySelector('[data-act="open"]').focus();
  }

  async function handleFile(file) {
    if (!file) return;
    const name = file.name;
    const ext = name.split('.').pop().toLowerCase();
    showStage('upload');
    if (ext === 'doc') {
      panelError(name, 'O formato .doc (Word antigo) não é suportado. Abra no Word e use “Salvar como” → .docx.');
      return;
    }
    if (!['docx', 'xlsx', 'xls', 'ods', 'csv'].includes(ext)) {
      panelError(name, 'Formato não suportado. Use .docx, .xlsx, .xls, .ods ou .csv.');
      return;
    }
    panelBusy(name);
    try {
      const buf = await file.arrayBuffer();
      const rows = ext === 'docx' ? await readDocx(buf) : await readSheet(buf);
      const session = rowsToSession(rows, name.replace(/\.[^.]+$/, ''));
      session.sourceName = name;
      session.importedAt = new Date().toISOString();
      const kept = mergePrevious(session, state);
      pending = session;
      panelConfirm(session, kept);
    } catch (err) {
      console.error(err);
      panelError(name, err.message || String(err));
    }
  }

  function openPending() {
    if (!pending) return;
    const { warnings, columns, ...data } = pending;
    state = data;
    pending = null;
    els.panel.hidden = true;
    renderEditor();
    saveNow();
    toast('Layout gerado. O preenchimento é salvo automaticamente.');
  }

  /* ---------- Etapa 2 ---------- */
  function relatorLabel(name) {
    if (relatorLabels[name]) return relatorLabels[name];
    if (/\b(desa\.?|desembargadora|ju[ií]za)\b/i.test(name)) return 'RELATORA';
    // palpite simples: primeiro nome terminado em "a" (Rosana, Maria…). Corrigível com um clique.
    const first = name.replace(/^(des\.?|desembargador|juiz|conv\.?|\s)+/gi, '').split(/\s+/)[0] || '';
    return /a$/i.test(first) ? 'RELATORA' : 'RELATOR';
  }

  function cardHTML(it) {
    const partes = it.partes.length ? it.partes : [{ parte: '', advogado: '', modalidade: '' }];
    // Uma linha por sustentação: [V|P] Nome do advogado (condição da parte)
    const partesHTML = partes.map(p => {
      const mod = p.modalidade || '';
      return `
            <li class="party">
              <span class="mod-tag${mod ? '' : ' mod-tag--empty'}" title="${esc(MOD_NAMES[mod] || (mod ? 'Modalidade: ' + mod : 'Modalidade não informada'))}">${esc(mod || '?')}</span>
              <span class="party__text"><span class="party__name">${esc(p.advogado) || '—'}</span>${p.parte ? ` <span class="party__role">(${esc(p.parte)})</span>` : ''}</span>
            </li>`;
    }).join('');

    return `
    <section class="card" data-id="${esc(it.id)}">
      <div class="card__top">
        <div class="card__meta">
          <div class="ordem-box"><span class="ordem">${esc(it.ordem)}</span><span class="pauta">PAUTA: ${esc(it.pauta)}</span></div>
          <div class="meta-row"><b>PROCESSO:</b>${esc(it.processo)}</div>
          <div class="meta-row"><button type="button" class="label-toggle" data-relator="${esc(it.relator)}" title="Clique para alternar Relator / Relatora">${relatorLabel(it.relator)}:</button>${esc(it.relator)}</div>
        </div>
        <div class="card__parties" data-density="${partyDensity(partes)}">
          <h3 class="bar">PARTES E PROCURADORES</h3>
          <ul class="parties">${partesHTML}
          </ul>
        </div>
      </div>
      <h3 class="bar bar--full">DECISÃO E OBSERVAÇÕES</h3>
      <div class="card__decision">
        <textarea class="field field--desfecho" data-field="desfecho" rows="2" placeholder="Digite a decisão…" aria-label="Decisão do processo ${esc(it.processo)}">${esc(it.desfecho)}</textarea>
        <div class="print-text print-text--desfecho" data-mirror="desfecho">${esc(it.desfecho)}</div>
        <div class="obs-line${it.obs ? '' : ' is-empty'}">
          <b>OBS.:</b>
          <textarea class="field field--obs" data-field="obs" rows="1" placeholder="${esc(DEFAULT_OBS)}" aria-label="Observação do processo ${esc(it.processo)}">${esc(it.obs)}</textarea>
          <span class="print-text" data-mirror="obs">${esc(it.obs)}</span>
          <button type="button" class="obs-default" title="Preencher com: ${esc(DEFAULT_OBS)}"${it.obs ? ' hidden' : ''}>Usar padrão</button>
        </div>
      </div>
    </section>`;
  }

  /* Altura acompanha o conteúdo (quebras automáticas e Enter), tanto na decisão quanto na OBS. */
  function autosize(ta) {
    ta.style.height = 'auto';
    ta.style.height = ta.scrollHeight + 2 + 'px';
  }

  function renderEditor() {
    els.title.textContent = state.title;
    els.cards.innerHTML = state.items.map(cardHTML).join('');
    showStage('editor');
    els.cards.querySelectorAll('textarea').forEach(autosize);
    setStatus('ok', state.updatedAt ? 'Salvo às ' + hhmm(new Date(state.updatedAt)) : 'Salvo neste navegador');
  }

  function findItem(el) {
    const card = el.closest('.card');
    return card && state.items.find(it => it.id === card.dataset.id);
  }

  els.cards.addEventListener('input', e => {
    const field = e.target.dataset.field;
    if (!field) return;
    const it = findItem(e.target);
    if (!it) return;
    it[field] = e.target.value;
    const card = e.target.closest('.card');
    card.querySelector(`[data-mirror="${field}"]`).textContent = e.target.value;
    autosize(e.target);
    if (field === 'obs') {
      const empty = !e.target.value.trim();
      const line = e.target.closest('.obs-line');
      line.classList.toggle('is-empty', empty);
      line.querySelector('.obs-default').hidden = !empty;
    }
    scheduleSave();
  });

  els.cards.addEventListener('click', e => {
    const fill = e.target.closest('.obs-default');
    if (fill) {
      const ta = fill.closest('.obs-line').querySelector('[data-field="obs"]');
      ta.value = DEFAULT_OBS;
      ta.dispatchEvent(new Event('input', { bubbles: true }));
      ta.focus();
      return;
    }
    const btn = e.target.closest('.label-toggle');
    if (!btn) return;
    const name = btn.dataset.relator;
    const next = relatorLabel(name) === 'RELATOR' ? 'RELATORA' : 'RELATOR';
    relatorLabels[name] = next;
    try { localStorage.setItem(LABELS_KEY, JSON.stringify(relatorLabels)); } catch { /* ignora */ }
    els.cards.querySelectorAll('.label-toggle').forEach(b => {
      if (b.dataset.relator === name) b.textContent = next + ':';
    });
  });

  els.title.addEventListener('input', () => {
    state.title = els.title.textContent.replace(/\s+/g, ' ').trim();
    scheduleSave();
  });
  els.title.addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); els.title.blur(); } });
  els.title.addEventListener('paste', e => {
    e.preventDefault();
    document.execCommand('insertText', false, (e.clipboardData || window.clipboardData).getData('text').replace(/\s+/g, ' '));
  });

  let resizeTimer;
  window.addEventListener('resize', () => {
    clearTimeout(resizeTimer);
    resizeTimer = setTimeout(() => els.cards.querySelectorAll('textarea').forEach(autosize), 150);
  });

  /* ---------- Impressão ---------- */
  function printSheet() {
    if (!state) return;
    saveNow();
    const prevTitle = document.title;
    document.title = state.title;           // vira o nome sugerido do PDF
    window.addEventListener('afterprint', () => { document.title = prevTitle; }, { once: true });
    window.print();
  }

  /* ---------- Exportação Excel ---------- */
  function slug(s) {
    return String(s).replace(/[ªº]/g, '').normalize('NFD').replace(/[̀-ͯ]/g, '')
      .toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 60) || 'pauta';
  }

  /* Mesma convenção da pauta original: uma letra só quando todas as sustentações coincidem */
  function modCell(it) {
    const mods = it.partes.map(p => p.modalidade || '');
    return new Set(mods).size <= 1 ? (mods[0] || '') : mods.join('\n');
  }

  async function exportExcel() {
    if (!state) return;
    saveNow();
    try {
      await loadScript('xlsx');
      const XLSX = window.XLSX;
      const aoa = [[state.title], [], EXPORT_HEADERS];
      state.items.forEach(it => {
        aoa.push([
          it.ordem, it.pauta, it.relator, it.processo,
          it.partes.map(p => p.advogado).join('\n'),
          modCell(it),
          it.partes.map(p => p.parte).join('\n'),
          it.desfecho, it.obs
        ]);
      });
      const ws = XLSX.utils.aoa_to_sheet(aoa);
      ws['!cols'] = [6, 6, 26, 24, 34, 11, 26, 60, 36].map(wch => ({ wch }));
      ws['!merges'] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: EXPORT_HEADERS.length - 1 } }];
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, 'Pauta');
      const stamp = new Date().toISOString().slice(0, 10);
      XLSX.writeFile(wb, `${slug(state.title)}_${stamp}.xlsx`);
      toast('Arquivo Excel gerado.');
    } catch (err) {
      console.error(err);
      toast(err.message || 'Falha ao gerar o Excel.');
    }
  }

  /* ---------- Limpar ---------- */
  function resetSession() {
    if (!state) return;
    const ok = window.confirm('Apagar a sessão atual deste navegador?\n\nDica: exporte para Excel antes, se quiser guardar o que foi preenchido.');
    if (!ok) return;
    state = null;
    pending = null;
    try { localStorage.removeItem(STORAGE_KEY); } catch { /* ignora */ }
    els.panel.hidden = true;
    els.cards.innerHTML = '';
    showStage('upload');
    toast('Sessão apagada.');
  }

  /* ---------------------------------------------------------
     Eventos gerais
     --------------------------------------------------------- */
  const pickFile = () => { els.fileInput.value = ''; els.fileInput.click(); };

  els.btnImport.addEventListener('click', pickFile);
  els.btnPrint.addEventListener('click', printSheet);
  els.btnExport.addEventListener('click', exportExcel);
  els.btnReset.addEventListener('click', resetSession);
  els.fileInput.addEventListener('change', () => handleFile(els.fileInput.files[0]));

  els.dropzone.addEventListener('click', pickFile);
  els.dropzone.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); pickFile(); } });
  ['dragenter', 'dragover'].forEach(t => els.dropzone.addEventListener(t, e => { e.preventDefault(); els.dropzone.classList.add('is-over'); }));
  ['dragleave', 'drop'].forEach(t => els.dropzone.addEventListener(t, () => els.dropzone.classList.remove('is-over')));
  els.dropzone.addEventListener('drop', e => { e.preventDefault(); handleFile(e.dataTransfer.files[0]); });
  window.addEventListener('dragover', e => e.preventDefault());
  window.addEventListener('drop', e => e.preventDefault());

  els.panel.addEventListener('click', e => {
    const act = e.target.closest('[data-act]')?.dataset.act;
    if (act === 'pick') pickFile();
    if (act === 'open') openPending();
    if (act === 'back') { pending = null; els.panel.hidden = true; renderEditor(); }
  });

  document.addEventListener('keydown', e => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') {
      e.preventDefault();
      if (state) { saveNow(); toast('Salvo.'); }
    }
  });

  /* ---------- Inicialização ---------- */
  const saved = loadJSON(STORAGE_KEY);
  if (saved && Array.isArray(saved.items) && saved.items.length) {
    state = migrate(saved);
    renderEditor();
  } else {
    showStage('upload');
  }
})();
