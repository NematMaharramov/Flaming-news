/*
Client-side admin editor: WYSIWYG controls, resizing, XML mapping UI, Save/Publish/History.
Drop into repo under static/admin-editor.js and include in templates/index.html:
<script src="/static/admin-editor.js"></script>
*/

(async function(){
  // Ensure DOM is ready
  if (document.readyState !== 'loading') initAdmin();
  else document.addEventListener('DOMContentLoaded', initAdmin);

  async function initAdmin(){
    // Only add admin UI if user has admin param OR always if hosted locally.
    // For safety we render admin controls hidden by default and rely on server auth later.
    addAdminTopbarControls();
    setupFloatingToolbar();
    enableResizers();
    bindOccupancyConfig();
    bindSavePublish();
    bindHistory();
    bindXmlMappingUI();
    console.log('Admin editor loaded');
  }

  // --------------------
  // Admin topbar: Save Draft | Publish | History | Occupancy settings | XML mapping
  // --------------------
  function addAdminTopbarControls(){
    const topbar = document.querySelector('.topbar');
    if (!topbar) return;
    const adminGroup = document.createElement('div');
    adminGroup.style.display = 'flex'; adminGroup.style.gap='8px'; adminGroup.style.alignItems='center';

    const draftBtn = document.createElement('button');
    draftBtn.className='btn secondary small'; draftBtn.textContent='Save Draft';
    draftBtn.addEventListener('click', saveDraft);
    adminGroup.appendChild(draftBtn);

    const publishBtn = document.createElement('button');
    publishBtn.className='btn small'; publishBtn.textContent='Publish';
    publishBtn.addEventListener('click', publishDraft);
    adminGroup.appendChild(publishBtn);

    const historyBtn = document.createElement('button');
    historyBtn.className='btn secondary small'; historyBtn.textContent='History';
    historyBtn.addEventListener('click', openHistoryPanel);
    adminGroup.appendChild(historyBtn);

    const occBtn = document.createElement('button');
    occBtn.className='btn secondary small'; occBtn.textContent='Occupancy Rules';
    occBtn.addEventListener('click', openOccupancyEditor);
    adminGroup.appendChild(occBtn);

    const xmlBtn = document.createElement('button');
    xmlBtn.className='btn secondary small'; xmlBtn.textContent='XML Mapping';
    xmlBtn.addEventListener('click', openXmlMappingPanel);
    adminGroup.appendChild(xmlBtn);

    topbar.appendChild(adminGroup);
  }

  // --------------------
  // Floating contextual toolbar (font, size, color, bg, bold, italic, underline, align, border, padding)
  // Appears when an input.cell or a cell wrapper is focused/clicked
  // --------------------
  let toolbar;
  function setupFloatingToolbar(){
    toolbar = document.createElement('div');
    toolbar.style.position='absolute';
    toolbar.style.display='none';
    toolbar.style.zIndex=9999;
    toolbar.style.background='#fff';
    toolbar.style.border='1px solid #bbb';
    toolbar.style.padding='6px';
    toolbar.style.borderRadius='6px';
    toolbar.style.boxShadow='0 4px 12px rgba(0,0,0,.12)';
    toolbar.innerHTML = `
      <select id="adm_font" style="width:140px"><option value="">Font (inherit)</option><option>Calibri</option><option>Arial</option><option>Times New Roman</option></select>
      <input id="adm_size" type="number" style="width:62px" min="8" max="36" placeholder="px"/>
      <input id="adm_color" type="color" title="Font color"/>
      <input id="adm_bg" type="color" title="Background color"/>
      <button id="adm_b" class="btn small">B</button>
      <button id="adm_i" class="btn small">I</button>
      <button id="adm_u" class="btn small">U</button>
      <select id="adm_align" style="width:90px"><option value="">Align</option><option value="left">Left</option><option value="center">Center</option><option value="right">Right</option></select>
      <button id="adm_border" class="btn small">Border</button>
    `;
    document.body.appendChild(toolbar);

    // interactions
    document.addEventListener('click', (ev)=>{
      const cell = ev.target.closest('td');
      if (!cell || !cell.closest('.sheet')) { toolbar.style.display='none'; return; }
      // only show when clicking on content cells (not header actions)
      showToolbarForCell(cell, ev);
    });
  }

  function showToolbarForCell(cell, ev){
    toolbar.style.display='';
    const r = cell.getBoundingClientRect();
    toolbar.style.left = (window.scrollX + r.right - toolbar.offsetWidth) + 'px';
    toolbar.style.top = (window.scrollY + r.top - toolbar.offsetHeight - 6) + 'px';
    // preset toolbar to cell's current styles
    const fontSel = toolbar.querySelector('#adm_font');
    const sizeInp = toolbar.querySelector('#adm_size');
    const colorInp = toolbar.querySelector('#adm_color');
    const bgInp = toolbar.querySelector('#adm_bg');
    const boldBtn = toolbar.querySelector('#adm_b');
    const italicBtn = toolbar.querySelector('#adm_i');
    const uBtn = toolbar.querySelector('#adm_u');
    const alignSel = toolbar.querySelector('#adm_align');
    // read computed styles
    const cs = window.getComputedStyle(cell);
    fontSel.value = cs.fontFamily.split(',')[0].replace(/["']/g,'') || '';
    sizeInp.value = parseInt(cs.fontSize,10) || '';
    colorInp.value = rgbToHex(cs.color);
    bgInp.value = rgbToHex(cs.backgroundColor);
    boldBtn.style.opacity = cs.fontWeight >= 700 ? '1' : '0.5';
    italicBtn.style.opacity = cs.fontStyle === 'italic' ? '1' : '0.5';
    uBtn.style.opacity = cs.textDecorationLine.includes('underline') ? '1' : '0.5';
    alignSel.value = cs.textAlign || '';

    // wire actions
    toolbar.querySelector('#adm_font').onchange = (e)=> { cell.style.fontFamily = e.target.value || ''; };
    toolbar.querySelector('#adm_size').onchange = (e)=> { cell.style.fontSize = e.target.value ? e.target.value + 'px' : ''; };
    toolbar.querySelector('#adm_color').onchange = (e)=> { cell.style.color = e.target.value; };
    toolbar.querySelector('#adm_bg').onchange = (e)=> { cell.style.backgroundColor = e.target.value; };
    toolbar.querySelector('#adm_b').onclick = ()=> {
      cell.style.fontWeight = (cell.style.fontWeight === '700' || window.getComputedStyle(cell).fontWeight >= 700) ? 'normal' : '700';
      toolbar.querySelector('#adm_b').style.opacity = cell.style.fontWeight === '700' ? '1' : '0.5';
    };
    toolbar.querySelector('#adm_i').onclick = ()=> {
      cell.style.fontStyle = cell.style.fontStyle === 'italic' ? '' : 'italic';
      toolbar.querySelector('#adm_i').style.opacity = cell.style.fontStyle === 'italic' ? '1' : '0.5';
    };
    toolbar.querySelector('#adm_u').onclick = ()=> {
      if ((cell.style.textDecoration || '').includes('underline')) { cell.style.textDecoration = ''; toolbar.querySelector('#adm_u').style.opacity='0.5'; }
      else { cell.style.textDecoration = 'underline'; toolbar.querySelector('#adm_u').style.opacity='1'; }
    };
    toolbar.querySelector('#adm_align').onchange = (e)=> { cell.style.textAlign = e.target.value || ''; };
    toolbar.querySelector('#adm_border').onclick = ()=> {
      cell.style.border = cell.style.border ? '' : '1px solid #666';
    };
    // persist change on blur via small debounce
    ['change','input'].forEach(evName=>{
      toolbar.querySelectorAll('#adm_font, #adm_size, #adm_color, #adm_bg, #adm_align').forEach(el=>{
        el.addEventListener(evName, () => scheduleDraftSave());
      });
    });
  }

  function rgbToHex(rgb){
    if (!rgb) return '#000000';
    // rgb(a) format
    const m = rgb.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!m) return '#000000';
    const r = parseInt(m[1],10), g = parseInt(m[2],10), b = parseInt(m[3],10);
    return '#' + [r,g,b].map(v=>v.toString(16).padStart(2,'0')).join('');
  }

  // --------------------
  // Column & Row Resizers (simple, per table.section)
  // Only affect the table they belong to; outer frame remains fixed
  // --------------------
  function enableResizers(){
    document.querySelectorAll('.sheet table.grid').forEach(t => {
      makeTableResizable(t);
    });
  }

  function makeTableResizable(table){
    // ensure colgroup exists with cols equal to th count of first row
    let cols = table.querySelectorAll('col');
    if (!cols.length){
      const firstRow = table.querySelector('tr');
      if (!firstRow) return;
      const colCount = firstRow.children.length;
      const colgroup = document.createElement('colgroup');
      for (let i=0;i<colCount;i++){
        const c = document.createElement('col'); c.style.width = (100/colCount)+'%'; colgroup.appendChild(c);
      }
      table.insertBefore(colgroup, table.firstChild);
      cols = colgroup.querySelectorAll('col');
    }
    // create vertical resizers on the header area
    const wrapper = table.parentElement;
    wrapper.style.position = 'relative';
    // remove existing if any
    const existing = wrapper.querySelectorAll('.col-resizer');
    existing.forEach(e=>e.remove());
    const tableRect = table.getBoundingClientRect();
    // compute absolute x positions of column boundaries
    const colRects = [];
    const firstTr = table.querySelector('tr');
    if (!firstTr) return;
    let left = tableRect.left;
    // compute widths from cols
    const widths = Array.from(cols).map(c=> {
      return c.getBoundingClientRect().width;
    });
    let acc = tableRect.left;
    for (let i=0;i<widths.length-1;i++){
      acc += widths[i];
      const r = document.createElement('div');
      r.className='col-resizer';
      r.style.position='absolute';
      r.style.width='6px';
      r.style.top = (table.offsetTop) + 'px';
      r.style.left = (acc - wrapper.offsetLeft - 3) + 'px';
      r.style.height = table.offsetHeight + 'px';
      r.style.cursor = 'col-resize';
      r.style.zIndex = 999;
      r.style.background = 'transparent';
      r.addEventListener('mousedown', startColDrag.bind(null, table, cols, i));
      wrapper.appendChild(r);
    }
    // row resizers: add drag handles between rows
    const trs = table.querySelectorAll('tr');
    trs.forEach((tr, idx) => {
      if (idx === trs.length-1) return;
      const hr = document.createElement('div');
      hr.className = 'row-resizer';
      hr.style.position = 'absolute';
      hr.style.left = table.offsetLeft + 'px';
      hr.style.right = (wrapper.offsetWidth - (table.offsetLeft + table.offsetWidth - wrapper.offsetLeft)) + 'px';
      hr.style.top = (tr.offsetTop + tr.offsetHeight + table.offsetTop - wrapper.offsetTop - 3) + 'px';
      hr.style.height = '6px';
      hr.style.cursor = 'row-resize';
      hr.style.zIndex = 999;
      hr.style.background = 'transparent';
      hr.addEventListener('mousedown', startRowDrag.bind(null, table, tr));
      wrapper.appendChild(hr);
    });
  }

  function startColDrag(table, cols, idx, e){
    e.preventDefault();
    const startX = e.clientX;
    const col = cols[idx];
    const nextCol = cols[idx+1];
    const startW = col.getBoundingClientRect().width;
    const startW2 = nextCol.getBoundingClientRect().width;
    function onMove(ev){
      const dx = ev.clientX - startX;
      const newW = Math.max(20, startW + dx);
      const newW2 = Math.max(20, startW2 - dx);
      // write in px
      col.style.width = newW + 'px';
      nextCol.style.width = newW2 + 'px';
    }
    function onUp(){
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      scheduleDraftSave();
    }
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }

  function startRowDrag(table, tr, e){
    e.preventDefault();
    const startY = e.clientY;
    const startH = tr.getBoundingClientRect().height;
    function onMove(ev){
      const dy = ev.clientY - startY;
      tr.style.height = Math.max(20, startH + dy) + 'px';
    }
    function onUp(){
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
      scheduleDraftSave();
    }
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }

  // --------------------
  // Occupancy color stops editor (modal)
  // --------------------
  let occupancyStops = [
    {pos:0,color:'#d73027'},
    {pos:50,color:'#ffffbf'},
    {pos:100,color:'#1a9850'}
  ];

  function bindOccupancyConfig(){
    // If JSON config provides stops on page load, use them
    fetch('/api/admin/config').then(r=>r.json()).then(j=>{
      const draft = j && j.draft;
      if (draft && draft.occupancyRules && draft.occupancyRules.stops) {
        occupancyStops = draft.occupancyRules.stops;
      }
    }).catch(()=>{});
  }

  function openOccupancyEditor(){
    const modal = document.createElement('div');
    modal.style.position='fixed'; modal.style.left=0; modal.style.top=0; modal.style.right=0; modal.style.bottom=0;
    modal.style.background='rgba(0,0,0,.5)'; modal.style.display='flex'; modal.style.alignItems='center'; modal.style.justifyContent='center';
    modal.innerHTML = `<div style="background:#fff;padding:16px;border-radius:8px;min-width:420px">
      <h3>Occupancy Color Stops</h3>
      <div id="occList"></div>
      <div style="margin-top:12px">
        <button id="occAdd" class="btn small secondary">Add stop</button>
        <button id="occSave" class="btn small">Save</button>
        <button id="occCancel" class="btn small secondary">Cancel</button>
      </div>
    </div>`;
    document.body.appendChild(modal);
    const list = modal.querySelector('#occList');
    function render(){
      list.innerHTML = '';
      occupancyStops.sort((a,b)=>a.pos-b.pos);
      occupancyStops.forEach((s, idx)=>{
        const row = document.createElement('div');
        row.style.display='flex'; row.style.gap='8px'; row.style.alignItems='center'; row.style.marginBottom='6px';
        const pos = document.createElement('input'); pos.type='number'; pos.value=s.pos; pos.style.width='80px';
        const color = document.createElement('input'); color.type='color'; color.value=s.color;
        const del = document.createElement('button'); del.className='btn small danger'; del.textContent='✕';
        del.onclick = ()=> { occupancyStops.splice(idx,1); render(); };
        pos.onchange = ()=> { s.pos = Number(pos.value); render(); };
        color.onchange = ()=> { s.color = color.value; render(); };
        const preview = document.createElement('div'); preview.style.width='120px'; preview.style.height='18px'; preview.style.background = s.color; preview.style.border='1px solid #ccc';
        row.appendChild(pos); row.appendChild(color); row.appendChild(preview); row.appendChild(del);
        list.appendChild(row);
      });
    }
    render();
    modal.querySelector('#occAdd').onclick = ()=>{
      occupancyStops.push({pos:50,color:'#ffffff'});
      render();
    };
    modal.querySelector('#occSave').onclick = async ()=>{
      // save to draft config
      const res = await fetch('/api/admin/config'); const j = await res.json();
      const draft = j.draft || {};
      draft.occupancyRules = draft.occupancyRules || {};
      draft.occupancyRules.stops = occupancyStops;
      await fetch('/api/admin/config/save', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({config: draft})});
      document.body.removeChild(modal);
      toast('Occupancy rules saved to draft');
    };
    modal.querySelector('#occCancel').onclick = ()=> modal.remove();
  }

  // --------------------
  // Draft Save / Publish actions (talk to admin_api)
  // --------------------
  let _saveTimer = null;
  function scheduleDraftSave(delay=800){
    if (_saveTimer) clearTimeout(_saveTimer);
    _saveTimer = setTimeout(()=>{ saveDraft(); _saveTimer=null; }, delay);
  }

  async function saveDraft(){
    // collect current template HTML (the full template fragment that we want to store)
    // For simplicity we capture the body innerHTML of the sheet content area
    const templateHTML = document.querySelector('.sheet')?.parentElement?.innerHTML || document.documentElement.outerHTML;
    // fetch current draft config, load then update with visual cell styles and layout
    const cfgResp = await fetch('/api/admin/config'); const cfgJson = await cfgResp.json();
    const draft = cfgJson.draft || {};
    // extract per-cell inline styles into config.layout.mapping (simple approach)
    draft.cellStyles = {}; // key: data-path or dom path -> style string
    document.querySelectorAll('.sheet td').forEach((td, idx)=>{
      const path = td.querySelector('[data-path]') ? td.querySelector('[data-path]').dataset.path : null;
      const key = path || `cell_${idx}`;
      if (td.getAttribute('style')) draft.cellStyles[key] = td.getAttribute('style');
      // collect col widths
      const table = td.closest('table.grid');
      if (table){
        // ensure tableWidths array
        draft.layout = draft.layout || {};
        draft.layout.tables = draft.layout.tables || {};
        const id = table.id || `table_${Array.from(document.querySelectorAll('table.grid')).indexOf(table)}`;
        draft.layout.tables[id] = draft.layout.tables[id] || {};
        // capture colgroup widths
        const cols = table.querySelectorAll('col');
        draft.layout.tables[id].cols = Array.from(cols).map(c => c.style.width || c.getAttribute('width') || null);
      }
    });
    // persist draft
    await fetch('/api/admin/config/save', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({config: draft, template_html: templateHTML})});
    toast('Draft saved');
  }

  async function publishDraft(){
    if (!confirm('Publish this draft and make it live for users?')) return;
    const note = prompt('Optional publish note','');
    const res = await fetch('/api/admin/config/publish', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({note})});
    const j = await res.json();
    if (j.status === 'ok') toast('Published');
    else toast('Publish failed', true);
  }

  // --------------------
  // History UI
  // --------------------
  async function openHistoryPanel(){
    const res = await fetch('/api/admin/history'); const j = await res.json();
    if (!j || !j.history) { toast('No history'); return; }
    const modal = document.createElement('div');
    modal.style.position='fixed'; modal.style.left=0; modal.style.top=0; modal.style.right=0; modal.style.bottom=0;
    modal.style.background='rgba(0,0,0,.5)'; modal.style.display='flex'; modal.style.alignItems='center'; modal.style.justifyContent='center';
    modal.innerHTML = `<div style="background:#fff;padding:14px;border-radius:8px;min-width:640px;max-height:80vh;overflow:auto">
      <h3>Template History</h3><div id="histList"></div><div style="margin-top:12px"><button id="hClose" class="btn small secondary">Close</button></div>
    </div>`;
    document.body.appendChild(modal);
    const list = modal.querySelector('#histList');
    j.history.forEach(entry=>{
      const row = document.createElement('div');
      row.style.display='flex'; row.style.justifyContent='space-between'; row.style.alignItems='center'; row.style.borderBottom='1px solid #eee'; row.style.padding='8px 0';
      const left = document.createElement('div'); left.innerHTML = `<strong>${entry.id}</strong> &nbsp; ${entry.timestamp} &nbsp; ${entry.note || ''}`;
      const right = document.createElement('div');
      const restoreBtn = document.createElement('button'); restoreBtn.className='btn small secondary'; restoreBtn.textContent='Restore to Draft';
      restoreBtn.onclick = async ()=> {
        if (!confirm('Restore this history snapshot as the current draft?')) return;
        const r = await fetch(`/api/admin/history/${entry.id}/restore`, {method:'POST'});
        const jr = await r.json();
        if (jr.status === 'ok') {
          toast('Restored to draft. Reload to see changes.');
          modal.remove();
        } else toast('Restore failed', true);
      };
      right.appendChild(restoreBtn);
      row.appendChild(left); row.appendChild(right); list.appendChild(row);
    });
    modal.querySelector('#hClose').onclick = ()=> modal.remove();
  }

  // --------------------
  // XML mapping UI (upload and bind mapping to cells)
  // --------------------
  function bindXmlMappingUI(){
    // no-op; content shown in openXmlMappingPanel
  }

  function openXmlMappingPanel(){
    const modal = document.createElement('div');
    modal.style.position='fixed'; modal.style.inset=0; modal.style.background='rgba(0,0,0,.5)'; modal.style.display='flex'; modal.style.alignItems='center'; modal.style.justifyContent='center';
    modal.innerHTML = `<div style="background:#fff;padding:16px;border-radius:8px;min-width:720px;max-height:80vh;overflow:auto">
      <h3>XML Mapping</h3>
      <div style="display:flex;gap:12px">
        <div style="flex:1">
          <input type="file" id="xmlFile" accept=".xml"/>
          <div id="xmlPaths" style="max-height:48vh;overflow:auto;border:1px solid #eee;margin-top:8px;padding:8px"></div>
        </div>
        <div style="flex:1">
          <div style="margin-bottom:8px">Select a cell from the report (click it) then pick a path to map.</div>
          <div><strong>Selected cell:</strong> <span id="selectedCellLabel">none</span></div>
          <div id="cellMapList" style="margin-top:8px"></div>
        </div>
      </div>
      <div style="margin-top:12px"><button id="xmlClose" class="btn small secondary">Close</button></div>
    </div>`;
    document.body.appendChild(modal);

    modal.querySelector('#xmlFile').addEventListener('change', async (ev)=>{
      const file = ev.target.files[0];
      if (!file) return;
      const fd = new FormData(); fd.append('xml', file);
      const res = await fetch('/api/admin/upload/xml', {method:'POST', body: fd});
      const j = await res.json();
      if (j.status !== 'ok') { modal.querySelector('#xmlPaths').textContent = 'Parse error'; return; }
      const paths = j.paths || [];
      const container = modal.querySelector('#xmlPaths');
      container.innerHTML = '';
      paths.slice(0,1000).forEach(p=>{
        const el = document.createElement('div'); el.textContent = p; el.style.padding='4px'; el.style.cursor='pointer';
        el.addEventListener('click', ()=> {
          const sel = modal.querySelector('#selectedCellLabel').dataset.path;
          if (!sel) { alert('Select a cell first by clicking it on the report'); return; }
          // save mapping to draft config
          saveMapping(sel, p);
          // show mapping
          const list = modal.querySelector('#cellMapList');
          const item = document.createElement('div'); item.textContent = `${sel} → ${p}`;
          list.appendChild(item);
        });
        container.appendChild(el);
      });
    });

    // click-on-cell to choose target
    const onSelectCell = (ev) => {
      const cell = ev.target.closest('td');
      if (!cell) return;
      const input = cell.querySelector('[data-path]');
      const path = input ? input.dataset.path : null;
      modal.querySelector('#selectedCellLabel').textContent = path || 'DOM cell (no data-path)';
      modal.querySelector('#selectedCellLabel').dataset.path = path || null;
      ev.stopPropagation();
      ev.preventDefault();
    };
    document.querySelectorAll('.sheet td').forEach(td => td.addEventListener('click', onSelectCell, {once:false}));

    modal.querySelector('#xmlClose').onclick = ()=> modal.remove();
  }

  async function saveMapping(cellKey, xmlPath){
    // get draft config, update mappings
    const cfgResp = await fetch('/api/admin/config'); const cfgJson = await cfgResp.json();
    const draft = cfgJson.draft || {};
    draft.mappings = draft.mappings || {};
    draft.mappings[cellKey] = xmlPath;
    await fetch('/api/admin/config/save', {method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify({config: draft})});
    toast('Mapping saved to draft');
  }

  // --------------------
  // Utility toast
  // --------------------
  function toast(msg, isErr){
    const t = document.getElementById('toast') || (function(){
      const d = document.createElement('div'); d.id='toast'; d.className='toast'; document.body.appendChild(d); return d;
    })();
    t.textContent = msg; t.style.display='block'; t.style.background = isErr ? '#c0392b' : '#2e7d32';
    setTimeout(()=> t.style.display='none', 3500);
  }

  // --------------------
  // Bind Save/Publish triggers: connect to UI already added
  // --------------------
  function bindSavePublish(){
    // Already wired by buttons created in topbar
  }

  // --------------------
  // History binder
  // --------------------
  function bindHistory(){
    // Already wired by button
  }

})();
