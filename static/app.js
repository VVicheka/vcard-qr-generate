// ---- Tabs ----
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab-btn').forEach(b => {
        b.classList.remove('border-emerald-600', 'text-emerald-700');
        b.classList.add('border-transparent', 'text-slate-500');
      });
      btn.classList.add('border-emerald-600', 'text-emerald-700');
      btn.classList.remove('border-transparent', 'text-slate-500');

      document.querySelectorAll('.tab-panel').forEach(p => p.classList.add('hidden'));
      document.getElementById('tab-' + btn.dataset.tab).classList.remove('hidden');
    });
  });

  // ---- Dynamic field rows ----
  const FIELD_TYPES = [
    { v: 'title', t: 'Position / Title' },
    { v: 'org',   t: 'Organization' },
    { v: 'phone', t: 'Phone' },
    { v: 'email', t: 'Email' },
    { v: 'url',   t: 'URL (with label)' },
    { v: 'note',  t: 'Note' },
  ];

  function fieldRow({ withColumn = false, columns = [], preset = {} } = {}) {
    const wrap = document.createElement('div');
    wrap.className = 'grid grid-cols-12 gap-2 items-center';

    const typeSel = document.createElement('select');
    typeSel.className = 'col-span-3 border rounded p-2';
    FIELD_TYPES.forEach(t => {
      const o = document.createElement('option');
      o.value = t.v; o.textContent = t.t;
      if (preset.type === t.v) o.selected = true;
      typeSel.appendChild(o);
    });

    const label = document.createElement('input');
    label.placeholder = 'Label (e.g. Telegram)';
    label.className = 'col-span-3 border rounded p-2';
    label.value = preset.label || '';

    let valueEl;
    if (withColumn) {
      valueEl = document.createElement('select');
      columns.forEach(c => {
        const o = document.createElement('option');
        o.value = c; o.textContent = c;
        if (preset.col === c) o.selected = true;
        valueEl.appendChild(o);
      });
    } else {
      valueEl = document.createElement('input');
      valueEl.placeholder = 'Value';
      valueEl.value = preset.value || '';
    }
    valueEl.className = 'col-span-5 border rounded p-2';

    const del = document.createElement('button');
    del.type = 'button';
    del.textContent = '✕';
    del.className = 'col-span-1 text-red-500 hover:text-red-700';
    del.onclick = () => wrap.remove();

    wrap.append(typeSel, label, valueEl, del);
    wrap._read = () => ({
      type: typeSel.value,
      label: label.value.trim(),
      ...(withColumn ? { col: valueEl.value } : { value: valueEl.value.trim() }),
    });
    return wrap;
  }

  // ---- Single mode ----
  const singleFields = document.getElementById('single-fields');
  [
    { type: 'phone', label: 'Mobile' },
    { type: 'email', label: 'Email' },
    { type: 'url',   label: 'Telegram' },
    { type: 'url',   label: 'Website' },
  ].forEach(p => singleFields.appendChild(fieldRow({ preset: p })));

  document.getElementById('add-field-single').onclick = () =>
    singleFields.appendChild(fieldRow());

  document.getElementById('single-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status = document.getElementById('single-status');
    status.textContent = 'Generating…';

    const fd = new FormData(e.target);

    const title = document.getElementById('single-title').value.trim();
    const org   = document.getElementById('single-org').value.trim();
    const extras = [];
    if (title) extras.push({ type: 'title', label: 'Position',     value: title });
    if (org)   extras.push({ type: 'org',   label: 'Organization', value: org });

    const dynamic = [...singleFields.children].map(c => c._read());
    fd.append('fields_json', JSON.stringify([...extras, ...dynamic]));

    try {
      const res = await fetch('/api/generate-single', { method: 'POST', body: fd });
      if (!res.ok) throw new Error((await res.json()).error || 'Failed');
      const blob = await res.blob();
      triggerDownload(blob, (fd.get('filename') || 'qr') + '.png');
      status.textContent = 'Done!';
    } catch (err) {
      status.textContent = 'Error: ' + err.message;
    }
  });

  // ---- Bulk mode ----
  const excelFile = document.getElementById('excel-file');
  const mappingArea = document.getElementById('mapping-area');
  const firstColSel = document.getElementById('first-col');
  const lastColSel = document.getElementById('last-col');
  const titleColSel = document.getElementById('title-col');
  const orgColSel = document.getElementById('org-col');
  const bulkFields = document.getElementById('bulk-fields');
  let currentColumns = [];

  excelFile.addEventListener('change', async () => {
    if (!excelFile.files[0]) return;
    const fd = new FormData();
    fd.append('file', excelFile.files[0]);

    const res = await fetch('/api/excel-columns', { method: 'POST', body: fd });
    const data = await res.json();
    if (!res.ok) { alert(data.error); return; }

    currentColumns = data.columns;
    [firstColSel, lastColSel].forEach(sel => {
      sel.innerHTML = '';
      currentColumns.forEach(c => {
        const o = document.createElement('option');
        o.value = c; o.textContent = c;
        sel.appendChild(o);
      });
    });

    [titleColSel, orgColSel].forEach(sel => {
      sel.innerHTML = '<option value="">— none —</option>';
      currentColumns.forEach(c => {
        const o = document.createElement('option');
        o.value = c; o.textContent = c;
        sel.appendChild(o);
      });
    });

    // sensible defaults if names match
    const findAny = (needles) =>
      currentColumns.find(c => needles.some(n => c.toLowerCase().includes(n)));
    const fCol = findAny(['first']);
    const lCol = findAny(['last']);
    const tCol = findAny(['position', 'title', 'job']);
    const oCol = findAny(['organization', 'company', 'org']);
    if (fCol) firstColSel.value = fCol;
    if (lCol) lastColSel.value  = lCol;
    if (tCol) titleColSel.value = tCol;
    if (oCol) orgColSel.value   = oCol;

    bulkFields.innerHTML = '';
    const auto = [
      { needles: ['phone', 'mobile', 'tel'], type: 'phone', label: 'Phone' },
      { needles: ['email', 'mail'],          type: 'email', label: 'Email' },
      { needles: ['website', 'site', 'url'], type: 'url',   label: 'Website' },
      { needles: ['telegram'],               type: 'url',   label: 'Telegram' },
    ];
    auto.forEach(a => {
      const col = findAny(a.needles);
      if (col) bulkFields.appendChild(fieldRow({
        withColumn: true, columns: currentColumns,
        preset: { type: a.type, label: a.label || col, col }
      }));
    });

    mappingArea.classList.remove('hidden');
  });

  document.getElementById('add-field-bulk').onclick = () =>
    bulkFields.appendChild(fieldRow({ withColumn: true, columns: currentColumns }));

  document.getElementById('bulk-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    const status = document.getElementById('bulk-status');
    status.textContent = 'Generating…';

    const fd = new FormData(e.target);

    const extras = [];
    if (titleColSel.value) extras.push({ col: titleColSel.value, label: 'Position',     type: 'title' });
    if (orgColSel.value)   extras.push({ col: orgColSel.value,   label: 'Organization', type: 'org'   });

    const mapping = {
      first_col: firstColSel.value,
      last_col: lastColSel.value,
      fields: [...extras, ...[...bulkFields.children].map(c => c._read())],
    };
    fd.append('mapping_json', JSON.stringify(mapping));

    try {
      const res = await fetch('/api/generate-bulk', { method: 'POST', body: fd });
      if (!res.ok) throw new Error((await res.json()).error || 'Failed');
      const blob = await res.blob();
      triggerDownload(blob, (fd.get('filename') || 'contacts_with_qr') + '.xlsx');
      status.textContent = 'Done!';
    } catch (err) {
      status.textContent = 'Error: ' + err.message;
    }
  });

  function triggerDownload(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.click();
    URL.revokeObjectURL(a.href);
  }