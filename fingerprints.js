(() => {
  const ORIGINAL_SOURCES = {
    '8ffe615d7b035659910718f93e92565aa3cb5f72b9908742b0e5c9bffbab484b': 'aug 1 24 to dec 3124 check only cases.xls',
    '5c6d464a9b2406cedd574c1ca6e1c3b6df0b8af3ec5947f342db963daae12ec8': 'full income cases aug1 to dec 3124.xls',
    '08984bfd5cce9e26a000fc92ba195f896b06d61b15c6a0be6b55b3cd7116a972': 'jan 1 to dec 31 25 check income cases.xls',
    'ed9df53acf9b7f14dfefa81425d70aa6801a72d8c13a600bab660df724663c3d': 'jan 1 to dec 31 25 full income cases.xls',
    'f0cb0905ec140cddfbc514500fe847994c8c01c9ec138dbdc88772218ac8192f': 'jan 1 to jun 30 26 check caases.xls',
    '521176408a5c12d8d1775b4a126c2ad184771dfdafb38c01ea5ce4ee49eae1fe': 'jan 1 to jun 30 26 full income cases.xls'
  };
  const EXPECTED_NAMES = new Set(Object.values(ORIGINAL_SOURCES).map(x => x.toLowerCase()));

  async function digest(file) {
    const bytes = await file.arrayBuffer();
    const raw = await crypto.subtle.digest('SHA-256', bytes);
    return [...new Uint8Array(raw)].map(b => b.toString(16).padStart(2, '0')).join('');
  }

  function esc(value) {
    return String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  }

  async function verifySelectedFiles(event) {
    const box = document.getElementById('fingerprintResults');
    const summary = document.getElementById('fingerprintSummary');
    if (!box || !summary) return;
    const files = [...event.target.files];
    if (!files.length) {
      box.innerHTML = '';
      summary.textContent = 'No MyCases files selected yet.';
      summary.className = 'fingerprint-summary';
      return;
    }

    box.innerHTML = '<div>Checking SHA-256 fingerprints…</div>';
    const rows = [];
    let exact = 0;
    for (const file of files) {
      const sha = await digest(file);
      const expected = ORIGINAL_SOURCES[sha];
      if (expected) {
        exact += 1;
        rows.push({name:file.name, status:'EXACT ORIGINAL VERIFIED SOURCE', detail:expected, kind:'pass'});
      } else if (EXPECTED_NAMES.has(file.name.toLowerCase())) {
        rows.push({name:file.name, status:'DIFFERENT CONTENT', detail:'Filename matches an original source, but the file fingerprint does not. Do not use it for the historical benchmark.', kind:'fail'});
      } else {
        rows.push({name:file.name, status:'NOT IN ORIGINAL SIX-SOURCE BENCHMARK', detail:'This may be valid for a new project, but it is not an exact original benchmark source.', kind:'warn'});
      }
    }

    summary.textContent = `${exact} of 6 exact original source files detected.`;
    summary.className = `fingerprint-summary ${exact === 6 ? 'pass' : 'warn'}`;
    box.innerHTML = rows.map(r => `<div class="fingerprint-row ${r.kind}"><b>${esc(r.status)}</b><span>${esc(r.name)}</span><small>${esc(r.detail)}</small></div>`).join('');
  }

  window.addEventListener('DOMContentLoaded', () => {
    document.getElementById('mycases')?.addEventListener('change', verifySelectedFiles);
  });
})();
