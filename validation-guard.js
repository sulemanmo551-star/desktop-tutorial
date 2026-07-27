(() => {
  const DOWNLOAD_IDS = ['dlWorkbook', 'dlMaster', 'dlExceptions', 'dlZip'];
  let applying = false;

  function numberFrom(id) {
    const text = document.getElementById(id)?.textContent || '0';
    return Number(String(text).replace(/[^0-9.-]/g, '')) || 0;
  }

  function ensureBanner() {
    let banner = document.getElementById('invalidRunBanner');
    if (!banner) {
      banner = document.createElement('div');
      banner.id = 'invalidRunBanner';
      banner.style.cssText = 'display:none;margin:0 0 16px;padding:14px 16px;border:2px solid #b91c1c;border-radius:10px;background:#fee2e2;color:#7f1d1d;font-weight:800;line-height:1.45';
      const results = document.getElementById('results');
      results?.insertBefore(banner, results.firstChild);
    }
    return banner;
  }

  function setIfDifferent(element, property, value) {
    if (element[property] !== value) element[property] = value;
  }

  function setAttributeIfDifferent(element, name, value) {
    if (element.getAttribute(name) !== value) element.setAttribute(name, value);
  }

  function applyGuard() {
    if (applying) return;
    const results = document.getElementById('results');
    if (!results || results.hidden) return;
    applying = true;
    try {
      const gatesText = (document.getElementById('gates')?.textContent || '').toUpperCase();
      const unique = numberFrom('m2');
      const matched = numberFrom('m3');
      const hasFailedGate = gatesText.includes('FAIL');
      const zeroMatchFailure = unique > 0 && matched === 0;
      const invalid = hasFailedGate || zeroMatchFailure;
      const banner = ensureBanner();

      for (const id of DOWNLOAD_IDS) {
        const button = document.getElementById(id);
        if (!button) continue;
        setIfDifferent(button, 'disabled', invalid);
        setAttributeIfDifferent(button, 'aria-disabled', String(invalid));
        const opacity = invalid ? '0.45' : '';
        const cursor = invalid ? 'not-allowed' : '';
        const title = invalid ? 'Downloads are blocked because this run failed verification.' : '';
        if (button.style.opacity !== opacity) button.style.opacity = opacity;
        if (button.style.cursor !== cursor) button.style.cursor = cursor;
        if (button.title !== title) button.title = title;
      }

      const message = 'INVALID RUN — downloads are disabled. The selected workbook was not parsed into verified payment records. Do not use these counts or exception rows for boss review.';
      const display = invalid ? 'block' : 'none';
      if (banner.style.display !== display) banner.style.display = display;
      if (banner.textContent !== (invalid ? message : '')) banner.textContent = invalid ? message : '';
    } finally {
      applying = false;
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    ensureBanner();
    const results = document.getElementById('results');
    const gates = document.getElementById('gates');
    const metrics = ['m2','m3'].map(id => document.getElementById(id)).filter(Boolean);
    const observer = new MutationObserver(() => queueMicrotask(applyGuard));
    if (gates) observer.observe(gates, {subtree:true, childList:true, characterData:true});
    for (const metric of metrics) observer.observe(metric, {subtree:true, childList:true, characterData:true});
    if (results) observer.observe(results, {attributes:true, attributeFilter:['hidden']});
    applyGuard();
  });
})();