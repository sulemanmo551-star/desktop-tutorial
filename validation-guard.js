(() => {
  const DOWNLOAD_IDS = ['dlWorkbook', 'dlMaster', 'dlExceptions', 'dlZip'];

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

  function applyGuard() {
    const results = document.getElementById('results');
    if (!results || results.hidden) return;

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
      button.disabled = invalid;
      button.setAttribute('aria-disabled', String(invalid));
      button.style.opacity = invalid ? '0.45' : '';
      button.style.cursor = invalid ? 'not-allowed' : '';
      button.title = invalid ? 'Downloads are blocked because this run failed verification.' : '';
    }

    if (invalid) {
      banner.style.display = 'block';
      banner.textContent = 'INVALID RUN — downloads are disabled. The selected workbook was not parsed into verified payment records. Do not use these counts or exception rows for boss review.';
    } else {
      banner.style.display = 'none';
      banner.textContent = '';
    }
  }

  window.addEventListener('DOMContentLoaded', () => {
    ensureBanner();
    const target = document.getElementById('results');
    if (target) new MutationObserver(applyGuard).observe(target, {subtree:true, childList:true, characterData:true, attributes:true});
    applyGuard();
  });
})();
