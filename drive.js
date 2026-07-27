/* Google Drive import for GTMDJD Portal.
   OAuth tokens and selected files stay in browser memory only.
   Client configuration is saved only in this browser's localStorage. */
(() => {
  const DRIVE_CONFIG_KEY = 'gtmdjd_google_drive_config_v1';
  const DRIVE_SCOPE = 'openid email https://www.googleapis.com/auth/drive.file';
  const EXPECTED_ACCOUNT = 'care@pharmacy.net';
  let tokenClient = null;
  let accessToken = '';
  let pickerLoaded = false;
  let activeTarget = null;
  const $d = id => document.getElementById(id);

  function getConfig() {
    try { return JSON.parse(localStorage.getItem(DRIVE_CONFIG_KEY) || '{}'); }
    catch { return {}; }
  }
  function saveConfig() {
    const config = {
      clientId: $d('driveClientId').value.trim(),
      apiKey: $d('driveApiKey').value.trim(),
      appId: $d('driveAppId').value.trim()
    };
    localStorage.setItem(DRIVE_CONFIG_KEY, JSON.stringify(config));
    setDriveStatus('Google configuration saved on this browser only.', 'ok');
    initTokenClient();
  }
  function clearConfig() {
    localStorage.removeItem(DRIVE_CONFIG_KEY);
    ['driveClientId','driveApiKey','driveAppId'].forEach(id => $d(id).value = '');
    accessToken = '';
    tokenClient = null;
    setDriveStatus('Google configuration cleared.', 'warn');
  }
  function setDriveStatus(message, kind='') {
    const el = $d('driveStatus');
    if (!el) return;
    el.textContent = message;
    el.className = `drive-status ${kind}`;
  }
  function configComplete() {
    const c = getConfig();
    return Boolean(c.clientId && c.apiKey && c.appId);
  }
  function loadConfigIntoForm() {
    const c = getConfig();
    $d('driveClientId').value = c.clientId || '';
    $d('driveApiKey').value = c.apiKey || '';
    $d('driveAppId').value = c.appId || '';
  }
  function initTokenClient() {
    const c = getConfig();
    if (!c.clientId || !window.google?.accounts?.oauth2) return;
    tokenClient = google.accounts.oauth2.initTokenClient({
      client_id: c.clientId,
      scope: DRIVE_SCOPE,
      callback: async response => {
        if (response.error) {
          setDriveStatus(`Google authorization failed: ${response.error}`, 'bad');
          return;
        }
        accessToken = response.access_token;
        const email = await getSignedInEmail();
        if (email && email.toLowerCase() !== EXPECTED_ACCOUNT.toLowerCase()) {
          setDriveStatus(`Connected as ${email}. Expected ${EXPECTED_ACCOUNT}; switch accounts before importing protected files.`, 'warn');
        } else {
          setDriveStatus(`Connected securely${email ? ` as ${email}` : ''}. Choose files from Drive.`, 'ok');
        }
        openPicker(activeTarget);
      }
    });
  }
  async function getSignedInEmail() {
    try {
      const r = await fetch('https://openidconnect.googleapis.com/v1/userinfo', {
        headers: { Authorization: `Bearer ${accessToken}` }
      });
      if (!r.ok) return '';
      return (await r.json()).email || '';
    } catch { return ''; }
  }
  function requestDrive(target) {
    activeTarget = target;
    if (!configComplete()) {
      $d('driveSetup').open = true;
      setDriveStatus('Complete the one-time Google Cloud configuration below first.', 'warn');
      return;
    }
    if (!tokenClient) initTokenClient();
    if (!tokenClient) {
      setDriveStatus('Google Identity library is still loading. Wait a moment and try again.', 'warn');
      return;
    }
    if (!accessToken) tokenClient.requestAccessToken({ prompt: 'select_account consent' });
    else openPicker(target);
  }
  function onPickerApiLoad() {
    pickerLoaded = true;
  }
  function openPicker(target) {
    if (!accessToken) return;
    if (!pickerLoaded || !window.google?.picker) {
      setDriveStatus('Google Drive picker is loading. Try again in a few seconds.', 'warn');
      return;
    }
    const c = getConfig();
    const view = new google.picker.DocsView(google.picker.ViewId.DOCS)
      .setIncludeFolders(true)
      .setSelectFolderEnabled(false);
    const picker = new google.picker.PickerBuilder()
      .setDeveloperKey(c.apiKey)
      .setAppId(c.appId)
      .setOAuthToken(accessToken)
      .setOrigin(window.location.protocol + '//' + window.location.host)
      .addView(view)
      .enableFeature(google.picker.Feature.MULTISELECT_ENABLED)
      .setTitle(target === 'mycases' ? 'Select MyCases XLS/XLSX/CSV files' : 'Select GTMDJD PDFs, spreadsheets, images, or ZIP files')
      .setCallback(data => pickerCallback(data, target))
      .build();
    picker.setVisible(true);
  }
  async function pickerCallback(data, target) {
    if (data.action !== google.picker.Action.PICKED) return;
    const docs = data.docs || [];
    setDriveStatus(`Importing ${docs.length} selected Drive file(s)...`, '');
    try {
      const files = [];
      for (let i = 0; i < docs.length; i++) {
        setDriveStatus(`Importing ${i + 1}/${docs.length}: ${docs[i].name}`, '');
        files.push(await downloadDriveFile(docs[i]));
      }
      if (target === 'mycases') {
        S.my = await expand([...S.my, ...files]);
        $d('mycasesList').textContent = `${S.my.length} file(s) ready (${files.length} imported from Drive)`;
      } else {
        S.toc = await expand([...S.toc, ...files]);
        $d('tocList').textContent = `${S.toc.length} file(s) ready (${files.length} imported from Drive)`;
      }
      setDriveStatus(`${files.length} file(s) imported securely into browser memory. Nothing was added to GitHub or Airtable.`, 'ok');
    } catch (err) {
      console.error(err);
      setDriveStatus(`Drive import failed: ${err.message}`, 'bad');
    }
  }
  async function downloadDriveFile(doc) {
    const id = doc.id;
    let name = doc.name || `drive-file-${id}`;
    let mime = doc.mimeType || '';
    let url;
    if (mime === 'application/vnd.google-apps.spreadsheet') {
      url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}/export?mimeType=${encodeURIComponent('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}`;
      if (!/\.xlsx$/i.test(name)) name += '.xlsx';
      mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    } else if (mime.startsWith('application/vnd.google-apps.')) {
      throw new Error(`${name} is a Google Workspace file type that this comparison does not accept. Export it as XLSX, PDF, or CSV first.`);
    } else {
      url = `https://www.googleapis.com/drive/v3/files/${encodeURIComponent(id)}?alt=media`;
    }
    const response = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } });
    if (!response.ok) {
      let detail = '';
      try { detail = (await response.json()).error?.message || ''; } catch {}
      throw new Error(`${name}: ${detail || `Google Drive returned ${response.status}`}`);
    }
    const blob = await response.blob();
    return new File([blob], name, { type: mime || blob.type, lastModified: Date.now() });
  }
  function disconnectDrive() {
    if (accessToken && window.google?.accounts?.oauth2) google.accounts.oauth2.revoke(accessToken, () => {});
    accessToken = '';
    setDriveStatus('Disconnected. No Drive token is retained.', 'warn');
  }
  function initDriveUi() {
    loadConfigIntoForm();
    $d('saveDriveConfig').addEventListener('click', saveConfig);
    $d('clearDriveConfig').addEventListener('click', clearConfig);
    $d('driveMyCases').addEventListener('click', () => requestDrive('mycases'));
    $d('driveToc').addEventListener('click', () => requestDrive('toc'));
    $d('disconnectDrive').addEventListener('click', disconnectDrive);
    if (configComplete()) setDriveStatus('Drive connector configured. Connect the authorized Google account.', '');
  }
  window.gtmdjdPickerLoaded = onPickerApiLoad;
  window.addEventListener('load', () => {
    initDriveUi();
    initTokenClient();
    if (window.gapi) gapi.load('picker', { callback: onPickerApiLoad });
  });
})();