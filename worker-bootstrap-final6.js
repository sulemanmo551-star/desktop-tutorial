/* FINAL6 bootstrap: unique filename prevents stale worker caches. */
(() => {
  const XLSX_URL='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  const nativeImportScripts=self.importScripts.bind(self);
  nativeImportScripts(XLSX_URL);
  const nativeRead=XLSX.read.bind(XLSX);
  XLSX.read=(data,options={})=>nativeRead(data,{...options,dense:false});
  self.importScripts=(...urls)=>{for(const url of urls){if(String(url).includes('xlsx@0.18.5'))continue;nativeImportScripts(url);}};
  nativeImportScripts('worker-engine.js?v=20260727-final6-base');
  nativeImportScripts('worker-matching-fix.js?v=20260727-final6-match');
  nativeImportScripts('worker-final-fix-final6.js?v=20260727-final6-verified');
  if(typeof progress==='function')progress(0,'BUILD FINAL6 ACTIVE');
})();