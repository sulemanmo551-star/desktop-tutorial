/* Compatibility bootstrap for SheetJS 0.18.5 plus verified matching rules. */
(() => {
  const XLSX_URL='https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  const nativeImportScripts=self.importScripts.bind(self);
  nativeImportScripts(XLSX_URL);

  const nativeRead=XLSX.read.bind(XLSX);
  XLSX.read=(data,options={})=>nativeRead(data,{...options,dense:false});

  self.importScripts=(...urls)=>{
    for(const url of urls){
      if(String(url).includes('xlsx@0.18.5')) continue;
      nativeImportScripts(url);
    }
  };

  nativeImportScripts('worker-engine.js?v=20260727-logicfix1');
  nativeImportScripts('worker-matching-fix.js?v=20260727-logicfix1');
})();
