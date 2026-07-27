/*
 * Compatibility bootstrap for SheetJS 0.18.5.
 *
 * The existing engine's cell reader supports sparse worksheets and the newer
 * `!data` dense representation. SheetJS 0.18.5 represents dense worksheets as
 * arrays instead, which made every cell appear blank. Force sparse worksheets
 * so the existing bounded parser reads the real GTMDJD workbook correctly.
 */
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

  nativeImportScripts('worker-engine.js?v=20260727-schemafix1');
})();
