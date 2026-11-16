// Motor de cálculo: lê a planilha (.xlsm/.xlsx) e recalcula todas as fórmulas no navegador.
// Funciona tanto no navegador (globais XLSX/HyperFormula) quanto no Node (para testes).
(function (root) {
  const XLSX_ = root.XLSX || (typeof require !== 'undefined' && require('xlsx'));
  const HF_ = root.HyperFormula || (typeof require !== 'undefined' && require('hyperformula'));
  const HyperFormula = HF_.HyperFormula || HF_;
  const FunctionPlugin = HF_.FunctionPlugin || HyperFormula.FunctionPlugin;
  const FunctionArgumentType = HF_.FunctionArgumentType || HyperFormula.FunctionArgumentType;

  const num = v => typeof v === 'number' ? v : (v && typeof v === 'object' && typeof v.val === 'number') ? v.val : null;

  // XIRR (TIR com datas), não existe no HyperFormula
  class XirrPlugin extends FunctionPlugin {
    xirr(ast, state) {
      return this.runFunction(ast.args, state, this.metadata('XIRR'), (vals, dates, guess) => {
        const v = vals.valuesFromTopLeftCorner(), d = dates.valuesFromTopLeftCorner();
        const cf = [];
        for (let k = 0; k < v.length; k++) {
          const x = num(v[k]), t = num(d[k]);
          if (x !== null && x !== 0 && t !== null) cf.push([x, t]);
        }
        if (cf.length < 2) return 0;
        const t0 = cf[0][1];
        let r = guess ?? 0.1;
        for (let it = 0; it < 100; it++) {
          let f = 0, df = 0;
          for (const [x, t] of cf) { const e = (t - t0) / 365; f += x / Math.pow(1 + r, e); df -= e * x / Math.pow(1 + r, e + 1); }
          const nr = r - f / df;
          if (!isFinite(nr)) break;
          if (Math.abs(nr - r) < 1e-10) return nr;
          r = nr;
        }
        return r;
      });
    }
  }
  XirrPlugin.implementedFunctions = {
    XIRR: { method: 'xirr', parameters: [
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.RANGE },
      { argumentType: FunctionArgumentType.NUMBER, optionalArg: true, defaultValue: 0.1 } ] }
  };
  HyperFormula.registerFunctionPlugin(XirrPlugin, { enGB: { XIRR: 'XIRR' } });

  // HyperFormula exige TRUE()/FALSE() em vez de TRUE/FALSE
  const fixFormula = f => f.split('"').map((p, i) => i % 2 ? p : p.replace(/\b(TRUE|FALSE)\b(?!\s*\()/g, '$1()')).join('"');

  function load(arrayBuffer) {
    const wb = XLSX_.read(arrayBuffer, { type: 'array', cellFormula: true });
    const sheets = {};
    for (const name of wb.SheetNames) {
      const ws = wb.Sheets[name];
      if (!ws['!ref']) { sheets[name] = [[]]; continue; }
      const r = XLSX_.utils.decode_range(ws['!ref']), data = [];
      for (let R = 0; R <= r.e.r; R++) {
        const row = [];
        for (let C = 0; C <= r.e.c; C++) {
          const c = ws[XLSX_.utils.encode_cell({ r: R, c: C })];
          row.push(!c ? null : c.f ? '=' + fixFormula(c.f) : (c.t === 'e' ? null : c.v));
        }
        data.push(row);
      }
      sheets[name] = data;
    }
    const hf = HyperFormula.buildFromSheets(sheets, {
      licenseKey: 'gpl-v3',
      leapYear1900: true, nullDate: { year: 1899, month: 12, day: 31 }  // datas iguais ao Excel
    });
    const addr = ref => { const [s, c] = ref.split('!'); return hf.simpleCellAddressFromString(c, hf.getSheetId(s)); };
    return {
      set(values) { hf.batch(() => { for (const [ref, v] of Object.entries(values)) hf.setCellContents(addr(ref), v); }); },
      get(ref) { const v = hf.getCellValue(addr(ref)); return (v && typeof v === 'object') ? null : v; },
    };
  }
  root.Simulador = { load };
  if (typeof module !== 'undefined') module.exports = { load };
})(typeof window !== 'undefined' ? window : globalThis);
