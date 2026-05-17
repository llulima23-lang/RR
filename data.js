// data.js — Dados estáticos + parser dinâmico de Excel (SheetJS)

// ============================================================
// DADOS INICIAIS (planilha BASE PARA PPT_ATUAL.xlsx)
// ============================================================
let DADOS = {
  capacity: {
    operacoes: ['ALARES','EDUCACIONAL','M.DIAS BRANCO','AGORACRED'],
    meses: ['Jan','Fev','Mar','Abr'],
    valores: [[7,7,12,17],[5,5,1,1],[1,1,1,1],[3,3,3,3]],
    totais: [16,16,17,22]
  },
  funil: {
    abrilGeral: [
      {op:'AGORACRED',     disc:260960, alo:8346,  alo_pct:0.0320, cpc:1353, cpc_pct:0.1621, prom:274,  prom_pct:0.2025},
      {op:'ALARES VARIÁVEL',disc:1066345,alo:14016, alo_pct:0.0131, cpc:6144, cpc_pct:0.4384, prom:1257, prom_pct:0.2046},
      {op:'ALARES PA FIXA',disc:403852, alo:22191, alo_pct:0.0549, cpc:6166, cpc_pct:0.2779, prom:3040, prom_pct:0.4930},
      {op:'M.DIAS',        disc:118135, alo:1897,  alo_pct:0.0161, cpc:463,  cpc_pct:0.2441, prom:40,   prom_pct:0.0864},
    ],
    maio9Dias: [
      {op:'ALARES VARIÁVEL',disc:608040, alo:5429,  alo_pct:0.0089, cpc:2284, cpc_pct:0.4207, prom:484,  prom_pct:0.2119},
      {op:'ALARES PA FIXA',disc:201566, alo:15458, alo_pct:0.0767, cpc:4536, cpc_pct:0.2934, prom:2034, prom_pct:0.4484},
      {op:'M.DIAS',        disc:56121,  alo:963,   alo_pct:0.0172, cpc:298,  cpc_pct:0.3094, prom:23,   prom_pct:0.0772},
      {op:'AGORACRED',     disc:0, alo:0, alo_pct:0, cpc:0, cpc_pct:0, prom:0, prom_pct:0},
    ],
    abril9Dias: [
      {op:'ALARES VARIÁVEL',disc:627763, alo:5884,  alo_pct:0.0094, cpc:2494, cpc_pct:0.4239, prom:456,  prom_pct:0.1828},
      {op:'ALARES PA FIXA',disc:162912, alo:9385,  alo_pct:0.0576, cpc:2692, cpc_pct:0.2868, prom:1400, prom_pct:0.5201},
      {op:'M.DIAS',        disc:52888,  alo:916,   alo_pct:0.0173, cpc:261,  cpc_pct:0.2849, prom:15,   prom_pct:0.0575},
      {op:'AGORACRED',     disc:0, alo:0, alo_pct:0, cpc:0, cpc_pct:0, prom:0, prom_pct:0},
    ]
  },
  quartil: {
    tarde: {
      quartis:['1º Quartil','2º Quartil','3º Quartil','4º Quartil'],
      dispersao:[1.00,0.50,0.3031,0.17],
      mediapromessas:[494,318,193,109]
    },
    manha: {
      quartis:['1º Quartil','2º Quartil','3º Quartil','4º Quartil'],
      dispersao:[1.00,0.55,0.48,0.3325],
      mediapromessas:[471,283,246,35]
    }
  },
  acoesDigitais: {
    whatsapp: {
      operacoes:['ALARES','EQUATORIAL','M.DIAS','COPEL','ENEL SP','ENEL CE','ENEL RJ'],
      marco:[0,1,42,173,34,45,33], abril:[4,0,17,200,105,42,31], maio:[4,1,13,150,41,3,26]
    },
    chatbot: {
      operacoes:['ENEL SP','ENEL CE','ENEL RJ'],
      marco:[13,60,21], abril:[11,33,24], maio:[17,26,11]
    }
  },
  planoAcao: [
    {oportunidade:'Promessas',          motivo:'Feedbacks e acompanhamento',           responsavel:'Operação'},
    {oportunidade:'Nota Qualidade',      motivo:'Auditoria e feedbacks',                responsavel:'Qualidade'},
    {oportunidade:'Capacity',            motivo:'Operadores em curva de aprendizado',   responsavel:'Operação'},
    {oportunidade:'Divisão PA Fixa/Var.',motivo:'Quartil unificado',                    responsavel:'Operação / INT'},
  ],
  dadosGerais: {
    indicadores: ['ABS (meta 2%)','Qualidade (meta 95%)','Pausa (meta 16%)','Tempo Logado','Turnover','MC'],
    meses: ['Janeiro','Fevereiro','Março','Abril','Maio'],
    valores: [
      [0.0235,0.0128,3.92,5.12,0.0236],
      [99.28,98.67,99.04,0.93,0.9972],
      [0.13,0.14,0.14,0.15,0.15],
      [1,1,1,1,1],
      [0,0,0,0,0],
      [0.706,0.084,-0.43,null,null],
    ],
    metas: [0.02, 0.95, 0.16, null, null, null]
  },
  digital: [],
  mapaDigital: {
    canais: ['SMS','Email','WhatsApp','AGV','URA','Portal','Chatbot'],
    operacoes: [
      { op:'EDUCAÇÃO',    canais:[false,false,true, false,false,false,false] },
      { op:'M DIAS',      canais:[false,false,true, false,false,false,false] },
      { op:'AGORACRED',   canais:[true, false,false,false,false,false,false] },
      { op:'ALARES',      canais:[true, true, true, true, true, false,false] },
      { op:'TIM',         canais:[true, false,false,true, true, false,false] },
      { op:'COPEL',       canais:[true, true, true, true, true, false,false] },
      { op:'EQUATORIAL',  canais:[true, true, true, true, true, false,false] },
      { op:'SABESP',      canais:[true, true, false,false,false,true, false] },
      { op:'ENEL CE',     canais:[true, true, true, true, true, false,true ] },
      { op:'ENEL RJ',     canais:[true, true, true, true, true, false,true ] },
      { op:'ENEL SP',     canais:[true, true, true, true, true, false,true ] },
    ]
  },
  periodo: 'Abril / Maio 2026'
};

// ============================================================
// CONFIGURAÇÃO DO FIREBASE (Sincronização em Nuvem Real-Time)
// ============================================================
const firebaseConfig = {
  apiKey: "AIzaSyDfJc_gjYOBdnGRSw9q71TzBttXAi-07cw",
  authDomain: "rr2026-127c2.firebaseapp.com",
  projectId: "rr2026-127c2",
  storageBucket: "rr2026-127c2.firebasestorage.app",
  messagingSenderId: "306089663556",
  appId: "1:306089663556:web:57d224233e1e06ebe33011",
  databaseURL: "https://rr2026-127c2-default-rtdb.firebaseio.com"
};

// Inicializa o Firebase
if (!firebase.apps.length) {
  firebase.initializeApp(firebaseConfig);
}
const db = firebase.database();

// ============================================================
// UTILITÁRIO: Firebase converte arrays com null em objetos.
// Esta função reconstrói arrays recursivamente.
// ============================================================
function fixFirebaseData(obj) {
  if (obj === null || obj === undefined) return obj;
  if (typeof obj !== 'object') return obj;
  if (!Array.isArray(obj)) {
    const keys = Object.keys(obj);
    const allNumeric = keys.length > 0 && keys.every(k => /^\d+$/.test(k));
    if (allNumeric) {
      const maxIdx = Math.max(...keys.map(Number));
      const arr = [];
      for (let i = 0; i <= maxIdx; i++) {
        arr.push(obj[i] !== undefined ? fixFirebaseData(obj[i]) : null);
      }
      return arr;
    }
  }
  if (Array.isArray(obj)) {
    return obj.map(item => fixFirebaseData(item));
  }
  const result = {};
  for (const key of Object.keys(obj)) {
    result[key] = fixFirebaseData(obj[key]);
  }
  return result;
}

// ============================================================
// FIREBASE — Conexão defensiva (NÃO quebra se Firebase falhar)
// ============================================================
let _firebaseOk = false;
let _isUploading = false; // flag para evitar renderAll duplo durante upload

try {
  db.ref('dashboard/DADOS').on('value', (snapshot) => {
    try {
      const raw = snapshot.val();
      if (!raw) return;
      const cloudData = fixFirebaseData(raw);
      DADOS = cloudData;
      _firebaseOk = true;
      // Não re-renderiza se estamos no meio de um upload (o upload já cuida disso)
      if (!_isUploading && typeof window.renderAll === 'function') {
        window.renderAll();
      }
    } catch(e) {
      console.warn('Firebase: erro ao processar dados da nuvem, usando dados locais.', e);
    }
  });
} catch(e) {
  console.warn('Firebase: não foi possível conectar. Dashboard funcionará apenas localmente.', e);
}

// Quando a página terminar de carregar, se o Firebase já tiver respondido, redesenha.
document.addEventListener('DOMContentLoaded', () => {
  setTimeout(() => {
    if (_firebaseOk && typeof window.renderAll === 'function') {
      window.renderAll();
    }
  }, 800);
});

// ============================================================
// PARSER DE EXCEL (SheetJS) — atualiza DADOS dinamicamente
// ============================================================
async function lerExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(e.target.result, {type:'array'});
        const toArr = (nome) => {
          const target = nome.toUpperCase().trim();
          const sheetName = wb.SheetNames.find(n => n.toUpperCase().trim() === target);
          if (!sheetName) return null;
          return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], {header:1, defval:null});
        };

        const parseNum = (val) => {
          if (val === undefined || val === null || val === '') return 0;
          if (typeof val === 'number') return val;
          const s = String(val).replace('%', '').replace(',', '.').trim();
          const n = Number(s);
          return isNaN(n) ? 0 : n;
        };
        
        const parseNumOrNull = (val) => {
          if (val === undefined || val === null || val === '') return null;
          if (typeof val === 'number') return val;
          const s = String(val).replace('%', '').replace(',', '.').trim();
          const n = Number(s);
          return isNaN(n) ? null : n;
        };

        // CAPACITY
        try {
          const cap = toArr('CAPACITY');
          if (cap) {
            const ops=[], meses=[], vals=[], totais=[];
            const header = cap[0];
            for(let i=1;i<header.length;i++) if(header[i]) meses.push(String(header[i]));
            for(let r=1;r<cap.length;r++){
              const row=cap[r]; if(!row[0]) continue;
              ops.push(String(row[0]));
              const rv=[]; let tot=0;
              for(let i=1;i<=meses.length;i++){ const v=parseNum(row[i]); rv.push(v); tot+=v; }
              vals.push(rv);
            }
            for(let i=0;i<meses.length;i++) totais.push(vals.reduce((s,r)=>s+(r[i]||0),0));
            if(ops.length) DADOS.capacity={operacoes:ops,meses,valores:vals,totais};
          }
        } catch(e) { console.error('Erro CAPACITY', e); }

        // FUNIL2 — formato transposto (métricas nas linhas, operações+períodos nas colunas)
        try {
          const fun2 = toArr('FUNIL2');
          if (fun2 && fun2.length > 1) {
            const headers = fun2[0] || [];

            console.log('=== FUNIL2: Headers encontrados ===');
            headers.forEach((h, i) => console.log(`  Col ${i}: "${h}"`));

            // Encontra em qual linha está cada métrica
            const metricRows = {};
            for (let r = 1; r < fun2.length; r++) {
              const rawLabel = String(fun2[r]?.[0] || '').trim();
              const label = rawLabel.toLowerCase().replace(/[^a-záàãâéêíóôõúç%0-9]/g, '');
              console.log(`  Linha ${r}: "${rawLabel}" → limpo: "${label}"`);
              if (label.includes('discag')) metricRows.disc = r;
              else if (label === 'alô' || label === 'alo' || label === 'alo') metricRows.alo = r;
              else if (label === 'cpc') metricRows.cpc = r;
              else if (label.includes('promessa') && !label.includes('%')) metricRows.prom = r;
              else if (label.includes('%al') || label.includes('%alo')) metricRows.alo_pct = r;
              else if (label.includes('%cpc')) metricRows.cpc_pct = r;
              else if (label.includes('%prom')) metricRows.prom_pct = r;
            }
            console.log('=== FUNIL2: Mapeamento de linhas ===', JSON.stringify(metricRows));

            const abrilGeral = [], maio9Dias = [], abril9Dias = [];

            for (let c = 1; c < headers.length; c++) {
              const h = String(headers[c] || '').toUpperCase().trim();
              if (!h) continue;

              // Identifica a operação (ordem importa: mais específico primeiro)
              let opName = '';
              if (h.includes('AGORACRED')) opName = 'AGORACRED';
              else if (h.includes('VARIAVEL') || h.includes('VARIÁVEL')) opName = 'ALARES VARIÁVEL';
              else if (h.includes('FIXA')) opName = 'ALARES PA FIXA';
              else if (h.includes('M.DIAS') || h.includes('MDIAS') || h.includes('M DIAS') || h.includes('M.DIA')) opName = 'M.DIAS';
              else {
                console.warn(`  FUNIL2 Col ${c}: header "${h}" não reconhecido, pulando.`);
                continue;
              }

              // Identifica o período
              let period = null;
              const has9 = h.includes('9 DIAS') || h.includes('9DIAS');
              if (has9) {
                if (h.includes('MAIO')) period = 'maio9Dias';
                else period = 'abril9Dias';
              } else {
                period = 'abrilGeral';
              }

              console.log(`  FUNIL2 Col ${c}: "${h}" → op="${opName}", period="${period}"`);

              // Lê valores da coluna
              const getVal = (rowIdx) => rowIdx ? parseNum(fun2[rowIdx]?.[c]) : 0;
              const getPct = (rowIdx) => {
                if (!rowIdx) return 0;
                let v = parseNum(fun2[rowIdx]?.[c]);
                if (Math.abs(v) > 1) v = v / 100; // 16% → 0.16
                return v;
              };

              const entry = {
                op: opName,
                disc: getVal(metricRows.disc),
                alo: getVal(metricRows.alo),
                alo_pct: getPct(metricRows.alo_pct),
                cpc: getVal(metricRows.cpc),
                cpc_pct: getPct(metricRows.cpc_pct),
                prom: getVal(metricRows.prom),
                prom_pct: getPct(metricRows.prom_pct),
              };

              console.log(`    → disc=${entry.disc}, alo=${entry.alo}, cpc=${entry.cpc}, prom=${entry.prom}`);

              if (period === 'abrilGeral') abrilGeral.push(entry);
              else if (period === 'maio9Dias') maio9Dias.push(entry);
              else if (period === 'abril9Dias') abril9Dias.push(entry);
            }

            console.log('=== FUNIL2 RESULTADO ===');
            console.log('  abrilGeral:', abrilGeral.length, 'ops:', abrilGeral.map(e=>e.op));
            console.log('  maio9Dias:', maio9Dias.length, 'ops:', maio9Dias.map(e=>e.op));
            console.log('  abril9Dias:', abril9Dias.length, 'ops:', abril9Dias.map(e=>e.op));

            if (abrilGeral.length) DADOS.funil.abrilGeral = abrilGeral;
            if (maio9Dias.length) DADOS.funil.maio9Dias = maio9Dias;
            if (abril9Dias.length) DADOS.funil.abril9Dias = abril9Dias;
          }
        } catch(e) { console.error('Erro FUNIL2', e); }

        // QUARTIL2026
        try {
          const qraw = toArr('QUARTIL2026');
          if (qraw) {
            const parseQ=(rows,startRow)=>{
              const q={quartis:[],dispersao:[],mediapromessas:[]};
              const hr=rows[startRow]||[];
              for(let c=1;c<hr.length;c++) if(hr[c]) q.quartis.push(String(hr[c]));
              for(let r=startRow+1;r<startRow+5&&r<rows.length;r++){
                const row=rows[r]; if(!row) continue;
                const lbl=String(row[0]||'').toLowerCase();
                if(lbl.includes('dispers')) for(let c=1;c<=q.quartis.length;c++) q.dispersao.push(parseNum(row[c]));
                if(lbl.includes('prom')) for(let c=1;c<=q.quartis.length;c++) q.mediapromessas.push(parseNum(row[c]));
              }
              return q;
            };
            let tardeRow=0,manhaRow=0;
            for(let r=0;r<qraw.length;r++){
              const c=String(qraw[r]?.[0]||'').toLowerCase();
              if(c.includes('tarde')) tardeRow=r;
              if(c.includes('manh')) manhaRow=r;
            }
            const qt=parseQ(qraw,tardeRow); if(qt.quartis.length) DADOS.quartil.tarde=qt;
            const qm=parseQ(qraw,manhaRow); if(qm.quartis.length) DADOS.quartil.manha=qm;
          }
        } catch(e) { console.error('Erro QUARTIL', e); }

        // PLANO DE AÇÃO
        try {
          const praw = toArr('PLANO DE AÇÃO');
          if(praw){
            const pl=[];
            for(let r=1;r<praw.length;r++){
              const row=praw[r]; if(!row||!row[0]) continue;
              pl.push({oportunidade:String(row[0]),motivo:String(row[1]||''),responsavel:String(row[2]||'')});
            }
            if(pl.length) DADOS.planoAcao=pl;
          }
        } catch(e) { console.error('Erro PLANO', e); }

        // AÇÕES DIGITAIS
        try {
          const adraw = toArr('AÇÕES DIGITAIS');
          if(adraw && adraw.length>2){
            const header=adraw[0]||[];
            const ops=[], mar=[], abr=[], mai=[];
            for(let r=1;r<adraw.length;r++){
              const row=adraw[r]; if(!row||!row[0]) continue;
              ops.push(String(row[0]));
              mar.push(parseNum(row[1])); abr.push(parseNum(row[2])); mai.push(parseNum(row[3]));
            }
            if(ops.length) DADOS.acoesDigitais.whatsapp={operacoes:ops,marco:mar,abril:abr,maio:mai};
          }
        } catch(e) { console.error('Erro DIGITAL', e); }

        // DADOS GERAIS
        try {
          const dgraw = toArr('DADOS GERAIS');
          if(dgraw && dgraw.length>1){
            const header=dgraw[0]||[];
            const meses=[]; for(let c=1;c<header.length;c++) if(header[c]) meses.push(String(header[c]));
            const inds=[], vals=[], metas=[];
            for(let r=1;r<dgraw.length;r++){
              const row=dgraw[r]; if(!row||!row[0]) continue;
              inds.push(String(row[0]));
              const rv=[]; for(let c=1;c<=meses.length;c++) rv.push(parseNumOrNull(row[c]));
              vals.push(rv);
              // extrai meta do nome se houver (ex: "ABS- 2%" ou "meta 10")
              let metaVal = null;
              const text = String(row[0]);
              // Tenta achar "meta X%" ou "X%"
              let m = text.match(/(-?\d+(?:[\.,]\d+)?)\s*%/);
              if (m) {
                metaVal = parseNum(m[1]) / 100;
              } else {
                // Tenta achar "meta X"
                m = text.match(/meta\s*(-?\d+(?:[\.,]\d+)?)/i);
                if (m) metaVal = parseNum(m[1]);
              }
              metas.push(metaVal);
            }
            if(inds.length) DADOS.dadosGerais={indicadores:inds,meses,valores:vals,metas};
          }
        } catch(e) { console.error('Erro DADOS GERAIS', e); }

        // DIGITAL
        try {
          const digiraw = toArr('DIGITAL');
          if(digiraw) {
            const rows=[]; 
            for(const r of digiraw) if(r&&r.some(v=>v!==null&&v!=='')) rows.push(r);
            DADOS.digital=rows;
          }
        } catch(e) { console.error('Erro DIGITAL ABA', e); }

        // Salvar no Firebase para manter os dados no refresh e propagar para todos
        try {
          db.ref('dashboard/DADOS').set(DADOS);
          localStorage.setItem('MF_DADOS_2026', JSON.stringify(DADOS)); // fallback opcional
        } catch (e) {
          console.error("Erro ao salvar no Firebase/localStorage", e);
        }

        resolve(true);
      } catch(err) { reject(err); }
    };
    reader.readAsArrayBuffer(file);
  });
}
