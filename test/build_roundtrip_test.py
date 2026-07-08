#!/usr/bin/env python3
# Extract the REAL functions from the userscript and assemble a jsc test program
# that runs a full export -> simulated user edits -> import round trip.
import re, sys, pathlib

HERE = pathlib.Path(__file__).resolve().parent
SRC = (HERE.parent / "coupa-receipt-filter.user.js").read_text()

def slice_fn(name):
    # match '  function NAME(' or '  async function NAME('
    m = re.search(r"\n  (?:async )?function " + re.escape(name) + r"\(", SRC)
    if not m:
        sys.exit("FN NOT FOUND: " + name)
    start = m.start() + 1  # drop leading newline
    # find terminating line that is exactly '  }' after start
    idx = SRC.index("\n  }\n", m.end())
    end = idx + len("\n  }")  # include the closing brace line
    return SRC[start:end]

def slice_between(start_pat, end_pat):
    m = re.search(start_pat, SRC)
    if not m:
        sys.exit("PAT NOT FOUND: " + start_pat)
    start = m.start()
    idx = SRC.index(end_pat, m.end())
    end = idx + len(end_pat)
    return SRC[start:end].strip()

fns = ["parseAmt", "lineProblems", "collectAttendeeDirectory",
       "collectCurrencyDirectory", "collectUsedCategories",
       "_buildWorkbookInner", "parseUploadedWorkbook",
       "buildLineUpdatePatchBody"]
extracted = "\n\n".join(slice_fn(n) for n in fns)

all_categories = slice_between(r"\n  const ALL_CATEGORIES = \[", "\n  ];")
example_hdr = slice_between(r"\n  const EXAMPLE_ATTENDEE_HEADER = ", ";").strip()

harness = r"""
'use strict';
// ---- constants copied/sliced from source ----
const NEW_ATTENDEE_TYPE_ID = 6;
const GIFT_MEAL_CATEGORY_ID = 85;
const PROBLEM_USD_THRESHOLD = 25;
%(EXAMPLE_HDR)s;
%(ALL_CATEGORIES)s;

// ---- global stubs jsc lacks ----
globalThis.Blob = function(parts){ this.parts = parts; };
async function loadExcelJS(){ return MockExcelJS; }
if (typeof URLSearchParams === 'undefined') {
  globalThis.URLSearchParams = class {
    constructor(){ this._p = []; }
    append(k, v){ this._p.push([k, String(v)]); }
    toString(){ return this._p.map(([k,v])=> encodeURIComponent(k)+'='+encodeURIComponent(v)).join('&'); }
  };
}
// parse an application/x-www-form-urlencoded body into {uniqueKey:value, 'k[]':[...]}
function parseBody(s){
  const out = {}; const multi = {};
  s.split('&').forEach(pair=>{ if(!pair) return; const i=pair.indexOf('='); const k=decodeURIComponent(pair.slice(0,i)); const v=decodeURIComponent(pair.slice(i+1));
    if(k.endsWith('[]')){ (multi[k]=multi[k]||[]).push(v); } else { out[k]=v; } });
  Object.assign(out, multi); return out;
}

// ---- minimal in-memory ExcelJS mock (subset used by the two functions) ----
function colToNum(letters){ let n=0; for (const ch of letters){ n = n*26 + (ch.charCodeAt(0)-64); } return n; }
class Cell { constructor(v){ this.value = (v===undefined? null : v); } }
class Row {
  constructor(arr){ this.cells = []; (arr||[]).forEach((v,i)=>{ this.cells[i] = new Cell(v); }); this._font=null; }
  getCell(n){ while(this.cells.length < n) this.cells.push(new Cell(null)); if(!this.cells[n-1]) this.cells[n-1]=new Cell(null); return this.cells[n-1]; }
  get values(){ const out=[undefined]; for(const c of this.cells) out.push(c? c.value : undefined); return out; }
  set font(v){ this._font=v; } get font(){ return this._font; }
  eachCell(opts, cb){ this.cells.forEach((c,i)=>{ if(opts&&opts.includeEmpty===false){ if(c==null||c.value==null||c.value==='') return; } cb(c, i+1); }); }
}
class Column { constructor(){ this.width=undefined; } }
class Sheet {
  constructor(name, opts){ this.name=name; this.opts=opts||{}; this.rows=[]; this._views=[]; this._cols={}; this._cf=[]; }
  addRow(arr){ const r=new Row(arr); this.rows.push(r); return r; }
  getRow(n){ while(this.rows.length < n) this.rows.push(new Row([])); return this.rows[n-1]; }
  getColumn(n){ if(!this._cols[n]) this._cols[n]=new Column(); return this._cols[n]; }
  getCell(ref){ const m=/^([A-Z]+)(\d+)$/.exec(ref); if(!m) throw new Error('bad ref '+ref); const col=colToNum(m[1]); const rn=parseInt(m[2],10); return this.getRow(rn).getCell(col); }
  addConditionalFormatting(o){ this._cf.push(o); }
  set views(v){ this._views=v; } get views(){ return this._views; }
  eachRow(opts, cb){ this.rows.forEach((r,i)=>{ const idx=i+1; if(opts&&opts.includeEmpty===false){ const any=r.cells.some(c=>c&&c.value!=null&&c.value!==''); if(!any) return; } cb(r, idx); }); }
}
class Workbook {
  constructor(){ this._sheets=[]; this._byName={}; const self=this; this.xlsx={ writeBuffer: async()=>self._serialize(), load: async(buf)=>self._deserialize(buf) }; }
  addWorksheet(name, opts){ const ws=new Sheet(name,opts); this._sheets.push(ws); this._byName[name]=ws; return ws; }
  getWorksheet(name){ return this._byName[name]||null; }
  _serialize(){ return { sheets: this._sheets.map(s=>({ name:s.name, rows: s.rows.map(r=> r.cells.map(c=> c? c.value : null)) })) }; }
  _deserialize(buf){ buf.sheets.forEach(sd=>{ const ws=this.addWorksheet(sd.name); sd.rows.forEach(cells=> ws.addRow(cells)); }); }
}
const MockExcelJS = { Workbook };

%(EXTRACTED)s

// ================= TEST DRIVER =================
const usdRates = { EUR: 0.9, PLN: 4.0 }; // USD implicit = 1
function mkAtt(id, first, last){ return { id, expense_attendee_type_id: 6, first_name: first, last_name: last }; }
function mkLine(o){ return Object.assign({ merchant:'M', local_expense_date:'2026-01-01', receipt_amount:'100.00', expense_category_id:59, expense_category_name:'Meals', description:'d', expense_artifacts:[{}], accounts:[{account_id:1,account_type_id:4}], expense_attendees:[] }, o); }

const reports = [{
  title:'Rpt1',
  expense_lines: [
    mkLine({ id: 1001, receipt_amount:'50.00', receipt_currency:{id:1, code:'USD', decimals:2}, expense_attendees:[ mkAtt(500,'Ada','Lovelace') ] }),
    mkLine({ id: 1002, receipt_amount:'80.00', receipt_currency:{id:8, code:'EUR', decimals:2} }),
    mkLine({ id: 1003, receipt_amount:'20.00', receipt_currency:{id:20, code:'PLN', decimals:2}, amount_to_receive_currency:{id:1, code:'USD'} }),
  ],
}];

const assert = (cond, msg) => { if(!cond){ print('ASSERT FAIL: '+msg); globalThis.__fail=(globalThis.__fail||0)+1; } else { print('ok: '+msg); } };

function findHeaderRow(bufSheet){ return bufSheet.rows[0]; }
function colOf(headerRow, name){ return headerRow.indexOf(name); } // 0-based

(async () => {
  // ---- EXPORT (all lines) ----
  const status = {};
  const { blob } = await _buildWorkbookInner(MockExcelJS, reports, usdRates, status, {}, false);
  const buf0 = blob.parts[0];
  const linesBuf = buf0.sheets.find(s=>s.name==='Lines');
  const curBuf = buf0.sheets.find(s=>s.name==='_currencies');
  const hdr = linesBuf.rows[0];
  print('HEADERS: ' + JSON.stringify(hdr));

  // header layout assertions (1-based column numbers)
  assert(colOf(hdr,'new_currency')===6, "new_currency is column G (7th, 0-based 6)");
  assert(colOf(hdr,'currency')===5, "currency is column F");
  assert(colOf(hdr,'usd_eq')===7, "usd_eq shifted to column H");
  assert(colOf(hdr,'new_category')===9, "new_category is column J (0-based 9)");
  assert(colOf(hdr,'current_attendees')===13, "current_attendees is column N");
  assert(hdr[hdr.length-1]===EXAMPLE_ATTENDEE_HEADER, "last column is the example attendee header");
  const exampleColIdx0 = hdr.length-1;

  // _currencies helper
  assert(!!curBuf, "_currencies helper sheet exists");
  const curMap = {}; curBuf.rows.slice(1).forEach(r=> curMap[r[0]]=r[1]);
  print('CURRENCY MAP: '+JSON.stringify(curMap));
  assert(curMap.USD===1 && curMap.EUR===8 && curMap.PLN===20, "currency code->id map correct (USD1/EUR8/PLN20)");

  // ---- helper to deep-clone the buffer ----
  const clone = (b)=> JSON.parse(JSON.stringify(b));

  // ================= RUN 1: currency change + rename example -> new attendee =================
  {
    const b = clone(buf0);
    const L = b.sheets.find(s=>s.name==='Lines');
    const h = L.rows[0];
    const ncur = colOf(h,'new_currency');
    // line 1001 (row index 1): change USD -> EUR
    L.rows[1][ncur] = 'eur'; // lower-case to test normalization
    // rename example column header to a real name and mark x on line 1001
    h[exampleColIdx0] = 'Grace Hopper';
    L.rows[1][exampleColIdx0] = 'x';
    const file = { arrayBuffer: async()=> b };
    let res;
    try { res = await parseUploadedWorkbook(file); }
    catch(e){ print('RUN1 unexpected throw: '+ (e && e.message)); print(JSON.stringify(e.validationErrors||[])); globalThis.__fail=(globalThis.__fail||0)+1; return; }
    const ch1001 = res.changes.find(c=>c.line_id===1001);
    assert(!!ch1001, "RUN1: change recorded for line 1001");
    assert(ch1001 && ch1001.new_currency_id===8 && ch1001.new_currency_code==='EUR', "RUN1: currency 'eur' -> id 8, code EUR");
    const grace = res.attendees.find(a=> a.first_name==='Grace' && a.last_name==='Hopper');
    assert(!!grace && grace.id===null, "RUN1: new attendee 'Grace Hopper' synthesized with null id (to be created)");
    assert(ch1001 && ch1001.attendees.some(a=>a===grace), "RUN1: line 1001 marks the new Grace Hopper attendee");
    // existing attendee Ada (id 500) should also be marked (was pre-checked 'x')
    assert(ch1001 && ch1001.attendees.some(a=>a.id===500), "RUN1: pre-existing attendee Ada (500) still marked");
  }

  // ================= RUN 2: unknown currency -> validation error =================
  {
    const b = clone(buf0);
    const L = b.sheets.find(s=>s.name==='Lines');
    const h = L.rows[0];
    const ncur = colOf(h,'new_currency');
    L.rows[2][ncur] = 'XYZ'; // line 1002, bogus currency
    const file = { arrayBuffer: async()=> b };
    let threw=false, msg='';
    try { await parseUploadedWorkbook(file); }
    catch(e){ threw=true; msg=(e&&e.message)||''; }
    assert(threw && /new_currency "XYZ"/.test(msg), "RUN2: unknown currency 'XYZ' raises a validation error");
  }

  // ================= RUN 3: unmodified example column is ignored (even with a stray value) =================
  {
    const b = clone(buf0);
    const L = b.sheets.find(s=>s.name==='Lines');
    const h = L.rows[0];
    // header left as the sentinel; put a stray non-x value in an example cell
    L.rows[1][exampleColIdx0] = 'y';
    const file = { arrayBuffer: async()=> b };
    let res, threw=false, msg='';
    try { res = await parseUploadedWorkbook(file); }
    catch(e){ threw=true; msg=(e&&e.message)||''; }
    assert(!threw, "RUN3: stray value under unmodified example header does NOT error (msg="+msg+")");
    if(res){
      const anyGraceLike = res.attendees.some(a=> a.id===null);
      assert(!anyGraceLike, "RUN3: no phantom attendee created from the example column");
    }
  }

  // ================= RUN 4: no edits -> no currency/category/desc changes (attendee-only no-ops handled downstream) =================
  {
    const file = { arrayBuffer: async()=> clone(buf0) };
    const res = await parseUploadedWorkbook(file);
    const withCur = res.changes.filter(c=> c.new_currency_id!=null);
    const withCat = res.changes.filter(c=> c.new_category_id!=null);
    const withDesc = res.changes.filter(c=> c.new_description!=null);
    assert(withCur.length===0, "RUN4: no new_currency changes when nothing edited");
    assert(withCat.length===0 && withDesc.length===0, "RUN4: no category/description changes when nothing edited");
  }

  // ================= RUN 5: buildLineUpdatePatchBody — single-currency line, NO foreign fields =================
  {
    const line = { id:1001, receipt_amount:'50.00', receipt_currency:{id:1,code:'USD'}, expense_attendees:[] };
    const change = { line_id:1001, new_currency_id:8, new_currency_code:'EUR', attendees:[] };
    const body = parseBody(buildLineUpdatePatchBody(line, change));
    assert(body['expense_line[receipt_total_currency_id]']==='8', "RUN5: receipt_total_currency_id set to new id 8");
    assert(body['expense_line[receipt_total_amount]']==='50.00', "RUN5: receipt amount unchanged (50.00)");
    assert(body['expense_line[foreign_currency_id]']==='', "RUN5: foreign currency stays empty (no reimbursement tracking)");
    assert(body['expense_line[exchange_rate]']==='', "RUN5: exchange_rate stays empty");
    assert(body['expense_line[foreign_currency_amount]']==='', "RUN5: foreign amount stays empty");
    assert(body['expense_line[amount_to_receive]']==='', "RUN5: amount_to_receive stays empty");
  }

  // ================= RUN 6: buildLineUpdatePatchBody — single-currency line WITH foreign==receipt =================
  {
    const line = { id:1002, receipt_amount:'50.00', receipt_currency:{id:1,code:'USD'},
      amount_to_receive_currency:{id:1,code:'USD'}, amount_to_receive:'50.00', exchange_rate:'1', expense_attendees:[] };
    const change = { line_id:1002, new_currency_id:8, new_currency_code:'EUR', attendees:[] };
    const body = parseBody(buildLineUpdatePatchBody(line, change));
    assert(body['expense_line[receipt_total_currency_id]']==='8', "RUN6: receipt currency -> 8");
    assert(body['expense_line[foreign_currency_id]']==='8', "RUN6: reimbursement currency MOVED to 8 (stays single-currency, not stale USD)");
    assert(body['expense_line[exchange_rate]']==='1', "RUN6: exchange_rate forced to 1 (no stale rate)");
    assert(body['expense_line[foreign_currency_amount]']==='50.00', "RUN6: foreign amount == receipt amount 50.00");
    assert(body['expense_line[amount_to_receive]']==='50.00', "RUN6: amount_to_receive == receipt amount 50.00");
  }

  // ================= RUN 7: no currency change -> original values preserved verbatim =================
  {
    const line = { id:1003, receipt_amount:'50.00', receipt_currency:{id:1,code:'USD'},
      amount_to_receive_currency:{id:1,code:'USD'}, amount_to_receive:'50.00', exchange_rate:'1', expense_attendees:[], description:'old' };
    const change = { line_id:1003, new_description:'new desc', attendees:[] }; // no currency
    const body = parseBody(buildLineUpdatePatchBody(line, change));
    assert(body['expense_line[receipt_total_currency_id]']==='1', "RUN7: no currency change -> keeps id 1");
    assert(body['expense_line[foreign_currency_id]']==='1', "RUN7: foreign currency untouched (1)");
    assert(body['expense_line[exchange_rate]']==='1', "RUN7: exchange_rate untouched (1)");
    assert(body['expense_line[description]']==='new desc', "RUN7: description change still applied");
  }

  // ================= RUN 8: isCrossCurrency classification (ported guard, mirrors applyUpload) =================
  {
    const isCrossCurrency = (line) => {
      const fId = line.amount_to_receive_currency?.id;
      const rId = line.receipt_currency?.id;
      if (fId != null && fId !== '' && Number(fId) !== Number(rId)) return true;
      const rate = line.exchange_rate;
      if (rate != null && rate !== '' && Number(rate) !== 1) return true;
      return false;
    };
    assert(isCrossCurrency({ receipt_currency:{id:1} })===false, "RUN8: single-currency (no foreign) -> not cross");
    assert(isCrossCurrency({ receipt_currency:{id:1}, amount_to_receive_currency:{id:1}, exchange_rate:'1' })===false, "RUN8: foreign==receipt, rate 1 -> not cross");
    assert(isCrossCurrency({ receipt_currency:{id:8}, amount_to_receive_currency:{id:1} })===true, "RUN8: EUR receipt / USD reimbursement -> CROSS (refused)");
    assert(isCrossCurrency({ receipt_currency:{id:1}, amount_to_receive_currency:{id:1}, exchange_rate:'1.10' })===true, "RUN8: rate 1.10 -> CROSS (refused)");
  }

  print(globalThis.__fail ? ('\n==== FAILURES: '+globalThis.__fail+' ====') : '\n==== ALL ASSERTIONS PASSED ====');
})();
"""

out = harness % {
    "EXAMPLE_HDR": example_hdr,
    "ALL_CATEGORIES": all_categories.rstrip(),
    "EXTRACTED": extracted,
}
pathlib.Path(HERE / "roundtrip_test.generated.js").write_text(out)
print("wrote roundtrip_test.generated.js (%d chars)" % len(out))
