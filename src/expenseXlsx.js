/* 지출증빙 및 물품 검수 보고 — 엑셀 생성 (브라우저에서 바로 처리)
   public/expense/ 의 빈 양식(원본 엑셀에서 사진·샘플값만 지운 것)을 열어
   XML을 직접 고치기 때문에 원본 서식·결재란 그림이 그대로 유지됩니다. */
import JSZip from "jszip";

const EMU_PER_PX = 9525, EMU_PER_PT = 12700;

/* 양식별 위치 정보 */
export const LAYOUTS = {
  approval: { // 1번: 결재란 있음
    file: "/expense/tpl_approval.xlsx", sheet: "검수서", draftSheet: "기안사본",
    itemStart: 8, itemEnd: 11, total: 12, photoRows: 25,
    cols: { no: "B", date: "C", name: "D", qty: "E", amt: "F", amtEnd: "H", note: "I" },
    vMerge: ["date", "note"], // 품목이 여러 줄이면 세로 병합
    left: [1, 3], right: [4, 8], // 사진 칸 (0부터: B~D / E~I)
    draftBox: { c0: 1, c1: 26, r0: 1, r1: 24 }, // 기안사본 B1:AA24
    header: (d) => ({
      F4: "                         ▣ 납품업체 :   " + d.vendor,
      F5: "                         ▣ 검수일자 :   " + d.inspectDate,
      F6: "                         ▣ 검 수 자 :   " + spaced(d.inspector) + "    (인)",
      B6: "문서번호 : " + d.docNo,
    }),
  },
  simple: { // 2번: 결재란 없음
    file: "/expense/tpl_simple.xlsx", sheet: "검수서",
    itemStart: 9, itemEnd: 9, total: 10, photoRows: 28,
    cols: { no: "A", date: "B", name: "C", qty: "D", amt: "E", amtEnd: "F", note: "G" },
    vMerge: ["date", "note"],
    left: [0, 2], right: [3, 6], // A~C / D~G
    header: (d) => ({
      F5: d.vendor,
      F6: { date: d.inspectDate },
      F7: spaced(d.inspector),
      A7: "문서번호 : " + d.docNo,
    }),
  },
};

const spaced = (n) => { n = (n || "").replace(/\s/g, ""); return n.length <= 3 ? n.split("").join(" ") : n; };
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const colNum = (L) => [...L].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0); // A=1
const colLet = (n) => { let s = ""; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
const serial = (iso) => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };

function cellXml(ref, s, v) {
  const sa = s != null ? ` s="${s}"` : "";
  if (v == null || v === "") return `<c r="${ref}"${sa}/>`;
  if (typeof v === "object" && v.date) return `<c r="${ref}"${sa}><v>${serial(v.date)}</v></c>`;
  if (typeof v === "object" && v.f) return `<c r="${ref}"${sa}><f>${v.f}</f><v>${v.v ?? 0}</v></c>`;
  if (typeof v === "number") return `<c r="${ref}"${sa}><v>${v}</v></c>`;
  return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`;
}

/* 시트 XML 파서 (행 단위) */
function parseRows(xml) {
  const sd0 = xml.indexOf("<sheetData"), sd1 = xml.indexOf("</sheetData>");
  const open = xml.slice(sd0, xml.indexOf(">", sd0) + 1);
  const body = xml.slice(sd0 + open.length, sd1);
  const rows = new Map();
  for (const m of body.matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const r = +/\br="(\d+)"/.exec(m[1])[1];
    const cells = new Map();
    for (const c of (m[2] || "").matchAll(/<c\b[^>]*?(?:\/>|>[\s\S]*?<\/c>)/g)) {
      const ref = /\br="([A-Z]+)\d+"/.exec(c[0])[1];
      cells.set(ref, c[0]);
    }
    rows.set(r, { attrs: m[1].replace(/\s*\br="\d+"/, "").replace(/\s*\bspans="[^"]*"/, ""), cells });
  }
  return { head: xml.slice(0, sd0) + open, tail: xml.slice(sd1), rows };
}
function buildRows({ head, tail, rows }) {
  const out = [...rows.keys()].sort((a, b) => a - b).map((r) => {
    const { attrs, cells } = rows.get(r);
    const cs = [...cells.keys()].sort((a, b) => colNum(a) - colNum(b)).map((k) => cells.get(k).replace(/\br="[A-Z]+\d+"/, `r="${k}${r}"`));
    const a = /\bht="/.test(attrs) && !/customHeight=/.test(attrs) ? attrs + ' customHeight="1"' : attrs; // 행 높이 고정 (사진 칸 크기 유지)
    return `<row r="${r}"${a}>${cs.join("")}</row>`;
  });
  return head + out.join("") + tail;
}
const styleOf = (cx) => { const m = cx && /\bs="(\d+)"/.exec(cx); return m ? m[1] : null; };
function setCell(rows, ref, v, keepStyle = true, style) {
  const [, L, R] = /^([A-Z]+)(\d+)$/.exec(ref);
  const row = rows.get(+R) || { attrs: "", cells: new Map() };
  rows.set(+R, row);
  const s = style ?? (keepStyle ? styleOf(row.cells.get(L)) : null);
  row.cells.set(L, cellXml(ref, s, v));
}

/* 열 너비·행 높이 → EMU */
function geometry(xml, rowsMax) {
  const dw = +(/defaultColWidth="([\d.]+)"/.exec(xml)?.[1] || 8.43);
  const dh = +(/defaultRowHeight="([\d.]+)"/.exec(xml)?.[1] || 15);
  const colW = (i) => { // i: 1부터
    for (const m of xml.matchAll(/<col min="(\d+)" max="(\d+)"[^>]*?width="([\d.]+)"([^>]*)\/>/g))
      if (i >= +m[1] && i <= +m[2]) return /hidden="1"/.test(m[4]) ? 0 : +m[3];
    return dw;
  };
  const px = (w) => Math.trunc(((256 * w + Math.trunc(128 / 7)) / 256) * 7);
  const cols = Array.from({ length: 60 }, (_, i) => px(colW(i + 1)) * EMU_PER_PX);
  const rowH = new Map();
  for (const m of xml.matchAll(/<row r="(\d+)"[^>]*?\bht="([\d.]+)"/g)) rowH.set(+m[1], +m[2]);
  const rows = Array.from({ length: rowsMax + 5 }, (_, i) => (rowH.get(i + 1) ?? dh) * EMU_PER_PT);
  const start = (arr, i) => arr.slice(0, i).reduce((a, b) => a + b, 0);
  const locate = (arr, emu) => { let i = 0; while (i < arr.length - 1 && emu >= arr[i]) { emu -= arr[i]; i++; } return [i, Math.round(emu)]; };
  return { cols, rows, start, locate };
}

/* 사진 칸 안에 여러 장을 세로로 나눠 배치 */
function placeImages(geo, imgs, c0, c1, r0, r1, pad = 70000) {
  const x0 = geo.start(geo.cols, c0), x1 = geo.start(geo.cols, c1 + 1);
  const y0 = geo.start(geo.rows, r0 - 1), y1 = geo.start(geo.rows, r1);
  const W = x1 - x0 - pad * 2, slotH = (y1 - y0 - pad * (imgs.length + 1)) / imgs.length;
  return imgs.map((im, k) => {
    const sc = Math.min(W / im.w, slotH / im.h);
    const cx = Math.round(im.w * sc), cy = Math.round(im.h * sc);
    const x = x0 + pad + (W - cx) / 2, y = y0 + pad + k * (slotH + pad) + (slotH - cy) / 2;
    return { im, from: [...geo.locate(geo.cols, x), ...geo.locate(geo.rows, y)], to: [...geo.locate(geo.cols, x + cx), ...geo.locate(geo.rows, y + cy)], cx, cy };
  });
}

function anchorXml(p, id, rid) {
  const pt = ([c, co, r, ro]) => `<xdr:col>${c}</xdr:col><xdr:colOff>${co}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${ro}</xdr:rowOff>`;
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${pt(p.from)}</xdr:from><xdr:to>${pt(p.to)}</xdr:to><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="사진 ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>` +
    `<xdr:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>` +
    `<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${p.cx}" cy="${p.cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:twoCellAnchor>`;
}

const relsOf = (p) => p.replace(/([^/]+)$/, "_rels/$1.rels");
const resolve = (base, t) => { const parts = base.split("/").slice(0, -1); for (const s of t.split("/")) s === ".." ? parts.pop() : parts.push(s); return parts.join("/"); };

async function sheetPath(zip, name) {
  const wb = await zip.file("xl/workbook.xml").async("string");
  const rid = new RegExp(`<sheet [^>]*name="${name}"[^>]*r:id="([^"]+)"`).exec(wb)[1];
  const rels = await zip.file("xl/_rels/workbook.xml.rels").async("string");
  const t = new RegExp(`Id="${rid}"[^>]*Target="([^"]+)"`).exec(rels)?.[1] || new RegExp(`Target="([^"]+)"[^>]*Id="${rid}"`).exec(rels)[1];
  return t.startsWith("/") ? t.slice(1) : resolve("xl/workbook.xml", t);
}

let imgSeq = 0;
async function addPictures(zip, sheetFile, placements) {
  if (!placements.length) return;
  const srels = await zip.file(relsOf(sheetFile)).async("string");
  const dt = /Target="([^"]*drawings\/drawing\d+\.xml)"/.exec(srels)[1];
  const dPath = resolve(sheetFile, dt);
  let dx = await zip.file(dPath).async("string");
  const drPath = relsOf(dPath);
  let dr = zip.file(drPath) ? await zip.file(drPath).async("string")
    : `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  let anchors = "";
  for (const p of placements) {
    imgSeq++;
    const name = `exp${Date.now()}_${imgSeq}.jpeg`, rid = `rIdExp${imgSeq}`;
    zip.file(`xl/media/${name}`, p.im.bytes);
    dr = dr.replace("</Relationships>", `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/${name}"/></Relationships>`);
    anchors += anchorXml(p, 2000 + imgSeq, rid);
  }
  dx = dx.replace(/<\/xdr:wsDr>\s*$/, anchors + "</xdr:wsDr>");
  zip.file(dPath, dx); zip.file(drPath, dr);
}

/* 사진 파일 → 회전 보정 + 축소한 JPEG */
export async function prepImage(file, max = 1800) {
  let src;
  try { src = await createImageBitmap(file, { imageOrientation: "from-image" }); }
  catch {
    src = await new Promise((ok, no) => { const i = new Image(); i.onload = () => ok(i); i.onerror = () => no(new Error(`${file.name}: 이미지를 열 수 없어요 (JPG/PNG로 올려 주세요)`)); i.src = URL.createObjectURL(file); });
  }
  const w0 = src.width, h0 = src.height, k = Math.min(1, max / Math.max(w0, h0));
  const w = Math.round(w0 * k), h = Math.round(h0 * k);
  const cv = document.createElement("canvas"); cv.width = w; cv.height = h;
  const g = cv.getContext("2d"); g.fillStyle = "#fff"; g.fillRect(0, 0, w, h); g.drawImage(src, 0, 0, w, h);
  const blob = await new Promise((ok) => cv.toBlob(ok, "image/jpeg", 0.85));
  return { bytes: new Uint8Array(await blob.arrayBuffer()), w, h };
}

/* d = { layout, vendor, purchaseDate, inspectDate, inspector, docNo, note,
         items:[{name, qty, amount}], photos:[img], receipts:[img], drafts:[img] } */
export async function buildExpenseXlsx(d) {
  const L = LAYOUTS[d.layout];
  const zip = await JSZip.loadAsync(await (await fetch(L.file)).arrayBuffer());
  const sp = await sheetPath(zip, L.sheet);
  let xml = await zip.file(sp).async("string");
  const P = parseRows(xml);
  const { cols: C } = L;
  const n = Math.max(1, d.items.length), tplN = L.itemEnd - L.itemStart + 1, delta = n - tplN;

  // 1) 품목 행 스타일(첫/중간/마지막) 확보 후 품목 영역 비우기
  const styleRow = (r) => { const row = P.rows.get(r); return { attrs: row.attrs, s: Object.fromEntries([...row.cells].map(([k, v]) => [k, styleOf(v)])) }; };
  const sFirst = styleRow(L.itemStart), sLast = styleRow(L.itemEnd), sMid = styleRow(Math.min(L.itemStart + 1, L.itemEnd));
  for (let r = L.itemStart; r <= L.itemEnd; r++) P.rows.delete(r);

  // 2) 아래 행들 밀기
  if (delta) {
    const moved = new Map();
    for (const [r, v] of P.rows) moved.set(r > L.itemEnd ? r + delta : r, v);
    P.rows = moved;
  }
  const s0 = L.itemStart, s1 = L.itemStart + n - 1, T = L.total + delta;

  // 3) 품목 행 작성
  const amtCols = []; for (let c = colNum(C.amt); c <= colNum(C.amtEnd); c++) amtCols.push(colLet(c));
  d.items.forEach((it, i) => {
    const r = s0 + i;
    const st = i === 0 ? sFirst : i === n - 1 ? sLast : sMid;
    P.rows.set(r, { attrs: st.attrs, cells: new Map() });
    const put = (L_, v) => setCell(P.rows, `${L_}${r}`, v, false, st.s[L_]);
    for (const k of Object.keys(st.s)) put(k, null);
    put(C.no, i + 1);
    put(C.name, it.name);
    put(C.qty, Number(it.qty) || 0);
    put(C.amt, Number(it.amount) || 0);
    if (i === 0 || !L.vMerge.includes("date")) put(C.date, d.purchaseDate ? { date: d.purchaseDate } : null);
    if (i === 0 || !L.vMerge.includes("note")) put(C.note, d.note || null);
  });

  // 4) 합계
  const sumQ = d.items.reduce((a, b) => a + (Number(b.qty) || 0), 0);
  const sumA = d.items.reduce((a, b) => a + (Number(b.amount) || 0), 0);
  setCell(P.rows, `${C.qty}${T}`, { f: `SUM(${C.qty}${s0}:${C.qty}${s1})`, v: sumQ });
  setCell(P.rows, `${C.amt}${T}`, { f: `SUM(${C.amt}${s0}:${C.amtEnd}${s1})`, v: sumA });

  // 5) 머리말
  for (const [ref, v] of Object.entries(L.header(d))) setCell(P.rows, ref, v);

  xml = buildRows(P);

  // 6) 병합 셀 다시 계산
  xml = xml.replace(/<mergeCells[^>]*>([\s\S]*?)<\/mergeCells>/, (_, inner) => {
    const refs = [...inner.matchAll(/ref="([A-Z]+)(\d+):([A-Z]+)(\d+)"/g)]
      .map((m) => [m[1], +m[2], m[3], +m[4]])
      .filter(([, a, , b]) => b < L.itemStart || a > L.itemEnd)
      .map(([c1, a, c2, b]) => (a > L.itemEnd ? [c1, a + delta, c2, b + delta] : [c1, a, c2, b]));
    for (let r = s0; r <= s1; r++) refs.push([C.amt, r, C.amtEnd, r]);
    if (n > 1) for (const k of L.vMerge) refs.push([C[k], s0, C[k], s1]);
    return `<mergeCells count="${refs.length}">${refs.map(([a, b, c, e]) => `<mergeCell ref="${a}${b}:${c}${e}"/>`).join("")}</mergeCells>`;
  });
  xml = xml.replace(/<dimension ref="[^"]*"\/>/, "");
  zip.file(sp, xml);

  // 7) 인쇄 영역 행 보정 + 열 때 다시 계산
  let wb = await zip.file("xl/workbook.xml").async("string");
  const sheetIdx = [...wb.matchAll(/<sheet [^>]*name="([^"]+)"/g)].findIndex((m) => m[1] === L.sheet);
  wb = wb.replace(new RegExp(`(<definedName name="_xlnm.Print_Area" localSheetId="${sheetIdx}">[^<]*?\\$)(\\d+)(</definedName>)`), (_, a, r, c) => a + (+r + delta) + c);
  wb = /<calcPr\b/.test(wb) ? wb.replace(/<calcPr\b([^>]*?)\/>/, (m, a) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, "")} fullCalcOnLoad="1"/>`)
    : wb.replace("</workbook>", `<calcPr fullCalcOnLoad="1"/></workbook>`);
  zip.file("xl/workbook.xml", wb);

  // 8) 사진 배치 — 왼쪽: 검수사진 / 오른쪽: 영수증
  const geo = geometry(xml, T + L.photoRows + 2);
  const pr0 = T + 1, pr1 = T + L.photoRows - 1;
  const pl = [];
  if (d.photos.length) pl.push(...placeImages(geo, d.photos, L.left[0], L.left[1], pr0, pr1));
  if (d.receipts.length) pl.push(...placeImages(geo, d.receipts, L.right[0], L.right[1], pr0, pr1));
  await addPictures(zip, sp, pl);

  // 9) 기안사본 (1번 양식만)
  if (L.draftSheet && d.drafts?.length) {
    const dp = await sheetPath(zip, L.draftSheet);
    const dxml = await zip.file(dp).async("string");
    const b = L.draftBox;
    await addPictures(zip, dp, placeImages(geometry(dxml, b.r1 + 2), d.drafts.slice(0, 1), b.c0, b.c1, b.r0, b.r1, 40000));
  }

  return zip.generateAsync({ type: "blob", mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", compression: "DEFLATE" });
}
