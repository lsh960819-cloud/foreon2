"""원본 지출증빙 엑셀 → 웹용 빈 양식(public/expense/*.xlsx) 생성 (1회성 도구)
사용: python scripts/make_expense_templates.py 원본1.xlsx 원본2.xlsx
"""
import re, sys, zipfile, posixpath

def load(p):
    z = zipfile.ZipFile(p); return {n: z.read(n) for n in z.namelist()}

def rels_path(part):
    d, b = posixpath.split(part); return posixpath.join(d, "_rels", b + ".rels")

def targets(files, part):
    rp = rels_path(part)
    if rp not in files: return []
    out = []
    for m in re.finditer(rb'<Relationship [^>]*?Target="([^"]+)"[^>]*?/>', files[rp]):
        if b'TargetMode="External"' in m.group(0): continue
        t = m.group(1).decode()
        out.append(t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(part), t)))
    return out

def gc(files):
    seen, stack = set(), ["_rels/.rels"[:0] or ""]
    roots = [t for t in targets(files, "")]  # from _rels/.rels
    stack = list(roots)
    while stack:
        p = stack.pop()
        if p in seen or p not in files: continue
        seen.add(p); stack += targets(files, p)
    keep = {"[Content_Types].xml", "_rels/.rels"} | seen | {rels_path(p) for p in seen}
    for n in list(files):
        if n not in keep: del files[n]
    ct = files["[Content_Types].xml"].decode()
    ct = re.sub(r'<Override PartName="/([^"]+)"[^>]*/>', lambda m: m.group(0) if m.group(1) in files else "", ct)
    if 'Extension="jpeg"' not in ct:
        ct = ct.replace("<Default ", '<Default Extension="jpeg" ContentType="image/jpeg"/><Default ', 1)
    files["[Content_Types].xml"] = ct.encode()

def remove_sheets(files, names):
    wb = files["xl/workbook.xml"].decode()
    rels = files["xl/_rels/workbook.xml.rels"].decode()
    sheets = re.findall(r'<sheet [^>]*/>', wb)
    names_order = [re.search(r'name="([^"]+)"', s).group(1) for s in sheets]
    keep_idx = [i for i, n in enumerate(names_order) if n not in names]
    for s in sheets:
        n = re.search(r'name="([^"]+)"', s).group(1)
        if n in names:
            rid = re.search(r'r:id="([^"]+)"', s).group(1)
            wb = wb.replace(s, "")
            rels = re.sub(r'<Relationship Id="%s"[^>]*/>' % rid, "", rels)
    # calcChain 제거
    rels = re.sub(r'<Relationship [^>]*calcChain[^>]*/>', "", rels)
    # definedNames 재정리
    def fix(m):
        d = m.group(0)
        if "#REF!" in d: return ""
        li = re.search(r'localSheetId="(\d+)"', d)
        if li:
            i = int(li.group(1))
            if i not in keep_idx: return ""
            d = d.replace(li.group(0), 'localSheetId="%d"' % keep_idx.index(i))
        return d
    wb = re.sub(r'<definedName [^>]*>.*?</definedName>', fix, wb)
    wb = wb.replace("<definedNames></definedNames>", "")
    wb = re.sub(r'activeTab="\d+"', 'activeTab="0"', wb); wb = re.sub(r' firstSheet="\d+"', '', wb)
    files["xl/workbook.xml"] = wb.encode(); files["xl/_rels/workbook.xml.rels"] = rels.encode()
    files.pop("xl/calcChain.xml", None)

def sheet_part(files, name):
    wb = files["xl/workbook.xml"].decode()
    rid = re.search(r'<sheet name="%s"[^>]*r:id="([^"]+)"' % re.escape(name), wb).group(1)
    t = re.search(r'Id="%s"[^>]*Target="([^"]+)"' % rid, files["xl/_rels/workbook.xml.rels"].decode()).group(1)
    return posixpath.normpath(posixpath.join("xl", t.lstrip("/").replace("xl/", "", 1) if t.startswith("/") else t))

def drawing_of(files, sp):
    for t in targets(files, sp):
        if "/drawings/drawing" in t: return t

def strip_pics(files, dp, keep_camera=True):
    x = files[dp].decode()
    def f(m):
        a = m.group(0)
        return a if (keep_camera and "cameraTool" in a) else ""
    x = re.sub(r'<xdr:(twoCellAnchor|oneCellAnchor|absoluteAnchor)\b.*?</xdr:\1>', f, x, flags=re.S)
    files[dp] = x.encode()
    # 안 쓰는 이미지 관계 정리
    rp = rels_path(dp); used = set(re.findall(r'r:embed="([^"]+)"', x))
    files[rp] = re.sub(r'<Relationship Id="([^"]+)"[^>]*/>', lambda m: m.group(0) if m.group(1) in used else "", files[rp].decode()).encode()

def clear_cells(files, sp, refs):
    x = files[sp].decode()
    for r in refs:
        x = re.sub(r'<c r="%s"([^>]*?)(?:/>|>.*?</c>)' % r,
                   lambda m: '<c r="%s"%s/>' % (r, re.sub(r'\s+t="[^"]*"', "", m.group(1))), x, count=1)
    files[sp] = x.encode()

def save(files, out):
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", files.pop("[Content_Types].xml"))
        for n, b in files.items(): z.writestr(n, b)

# ── 1번 (결재란 있음): 검수서 + 기안사본 + 결제(결재란 원본) ──
f = load(sys.argv[1])
remove_sheets(f, ["검수사진", "Sheet7", "Sheet8", "Sheet1"])
for s in ["검수서", "기안사본"]:
    strip_pics(f, drawing_of(f, sheet_part(f, s)))
sp = sheet_part(f, "검수서")
clear_cells(f, sp, ["F4", "F5", "F6", "B6"] + [c + str(r) for r in range(8, 12) for c in "BCDEFI"] + ["E12", "F12"])
gc(f); save(f, "public/expense/tpl_approval.xlsx")

# ── 2번 (결재란 없음): 검수서 + Sheet1(결재란 원본) ──
f = load(sys.argv[2])
remove_sheets(f, ["사진 추가"])
strip_pics(f, drawing_of(f, sheet_part(f, "검수서")))
sp = sheet_part(f, "검수서")
clear_cells(f, sp, ["F5", "F6", "F7", "A9", "B9", "C9", "D9", "E9", "G9", "D10", "E10", "S31"])
gc(f); save(f, "public/expense/tpl_simple.xlsx")
print("ok")
