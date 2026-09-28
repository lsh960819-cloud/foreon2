# -*- coding: utf-8 -*-
"""
Vercel Python 함수: 바이비(BYB) 강좌 명단 → 출석부 엑셀(zip)
POST /api/attendance  {action: login | verify | windows | members | build, ...}
- 비밀번호/토큰은 저장하지 않음. 로그인 정보(auth)는 브라우저 메모리에만 보관되고 요청마다 전달됨.
"""
import base64
import calendar
import http.cookiejar
import io
import json
import re
import tempfile
import time
import urllib.error
import urllib.parse
import urllib.request
import zipfile
from collections import OrderedDict
from http.server import BaseHTTPRequestHandler
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.worksheet.formula import ArrayFormula

TEMPLATE_DIR = Path(__file__).resolve().parent / "_templates"
AUTH_BASE = "https://auth.htbeyond.net/byb"
API_BASE = "https://manager-v2-api.htbeyond.net/mid"
CLIENT_ID = "3mif4qb6d5l8fp19jdql2dhufi"
ADMIN_ORIGIN = "https://bybadmin.htbeyondcloud.com"
SITE_ID = "EDGE524"
PRODUCT_KEYWORD = "2단지"
DAYS = "월화수목금토일"
ID_KEYS = ["blueId", "userId", "username", "loginId", "memberId", "reserverId", "ownerId"]
KINDS = OrderedDict([("수영", "수영"), ("농구", "농구교실"), ("축구", "축구교실")])
# 채널 목록을 못 읽을 때 쓰는 예비값 (바이비 관리자 화면에서 확인한 값)
FALLBACK_CHANNELS = [65850,66949,66951,66948,66950,66952,66953,66277,65773,65774,65769,65770,65771,65772,66327,65764,65765,65768,65767,66420,68117,66235,66301,68115,66691,65853,65857,68118,66693,66687,65854,65858,66690,68119,66689,66236,68113,69025,68120,66688,65855,65859,69047,69046,66303,68114,65856,65860,69086,68109,65864,65867,65866,68111,65865,66698,66699,68110,65868,66116,66115,68112,65869,66118,66696,66117,66305,66624,66304,68808,68809,68810,68811,68812,68636,68915,68832,68838,68833,68834,68835,68830,68836,68827,68814,68828,68837,68831,68826,68829,65775,65776,65777,65778,66217,66712,68867,67556,67986,66346,67985,66523,67984,66342,67983,66522,66918,66917,66916,68787,68944,66915,65829,65830,65837,65838,65839,65840,65841,65842,68338,66355,69008,66692,66302,68116,68943]

LOG = []


def print(*a, **k):  # noqa: A001  엑셀 생성 로그를 모아서 화면에 돌려줌
    LOG.append(" ".join(str(x) for x in a))


class BybError(Exception):
    pass


# ───────────── 바이비 통신 ─────────────
def _req(url, method="GET", body=None, headers=None, cookies=None, form=False):
    h = {"Origin": ADMIN_ORIGIN, "Referer": ADMIN_ORIGIN + "/", "Accept": "application/json, text/plain, */*",
         "User-Agent": "Mozilla/5.0 foreon2-attendance"}
    if headers:
        h.update(headers)
    data = None
    if body is not None:
        if form:
            data = urllib.parse.urlencode(body).encode()
            h["Content-Type"] = "application/x-www-form-urlencoded"
        else:
            data = json.dumps(body).encode()
            h["Content-Type"] = "application/json"
    if cookies:
        h["Cookie"] = "; ".join(f"{k}={v}" for k, v in cookies.items())
    req = urllib.request.Request(url, data=data, method=method, headers=h)
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=25) as r:
                raw = r.read()
                status, set_cookie = r.status, r.headers.get_all("Set-Cookie") or []
                break
        except urllib.error.HTTPError as e:
            raw, status, set_cookie = e.read(), e.code, e.headers.get_all("Set-Cookie") or []
            if status in (429, 503) and attempt < 2:
                time.sleep(3 * (attempt + 1))
                continue
            break
    if cookies is not None:
        for c in set_cookie:
            kv = c.split(";", 1)[0]
            if "=" in kv:
                k, v = kv.split("=", 1)
                cookies[k.strip()] = v.strip()
    try:
        js = json.loads(raw.decode("utf-8")) if raw else None
    except Exception:
        js = None
    return status, js, raw[:300].decode("utf-8", "ignore")


def byb_login(user, pw, cookies):
    st, js, txt = _req(f"{AUTH_BASE}/oauth2/token", "POST",
                       {"grant_type": "password", "client_id": CLIENT_ID, "username": user, "password": pw},
                       cookies=cookies, form=True)
    return st, js, txt


def pick_auth(tok):
    """토큰 응답에서 manager API가 받아주는 Authorization 형식 찾기"""
    cands = []
    for key in ("access_token", "id_token", "accessToken", "idToken"):
        v = (tok or {}).get(key)
        if v:
            cands += [f"Bearer {v}", v]
    for a in cands:
        st, _, _ = _req(f"{API_BASE}/membermanage/manager/manager/current", headers={"Authorization": a})
        if st == 200:
            return a
    raise BybError(f"토큰 형식 확인 실패 (응답 키: {list((tok or {}).keys())})")


def api(auth, method, path, body=None):
    st, js, txt = _req(f"{API_BASE}{path}", method, body, headers={"Authorization": auth["a"]})
    if st == 401:
        raise BybError("로그인이 만료됐어요. 다시 로그인해 주세요.")
    if st != 200:
        raise BybError(f"바이비 응답 오류 {st}: {txt}")
    time.sleep(0.05)
    return js


def get_channels(auth):
    try:
        js = api(auth, "GET", "/community/manager/v3/community/channels")
        ids = []

        def walk(o):
            if isinstance(o, dict):
                if isinstance(o.get("id"), int) and ("name" in o or "title" in o):
                    ids.append(o["id"])
                for v in o.values():
                    walk(v)
            elif isinstance(o, list):
                for v in o:
                    walk(v)
        walk(js)
        return ids or FALLBACK_CHANNELS
    except BybError:
        return FALLBACK_CHANNELS


def list_windows(auth):
    payload = {"channelIds": get_channels(auth), "deleted": False, "title": PRODUCT_KEYWORD,
               "visible": None, "repeatSettings": ["MONTHLY_REPETITION", "NONE"]}
    out, seen, pg = [], set(), 0
    while pg < 20:
        js = api(auth, "POST", f"/community/manager/v3/klass/windows?page={pg}&size=100", payload)
        its = items_of(js)
        for it in its:
            wid = it.get("id") or it.get("windowId") or find_key(it, ["windowId", "id"])
            title = find_title(it)
            if wid and title and wid not in seen:
                seen.add(wid)
                out.append({"id": wid, "name": class_name_from_row(title) or title})
        if len(its) < 100:
            break
        pg += 1
    return out


def members_of(auth, wid, year, month, me):
    last = calendar.monthrange(year, month)[1]
    ids, pg = [], 0
    while pg < 20:
        q = urllib.parse.urlencode({
            "page": pg, "size": 100, "scenarioType": "KLASS", "latestRevision.status": "CONFIRMED",
            "sort": "createdAt,asc", "window.id": wid,
            "startOffsetDateTime": f"{year}-{month:02d}-01T00:00:00+09:00",
            "endOffsetDateTime": f"{year}-{month:02d}-{last}T23:59:59+09:00"})
        its = items_of(api(auth, "GET", f"/community/manager/v1/reservations?{q}"))
        for it in its:
            uid = find_key(it, ID_KEYS)
            if uid and uid != me and uid not in ids:
                ids.append(uid)
        if len(its) < 100:
            break
        pg += 1
    mem = []
    for i in range(0, len(ids), 50):
        js = api(auth, "POST", "/membermanage/manager/member/blueIdList", {"userList": ids[i:i + 50]})
        got = {u.get("blueId"): u for u in (js or []) if isinstance(u, dict)}
        mem += api_to_members([got[x] for x in ids[i:i + 50] if x in got])
    return mem


# ───────────── 엑셀 (로컬 프로그램과 동일한 로직) ─────────────
def parse_class(name):
    """'2단지 요가(화목07시김수정)' → base, 요일, 강사"""
    m = re.match(r"^(.*?)\((.*)\)\s*$", name)
    base, inner = (m.group(1), m.group(2)) if m else (name, "")
    d = re.search(rf"([{DAYS}]+)\d{{1,2}}시", inner)
    t = re.search(r"\d{1,2}시(?:\d{1,2}분)?/?([가-힣]+)$", inner)
    return base.strip(), (d.group(1) if d else ""), (t.group(1) if t else "")


def clean_phone(s):
    d = re.sub(r"\D", "", s or "")
    if len(d) == 11:
        return f"{d[:3]}-{d[3:7]}-{d[7:]}"
    if len(d) == 10:
        return f"{d[:3]}-{d[3:6]}-{d[6:]}"
    return (s or "").strip()


def api_to_members(items):
    out = []
    for u in items or []:
        dong = ho = ""
        for v in (u.get("residentVerifications") or []):
            if v.get("siteId") == SITE_ID and v.get("status") == "VERIFIED":
                dong, ho = v.get("dong") or "", v.get("ho") or ""
        if not dong:
            for v in (u.get("pin_verified") or []):
                if v.get("siteId") == SITE_ID:
                    dong, ho = v.get("dong") or "", v.get("ho") or ""
        phone = (u.get("phone_number") or "").replace("+82", "0", 1).replace("00", "0", 1) \
            if (u.get("phone_number") or "").startswith("+82") else u.get("phone_number") or ""
        m = {"name": u.get("name") or "", "dong": str(dong), "ho": str(ho), "phone": clean_phone(phone)}
        if m not in out:   # 중복 제거
            out.append(m)
    return out


def class_name_from_row(text):
    m = re.search(rf"{PRODUCT_KEYWORD}[^\t\n]*?\([^)\n]*\)", text)
    if not m:
        return None
    name = re.sub(r"\s+\(", "(", m.group(0).strip())
    return re.sub(r"[/\\?*\[\]:]", "", name)


def is_class_sheet(ws):
    return ws["B3"].value == "성명" and "(" in ws.title


def is_helper(ws):
    return isinstance(ws["A1"].value, int) and isinstance(ws["B1"].value, str) and "월" in ws["B1"].value


def slot_rows(ws):
    return [r for r in range(5, ws.max_row + 1) if isinstance(ws.cell(r, 1).value, int)]


def clone_sheet(wb, src, title):
    ws = wb.copy_worksheet(src)
    ws.title = title
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, ArrayFormula):
                c.value = ArrayFormula(c.value.ref, c.value.text)
    for cf in src.conditional_formatting:
        for rule in cf.rules:
            ws.conditional_formatting.add(str(cf.sqref), rule)
    if src.print_area:
        areas = [a.split("!")[-1] for a in re.split(r",(?=')|,(?=\$)", src.print_area)]
        ws.print_area = ",".join(areas)
    if src.print_title_rows:
        ws.print_title_rows = src.print_title_rows
    return ws


def set_instructor(ws, old, new):
    if not new:
        return
    for row in ws.iter_rows(min_row=1, max_row=2):
        for c in row:
            if isinstance(c.value, str) and "강사" in c.value:
                c.value = c.value.replace(old, new) if old and old in c.value else re.sub(
                    r"(강사\s*:?\s*)[가-힣]*|[가-힣]+(\s*강사)", lambda m: (m.group(1) or "") + new + (m.group(2) or ""), c.value, count=1)


def fill_sheet(ws, members):
    rows = slot_rows(ws)
    for r in rows:
        for col in range(2, 6):
            ws.cell(r, col).value = None
    for r, m in zip(rows, members):
        ws.cell(r, 2).value = m["name"]
        ws.cell(r, 3).value = m["dong"]
        ws.cell(r, 4).value = m["ho"]
        ws.cell(r, 5).value = m["phone"]
    if len(members) > len(rows):
        print(f"  ! {ws.title}: 인원 {len(members)}명 > 칸 {len(rows)}개 (초과분 누락)")


def ensure_helper(wb, day, src_day):
    """새 요일(예: 토) 날짜 계산 시트가 없으면 기존 것을 복사해 만듦"""
    if day in wb.sheetnames:
        return
    base = wb[src_day] if src_day in wb.sheetnames else next(ws for ws in wb.worksheets if is_helper(ws))
    h = clone_sheet(wb, base, day)
    h["A2"].value = day[0]
    h["A3"].value = day[1] if len(day) > 1 else None
    print(f"  + 날짜 시트 생성: {day}")


def swap_day_refs(ws, old, new):
    pat = re.compile(rf"(?<![가-힣])'?{old}'?!")
    for row in ws.iter_rows():
        for c in row:
            v = c.value
            if isinstance(v, ArrayFormula):
                c.value = ArrayFormula(v.ref, pat.sub(f"{new}!", v.text))
            elif isinstance(v, str) and v.startswith("="):
                c.value = pat.sub(f"{new}!", v)


def build_workbook(tpl_path, classes, data, year, month, holidays, out_path):
    wb = load_workbook(tpl_path)
    existing = {ws.title: ws for ws in wb.worksheets if is_class_sheet(ws)}
    samples = list(existing.values())
    made = []
    for cname in classes:
        title = cname[:31]
        if title in existing:
            ws = existing[title]
        else:
            _, day, instr = parse_class(cname)
            src = next((s for s in samples if parse_class(s.title)[1] == day), samples[0])
            ws = clone_sheet(wb, src, title)
            src_day = parse_class(src.title)[1]
            if day and src_day and day != src_day:
                ensure_helper(wb, day, src_day)
                swap_day_refs(ws, src_day, day)
            set_instructor(ws, parse_class(src.title)[2], instr)
            a1 = ws["A1"].value
            if isinstance(a1, str) and not a1.startswith("="):
                base, _, _ = parse_class(cname)
                inner = re.sub(rf"{instr}\)$", ")", cname[len(base):]) if instr else cname[len(base):]
                ws["A1"].value = (base.replace(PRODUCT_KEYWORD, "").strip() + " " + inner.strip("()")).strip()
            print(f"  + 새 시트 생성: {title}")
        fill_sheet(ws, data[cname])
        made.append(ws)

    for title, ws in existing.items():
        if ws not in made:
            wb.remove(ws)
    for ws in wb.worksheets:
        if is_helper(ws):
            ws["A1"].value = year
            ws["B1"].value = f"{month}월"
            for i in range(2, 11):
                ws.cell(i, 11).value = None
            for i, d in enumerate(holidays[:9]):
                ws.cell(2 + i, 11).value = d
    others = [ws for ws in wb._sheets if ws not in made]
    wb._sheets = made + others
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws is made[0]
    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    print(f"  저장: {out_path.name}")


def kind_of(text):
    return next((k for k in KINDS if k in text), "GX")


def items_of(js):
    """응답에서 목록 부분 꺼내기 (list 또는 {content:[...]} 등)"""
    if isinstance(js, list):
        return js
    if isinstance(js, dict):
        for k in ("content", "items", "data", "list", "results", "reservations", "windows"):
            v = js.get(k)
            if isinstance(v, list):
                return v
            if isinstance(v, dict):
                r = items_of(v)
                if r:
                    return r
        for v in js.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v
    return []


def find_key(obj, keys, depth=0):
    """중첩된 dict에서 keys 순서대로 첫 문자열 값 찾기"""
    if depth > 4:
        return None
    if isinstance(obj, dict):
        for k in keys:
            v = obj.get(k)
            if isinstance(v, str) and v:
                return v
        for v in obj.values():
            r = find_key(v, keys, depth + 1)
            if r:
                return r
    elif isinstance(obj, list):
        for v in obj:
            r = find_key(v, keys, depth + 1)
            if r:
                return r
    return None


def find_title(obj, depth=0):
    if depth > 4:
        return None
    if isinstance(obj, dict):
        for k in ("title", "name", "productName", "windowName"):
            v = obj.get(k)
            if isinstance(v, str) and PRODUCT_KEYWORD in v:
                return v
        for v in obj.values():
            r = find_title(v, depth + 1)
            if r:
                return r
    return None


def build_zip(data, year, month, holidays, skip_empty=True):
    tpls = OrderedDict()
    for p in sorted(TEMPLATE_DIR.glob("*.xlsx")):
        sheets = [ws.title for ws in load_workbook(p).worksheets if is_class_sheet(ws)]
        kinds = [kind_of(t) for t in sheets]
        tpls.setdefault(max(set(kinds), key=kinds.count), p)
    assign = OrderedDict((k, []) for k in tpls)
    empty = []
    for cname, mem in data.items():
        if not mem and skip_empty:
            empty.append(cname)
            continue
        k = kind_of(cname)
        if k in assign:
            assign[k].append(cname)
    if empty:
        print(f"확정 인원 0명이라 제외: {len(empty)}개 강좌")
    buf = io.BytesIO()
    with tempfile.TemporaryDirectory() as td, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        out = Path(td)
        for k, classes in assign.items():
            if not classes:
                continue
            label = KINDS.get(k, "GX")
            if k == "수영":
                groups = OrderedDict()
                for c in classes:
                    groups.setdefault(parse_class(c)[2] or "강사미상", []).append(c)
                jobs = [(cl, f"{year}년 {month}월 {label} 출석부_{ins}.xlsx") for ins, cl in groups.items()]
            else:
                jobs = [(classes, f"{year}년 {month}월 {label} 출석부.xlsx")]
            for cl, fname in jobs:
                build_workbook(tpls[k], cl, data, year, month, holidays, out / fname)
                z.write(out / fname, fname)
    return buf.getvalue()


# ───────────── HTTP 핸들러 ─────────────
def encode_auth(d):
    return base64.urlsafe_b64encode(json.dumps(d).encode()).decode()


def decode_auth(s):
    try:
        return json.loads(base64.urlsafe_b64decode(s.encode()).decode())
    except Exception:
        raise BybError("로그인 정보가 없어요. 다시 로그인해 주세요.")


def handle(b):
    global PRODUCT_KEYWORD
    act = b.get("action")
    if b.get("keyword"):
        PRODUCT_KEYWORD = b["keyword"]

    if act == "login":
        cookies = {}
        st, js, txt = byb_login(b["user"], b["password"], cookies)
        if st == 200 and js:
            return {"auth": encode_auth({"a": pick_auth(js), "u": b["user"]})}
        st2, _, txt2 = _req(f"{AUTH_BASE}/session/userinfo/otp/request", "POST", cookies=cookies)
        if st2 != 200:
            raise BybError(f"로그인 실패 ({st}): 아이디/비밀번호를 확인해 주세요. [{txt[:120]}] [{st2} {txt2[:120]}]")
        # 인증번호 확인 후 토큰을 다시 받으려면 비밀번호가 필요 → 쿠키와 함께 임시로 돌려줌(브라우저 메모리에만 보관)
        return {"needOtp": True, "pending": encode_auth({"c": cookies, "u": b["user"], "p": b["password"]})}

    if act == "verify":
        pend = decode_auth(b["pending"])
        cookies = pend["c"]
        st, _, txt = _req(f"{AUTH_BASE}/session/userinfo/otp/verify", "POST", {"code": b["code"]}, cookies=cookies)
        if st != 200:
            raise BybError(f"인증번호 확인 실패 ({st}): {txt[:150]}")
        st, js, txt = byb_login(pend["u"], pend["p"], cookies)
        if st != 200:
            raise BybError(f"토큰 발급 실패 ({st}): {txt[:150]}")
        return {"auth": encode_auth({"a": pick_auth(js), "u": pend["u"]})}

    if act == "build":
        LOG.clear()
        data = OrderedDict(b["data"])
        z = build_zip(data, int(b["year"]), int(b["month"]),
                      [int(x) for x in b.get("holidays", [])], b.get("skipEmpty", True))
        return {"zip": base64.b64encode(z).decode(), "log": LOG[-200:]}

    auth = decode_auth(b.get("auth", ""))

    if act == "windows":
        return {"windows": list_windows(auth)}

    if act == "members":
        y, m = int(b["year"]), int(b["month"])
        res = {}
        for w in b["windows"][:15]:
            res[w["name"]] = members_of(auth, w["id"], y, m, auth.get("u"))
        return {"members": res}

    raise BybError("알 수 없는 요청")


class handler(BaseHTTPRequestHandler):
    def _send(self, code, obj):
        raw = json.dumps(obj, ensure_ascii=False).encode()
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(raw)

    def do_POST(self):
        try:
            n = int(self.headers.get("Content-Length") or 0)
            body = json.loads(self.rfile.read(n) or b"{}")
            self._send(200, handle(body))
        except BybError as e:
            self._send(400, {"error": str(e)})
        except Exception as e:  # noqa
            self._send(500, {"error": f"서버 오류: {type(e).__name__}: {e}"})
