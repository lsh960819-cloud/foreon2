import React, { useState, useMemo } from "react";
import {
  LogIn, LogOut, Search, PackageSearch, ClipboardList, CalendarDays, Home as HomeIcon,
  Plus, X, User, Loader2, MessageSquareWarning, ArrowLeftRight, Inbox, Send, Upload,
  CheckCircle2, Clock, FileText, GraduationCap, PlayCircle, Pencil, Trash2, Copy
} from "lucide-react";
import * as XLSX from "xlsx";
import { supabase } from "./supabaseClient.js";

/* ═══════════════════════════════════════════════════════════
   계정: foreon1 ~ foreon9 (골프장·피트니스 계정은 폐지)
   부서 배정 (필요하면 여기 숫자만 바꾸면 됩니다)
   ═══════════════════════════════════════════════════════════ */
const DEPT_OF = {
  foreon1: "사무실", foreon2: "사무실", foreon3: "사무실", foreon4: "사무실", foreon5: "사무실",
  foreon6: "데스크", foreon7: "데스크", foreon8: "데스크", foreon9: "데스크",
};
/* 개인별 비밀번호 — 사무실: Office{번호}!26 (패턴형) / 데스크: 무작위(직원끼리 유추 방지) */
const PW_OF = {
  foreon1: "Office1!26", foreon2: "Office2!26", foreon3: "Office3!26", foreon4: "Office4!26", foreon5: "Office5!26",
  foreon6: "xQ4#nZt8Bw", foreon7: "Lp9$rVq3Xm", foreon8: "Gk6&wTf1Qz", foreon9: "Rn2@Xy7Lqe",
};
const AI_ADMIN = "foreon4"; // 관리자 계정 (수강 시작 연월 선택 · 접수기간 수정 권한)

const 민원카테고리 = ["시설 고장", "이용 안내", "회원 접수", "결제·환불", "기타"];
const 담당부서 = ["데스크", "피트니스", "골프장", "사무실"];

/* ── Supabase 영속 저장 ──
   [rows, setRows] 와 똑같이 쓰지만, 바뀐 부분을 자동으로 Supabase에 반영합니다. */
async function syncDiff(table, prev, next, toDb) {
  const pm = new Map(prev.map((r) => [r.id, r]));
  const nm = new Map(next.map((r) => [r.id, r]));
  try {
    for (const [id, r] of nm) {
      const old = pm.get(id);
      if (!old) {
        const { error } = await supabase.from(table).insert(toDb(r));
        if (error) throw new Error("저장 실패: " + error.message);
      } else if (JSON.stringify(toDb(old)) !== JSON.stringify(toDb(r))) {
        const { id: _drop, ...patch } = toDb(r);
        const { error } = await supabase.from(table).update(patch).eq("id", id);
        if (error) throw new Error("수정 실패: " + error.message);
      }
    }
    for (const [id] of pm) {
      if (!nm.has(id)) {
        const { error } = await supabase.from(table).delete().eq("id", id);
        if (error) throw new Error("삭제 실패: " + error.message);
      }
    }
  } catch (e) {
    alert(`[${table}] ${e.message}`);
  }
}

function usePersisted(table, toDb = (r) => r, fromDb = (r) => r) {
  const [rows, setRowsState] = useState([]);
  const rowsRef = React.useRef([]);
  const loaded = React.useRef(false);
  rowsRef.current = rows;

  React.useEffect(() => {
    let alive = true;
    supabase.from(table).select("*").order("id", { ascending: false })
      .then(({ data, error }) => {
        if (!alive) return;
        if (error) console.error(table, error.message);
        const list = (data || []).map(fromDb);
        rowsRef.current = list;
        setRowsState(list);
        loaded.current = true;
      });
    return () => { alive = false; };
  }, [table]);

  const setRows = (updater) => {
    const prev = rowsRef.current;
    const next = typeof updater === "function" ? updater(prev) : updater;
    rowsRef.current = next;
    setRowsState(next);
    if (loaded.current) syncDiff(table, prev, next, toDb);
  };
  return [rows, setRows];
}

/* 강습 등록·취소 — 분류 순서 및 강좌 목록 */
const LESSON_CATS = ["요가", "줌바", "타바타", "매트", "방송", "근력", "축구", "농구", "수영", "아쿠아로빅"];
const COURSES = {
  요가: [
    "2단지 요가(수금19시/한승지)", "2단지 요가(수금20시/한승지)", "2단지 요가(수금21시/한승지)",
    "2단지 요가(화목07시/김수정)", "2단지 요가(화목08시/김수정)", "2단지 요가(화목09시/도은진)",
    "2단지 요가(화목10시/도은진)", "2단지 요가(화목20시/이인지)", "2단지 요가(화목21시/이인지)",
    "2단지 요가(수금08시/윤명원)", "2단지 요가(수금09시/윤명원)", "2단지 요가(수금10시/윤명원)",
    "2단지 요가(수금18시/박상희)", "2단지 요가(수금19시/박상희)", "2단지 요가(토19시/송혜정)", "2단지 요가(토20시/송혜정)",
  ],
  줌바: [
    "2단지 줌바댄스(화목11시/김지희)", "2단지 줌바댄스(화목12시/김지희)",
    "2단지 줌바댄스(화목17시/신성희)", "2단지 줌바댄스(화목18시/신성희)",
    "2단지 줌바댄스(수금16시/김지인)", "2단지 줌바댄스(수금17시/김지인)",
    "2단지 줌바댄스(토10시/박나리)", "2단지 줌바댄스(토11시/박나리)",
  ],
  타바타: ["2단지 타바타(수금11시/신수경)", "2단지 타바타(수금12시/신수경)"],
  매트: ["2단지 매트필라테스(화목16시/김혜인)"],
  방송: ["2단지 방송댄스(화목19시/박선화)"],
  근력: ["2단지 근력(토08시/이재욱)", "2단지 근력(토09시/이재욱)"],
  축구: [
    "2단지 축구교실(3~4학년/화14시/이범영)", "2단지 축구교실(1~2학년/화15시/이범영)", "2단지 축구교실(6~7세/화16시/이범영)",
    "2단지 축구교실(성인여성/수10시/이범영)", "2단지 축구교실(1~2학년/수14시/이범영)", "2단지 축구교실(3~4학년/수15시/이범영)",
    "2단지 축구교실(6~7세/수16시/이범영)", "2단지 축구교실(1~2학년/목14시/이범영)", "2단지 축구교실(5~6학년/목15시/이범영)",
    "2단지 축구교실(6~7세/목16시/이범영)", "2단지 축구교실(성인여성/금10시/이범영)", "2단지 축구교실(1~2학년/금14시/이범영)",
    "2단지 축구교실(3~4학년/금15시/이범영)", "2단지 축구교실(6~7세/금16시/이범영)",
  ],
  농구: [
    "2단지 농구교실(1~2학년/화16시/배준태)", "2단지 농구교실(3~4학년/화17시/배준태)", "2단지 농구교실(5~6학년/화18시/배준태)", "2단지 농구교실(중학생/화19시/배준태)",
    "2단지 농구교실(1~2학년/수16시/배준태)", "2단지 농구교실(3~4학년/수17시/배준태)", "2단지 농구교실(5~6학년/수18시/배준태)", "2단지 농구교실(중학생/수19시/배준태)",
    "2단지 농구교실(1~2학년/목16시/배준태)", "2단지 농구교실(3~4학년/목17시/배준태)", "2단지 농구교실(5~6학년/목18시/배준태)", "2단지 농구교실(중학생/목19시/배준태)",
    "2단지 농구교실(1~2학년/금16시/배준태)", "2단지 농구교실(3~4학년/금17시/배준태)", "2단지 농구교실(5~6학년/금18시/배준태)", "2단지 농구교실(중학생/금19시/배준태)",
  ],
  수영: [
    "2단지 성인수영-중급(토20시/전유정)",
    "2단지 성인수영-기초(화목06시30분/손태환)", "2단지 성인수영-초급(화목06시30분/김소운)", "2단지 성인수영-중급(화목06시30분/이규상)", "2단지 성인수영-상급(화목06시30분/박범호)",
    "2단지 성인수영-기초(화목07시30분/손태환)", "2단지 성인수영-초급(화목07시30분/김소운)", "2단지 성인수영-중급(화목07시30분/이규상)", "2단지 성인수영-상급(화목07시30분/박범호)",
    "2단지 성인수영-기초(화목09시/손태환)", "2단지 성인수영-초급(화목09시/김소운)", "2단지 성인수영-중급(화목09시/이규상)", "2단지 성인수영-상급(화목09시/박범호)",
    "2단지 성인수영-기초(화목10시/손태환)", "2단지 성인수영-초급(화목10시/김소운)", "2단지 성인수영-중급(화목10시/이규상)", "2단지 성인수영-상급(화목10시/박범호)",
    "2단지 성인수영-기초(화목20시/윤성수)", "2단지 성인수영-초급(화목20시/박동환)", "2단지 성인수영-중급(화목20시/전유정)", "2단지 성인수영-상급(화목20시/정인혁)",
    "2단지 성인수영-기초(수금06시30분/박범호)", "2단지 성인수영-초급(수금06시30분/이규상)", "2단지 성인수영-중급(수금06시30분/김소운)", "2단지 성인수영-상급(수금06시30분/손태환)",
    "2단지 성인수영-기초(수금07시30분/박범호)", "2단지 성인수영-초급(수금07시30분/이규상)", "2단지 성인수영-중급(수금07시30분/김소운)", "2단지 성인수영-상급(수금07시30분/손태환)",
    "2단지 성인수영-기초(수금09시/박범호)", "2단지 성인수영-초급(수금09시/이규상)", "2단지 성인수영-중급(수금09시/김소운)", "2단지 성인수영-상급(수금09시/손태환)",
    "2단지 성인수영-기초(수금10시/박범호)", "2단지 성인수영-초급(수금10시/이규상)", "2단지 성인수영-중급(수금10시/김소운)", "2단지 성인수영-상급(수금10시/손태환)",
    "2단지 성인수영-기초(수금20시/윤성수)", "2단지 성인수영-초급(수금20시/박동환)", "2단지 성인수영-중급(수금20시/전유정)", "2단지 성인수영-상급(수금20시/정인혁)",
    "2단지 어린이수영-기초(화목16시/정인혁)", "2단지 어린이수영-초급(화목16시/윤성수)", "2단지 어린이수영-중급(화목16시/박동환)", "2단지 어린이수영-상급(화목16시/전유정)",
    "2단지 어린이수영-기초(화목17시/윤성수)", "2단지 어린이수영-초급(화목17시/정인혁)", "2단지 어린이수영-중급(화목17시/전유정)", "2단지 어린이수영-상급(화목17시/박동환)",
    "2단지 어린이수영-기초(화목19시/정인혁)", "2단지 어린이수영-초급(화목19시/윤성수)", "2단지 어린이수영-중급(화목19시/전유정)", "2단지 어린이수영-상급(화목19시/박동환)",
    "2단지 어린이수영-기초(수금16시/정인혁)", "2단지 어린이수영-초급(수금16시/전유정)", "2단지 어린이수영-중급(수금16시/윤성수)", "2단지 어린이수영-상급(수금16시/박동환)",
    "2단지 어린이수영-기초(수금17시/윤성수)", "2단지 어린이수영-초급(수금17시/정인혁)", "2단지 어린이수영-중급(수금17시/전유정)", "2단지 어린이수영-상급(수금17시/박동환)",
    "2단지 어린이수영-기초(수금19시/정인혁)", "2단지 어린이수영-초급(수금19시/윤성수)", "2단지 어린이수영-중급(수금19시/전유정)", "2단지 어린이수영-상급(수금19시/박동환)",
    "2단지 성인수영-기초(토06시30분/박범호)", "2단지 성인수영-초급(토06시30분/이규상)", "2단지 성인수영-중급(토06시30분/손태준)", "2단지 성인수영-상급(토06시30분/김소운)",
    "2단지 성인수영-기초(토07시30분/박범호)", "2단지 성인수영-초급(토07시30분/이규상)", "2단지 성인수영-중급(토07시30분/손태준)", "2단지 성인수영-상급(토07시30분/김소운)",
    "2단지 성인수영-기초초급(토09시/이규상)", "2단지 성인수영-중상급(토09시/김소운)", "2단지 성인수영-상급(토09시/박범호)",
    "2단지 성인수영-기초초급(토10시/이규상)", "2단지 성인수영-중상급(토10시/김소운)", "2단지 성인수영-상급(토10시/박범호)",
    "2단지 성인수영-기초(토16시/박동환)", "2단지 성인수영-상급(토17시/윤성수)",
    "2단지 성인수영-기초초급(토19시/박동환)", "2단지 성인수영-중상급(토19시/정인혁)",
    "2단지 성인수영-기초초급(토20시/박동환)", "2단지 성인수영-상급(토20시/정인혁)",
    "2단지 어린이수영-기초초급(토09시/손태준)", "2단지 어린이수영-중상급(토10시/손태준)",
    "2단지 어린이수영-기초초급(토16시/정인혁)", "2단지 어린이수영-중상급(토16시/윤성수)",
    "2단지 어린이수영-기초초급(토17시/박동환)", "2단지 어린이수영-중상급(토17시/정인혁)",
    "2단지 어린이수영-기초(토19시/윤성수)", "2단지 어린이수영-상급(토20시/윤성수)",
  ],
  아쿠아로빅: [
    "2단지 아쿠아로빅(화목11시/최성준)", "2단지 아쿠아로빅(화목12시/최성준)", "2단지 아쿠아로빅(화목14시/최성준)", "2단지 아쿠아로빅(화목15시/최성준)",
    "2단지 아쿠아로빅(수금11시/오미화)", "2단지 아쿠아로빅(수금12시/오미화)", "2단지 아쿠아로빅(수금14시/오미화)", "2단지 아쿠아로빅(수금15시/오미화)",
    "2단지 아쿠아로빅(토15시/박원정)", "2단지 아쿠아로빅(토14시/박원정)", "2단지 아쿠아로빅(토12시/박원정)", "2단지 아쿠아로빅(토11시/박원정)",
  ],
};

/* ── 강좌 단계 선택(층→종목→반→요일→강사→시간)용 파싱 ── */
const uniqArr = (a) => [...new Set(a)];
const DAY_ORDER = ["화목", "수금", "토", "월", "화", "수", "목", "금", "일"];
const daySort = (a, b) => DAY_ORDER.indexOf(a) - DAY_ORDER.indexOf(b);
const timeMin = (t) => { const m = String(t).match(/(\d{1,2})시(?:\s*(\d{2})분)?/); return m ? (+m[1]) * 60 + (+(m[2] || 0)) : 0; };
const timeSort = (a, b) => timeMin(a) - timeMin(b);
const dayLabel = (d) => ({ 화목: "화/목", 수금: "수/금" }[d] || d);
const FLOOR_INFO = { B1: "GX·구기 (요가·줌바·축구·농구 등)", B2: "수영장 (수영·아쿠아로빅)" };

const ALL_COURSES = Object.entries(COURSES).flatMap(([cat, arr]) =>
  arr.map((full) => {
    const m = full.match(/^2단지\s*([^(]+)\(([^)]*)\)/);
    if (!m) return null;
    let sport = m[1].trim(), grp = "";
    const di = sport.indexOf("-");
    if (di > -1) { grp = sport.slice(di + 1); sport = sport.slice(0, di); }
    const parts = m[2].split("/");
    const teacher = parts[parts.length - 1];
    let dayTime = parts[0];
    if (parts.length === 3) { grp = parts[0]; dayTime = parts[1]; }
    const dm = dayTime.match(/^(화목|수금|월|화|수|목|금|토|일)(.*)$/);
    const days = dm ? dm[1] : dayTime;
    const time = dm ? dm[2] : "";
    const floor = /수영|아쿠아/.test(sport) ? "B2" : "B1";
    return { full, cat, sport, grp, days, time, teacher, floor };
  }).filter(Boolean)
);

function PickRow({ label, options, value, onPick, render }) {
  if (!options.length) return null;
  return (
    <div>
      <span className="text-xs font-medium text-slate-500 block mb-1.5">{label}</span>
      <div className="flex flex-wrap gap-1.5">
        {options.map((o) => (
          <button key={o} type="button" onClick={() => onPick(o)}
            className={`text-sm px-3 py-1.5 rounded-full border transition ${value === o ? "bg-emerald-600 border-emerald-600 text-white font-medium" : "bg-white border-slate-200 text-slate-600 hover:border-emerald-300"}`}>
            {render ? render(o) : o}
          </button>
        ))}
      </div>
    </div>
  );
}

/* 수강 시작 연월 (YYYY-MM) */
const ymNow = () => new Date().toISOString().slice(0, 7);
const ymList = (n = 4) => {
  const d = new Date(); d.setDate(1);
  return Array.from({ length: n }, (_, i) => {
    const x = new Date(d.getFullYear(), d.getMonth() + i, 1);
    return `${x.getFullYear()}-${String(x.getMonth() + 1).padStart(2, "0")}`;
  });
};
const ymLabel = (ym) => {
  const [y, m] = ym.split("-");
  return `${Number(m)}월${ym === ymNow() ? " (이번 달)" : ` (${y})`}`;
};

/* ── 강좌별 정원 (BYB 화면값 대신 실제 운영 정원을 사용) ── */
const UNDER_RATIO = 0.8;   // 정원의 80% 미만이면 '정원 미달'

function capacityOf(course = "") {
  const c = String(course);
  if (c.includes("아쿠아")) return 50;
  if (c.includes("수영")) return 25;
  if (c.includes("축구교실")) return c.includes("6~7세") ? 16 : 18;
  if (c.includes("농구교실")) return c.includes("4~6학년") ? 15 : 12;
  if (/요가|줌바|매트필라테스|방송댄스|타바타|근력/.test(c)) return 30;
  return 30;   // 그 외 기본값
}
const isUnder = (s) => s.enrolled < Math.ceil(capacityOf(s.course) * UNDER_RATIO);
const isOver = (s) => s.enrolled >= capacityOf(s.course);

/* 확인창이 차단된 환경(회사 PC 등)에서도 동작하도록 —
   대화상자가 뜨지 않고 즉시 false 가 오면 차단으로 보고 진행합니다. */
function askYesNo(msg) {
  try {
    const t0 = Date.now();
    const ok = window.confirm(msg);
    if (ok) return true;
    return Date.now() - t0 < 10;   // 사람이 누른 게 아니라 즉시 반환 → 차단된 것
  } catch (e) {
    return true;
  }
}

const today = () => new Date().toISOString().slice(0, 10);
const nowLocal = () => {
  const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 16);
};
const won = (n) => n.toLocaleString("ko-KR") + "원";

/* ── 독서실·골프백 금액표 (날짜 1~31일, 첨부 금액표 그대로 내장) ──
   DEDUCT = 등록 시 차감입력 / REFUND = 취소 시 환불입력 */
const AMT = {
  독서실_개인: {
    price: 130000,
    deduct: [0,0,4330,8660,13000,17330,21660,26000,30330,34660,39000,43330,47660,52000,56330,60660,65000,69330,73660,78000,82330,86660,91000,95330,99660,104000,108330,112660,117000,121330,125660,125660],
    refund: [0,125660,121330,117000,112660,108330,104000,99660,95330,91000,86660,82330,78000,73660,69330,65000,60660,56330,52000,47660,43330,39000,34660,30330,26000,21660,17330,13000,8660,4330,0,0],
  },
  독서실_일반: {
    price: 70000,
    deduct: [0,0,2330,4660,7000,9330,11660,14000,16330,18660,21000,23330,25660,28000,30330,32660,35000,37330,39660,42000,44330,46660,49000,51330,53660,56000,58330,60660,63000,65330,67660,67660],
    refund: [0,67660,65330,63000,60660,58330,56000,53660,51330,49000,46660,44330,42000,39660,37330,35000,32660,30330,28000,25660,23330,21000,18660,16330,14000,11660,9330,7000,4660,2330,0,0],
  },
  골프백_상단: {
    price: 10000,
    deduct: [0,0,330,660,1000,1330,1660,2000,2330,2660,3000,3330,3660,4000,4330,4660,5000,5330,5660,6000,6330,6660,7000,7330,7660,8000,8330,8660,9000,9330,9660,9660],
    refund: [0,9660,9330,9000,8660,8330,8000,7660,7330,7000,6660,6330,6000,5660,5330,5000,4660,4330,4000,3660,3330,3000,2660,2330,2000,1660,1330,1000,660,330,0,0],
  },
  골프백_하단: {
    price: 13000,
    deduct: [0,0,430,860,1300,1730,2160,2600,3030,3460,3900,4330,4760,5200,5630,6060,6500,6930,7360,7800,8230,8660,9100,9530,9960,10400,10830,11260,11700,12130,12560,12560],
    refund: [0,12560,12130,11700,11260,10830,10400,9960,9530,9100,8660,8230,7800,7360,6930,6500,6060,5630,5200,4760,4330,3900,3460,3030,2600,2160,1730,1300,860,430,0,0],
  },
};
// 31일까지 있는 달의 31일은 30일로 취급 (배열은 이미 31=30값과 동일하게 채움)
function calcAmount(product, action, day) {
  const t = AMT[product]; if (!t || !day) return null;
  const d = Math.min(Math.max(day, 1), 31);
  return (action === "등록" ? t.deduct : t.refund)[d];
}
const stamp = () => {
  const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 16).replace("T", " ");
};

/* 수정·삭제 이력 표시 줄 */
function AuditLine({ r }) {
  const last = r.hist && r.hist.length ? r.hist[r.hist.length - 1] : null;
  if (!last) return null;
  return <p className="text-[11px] text-amber-600 mt-1 font-mono">✎ {last.at} {last.by} 수정{r.hist.length > 1 ? ` (총 ${r.hist.length}회)` : ""}</p>;
}
/* 삭제된 항목 표시 (소프트 삭제) */
function DeletedCard({ r, label }) {
  return (
    <div className="bg-slate-50 rounded-xl border border-dashed border-slate-300 p-3 text-xs text-slate-400">
      🗑 삭제된 {label} · <span className="font-mono">{r.deleted.at} {r.deleted.by}</span> 삭제함
    </div>
  );
}

const SESSION_KEY = "foreon2_session";
const SESSION_HOURS = 8;

function loadSession() {
  try {
    const raw = localStorage.getItem(SESSION_KEY);
    if (!raw) return null;
    const s = JSON.parse(raw);
    if (Date.now() - s.at > SESSION_HOURS * 3600 * 1000) {
      localStorage.removeItem(SESSION_KEY);
      return null;
    }
    return { id: s.id, dept: s.dept };
  } catch { return null; }
}
function saveSession(me) {
  try { localStorage.setItem(SESSION_KEY, JSON.stringify({ id: me.id, dept: me.dept, at: Date.now() })); } catch {}
}
function clearSession() {
  try { localStorage.removeItem(SESSION_KEY); } catch {}
}

export default function App() {
  const [me, setMe] = useState(() => loadSession());
  const [tab, setTab] = useState("home");

  const login = (m) => { saveSession(m); setMe(m); };
  const logout = () => { clearSession(); setMe(null); };

  // 우클릭·단축키 방해 (참고: 브라우저 코드 특성상 100% 차단은 불가능합니다.
  // 마음먹은 사용자는 주소창 코드 입력 등으로 우회할 수 있어요 — 초보 접근만 늦추는 수준입니다.)
  React.useEffect(() => {
    const blockKeys = (e) => {
      const k = e.key?.toUpperCase();
      if (k === "F12" || (e.ctrlKey && e.shiftKey && ["I", "J", "C"].includes(k)) || (e.ctrlKey && k === "U")) {
        e.preventDefault();
      }
    };
    const blockRight = (e) => e.preventDefault();
    document.addEventListener("keydown", blockKeys);
    document.addEventListener("contextmenu", blockRight);
    return () => {
      document.removeEventListener("keydown", blockKeys);
      document.removeEventListener("contextmenu", blockRight);
    };
  }, []);

  const [complaints, setComplaints] = usePersisted("complaints");
  const [handovers, setHandovers] = usePersisted("handovers");
  const [requests, setRequests] = usePersisted(
    "requests",
    ({ officeOnly, ...r }) => ({ ...r, officeonly: officeOnly }),
    ({ officeonly, ...r }) => ({ ...r, officeOnly: officeonly }),
  );
  const [events, setEvents] = usePersisted("events");
  const [lost, setLost] = usePersisted("lost");

  if (!me) return <Login onLogin={login} />;
  const isAI = me.id === AI_ADMIN;
  const isOffice = me.dept === "사무실";

  const tabs = [
    { id: "home", label: "홈", icon: HomeIcon },
    { id: "worklog", label: "업무일지", icon: ClipboardList },
    { id: "search", label: "기록 검색", icon: Search },
    { id: "lost", label: "분실물", icon: PackageSearch },
    { id: "events", label: "주요 행사·일정", icon: CalendarDays },
    { id: "lessons", label: "등록·취소·이월", icon: GraduationCap },
    { id: "worklogs", label: "작업 기록", icon: FileText },
    ...(isOffice ? [{ id: "officelog", label: "사무실 업무일지", icon: ClipboardList }] : []),
    { id: "link", label: "외부 연동", icon: ArrowLeftRight },
  ];

  return (
    <div className="min-h-screen bg-slate-50 text-slate-800">
      <div className="bg-white border-b border-slate-200 sticky top-0 z-10">
        <div className="max-w-4xl mx-auto px-4 h-14 flex items-center justify-between">
          <div className="flex items-center gap-2"><span className="w-2 h-2 rounded-full bg-emerald-500" /><span className="font-bold">커뮤니티 운영</span></div>
          <div className="flex items-center gap-2">
            <span className="flex items-center gap-1.5 text-sm rounded-lg bg-slate-100 px-2.5 py-1.5">
              <User size={14} className="text-emerald-600" /><b>{me.id}</b>
              <span className="text-xs text-slate-400 font-mono">{me.dept}</span>
            </span>
            <button onClick={logout} className="text-slate-400 hover:text-slate-600 p-1.5"><LogOut size={16} /></button>
          </div>
        </div>
      </div>

      <div className="max-w-4xl mx-auto px-4 pb-20">
        <div className="flex gap-1 pt-4 pb-3 overflow-x-auto">
          {tabs.map((t) => {
            const on = tab === t.id; const Icon = t.icon;
            return (
              <button key={t.id} onClick={() => setTab(t.id)}
                className={`flex items-center gap-1.5 rounded-lg px-3.5 py-2 text-sm font-medium whitespace-nowrap transition
                  ${on ? "bg-emerald-600 text-white shadow-sm" : "text-slate-500 hover:bg-slate-100"}`}>
                <Icon size={15} /> {t.label}
              </button>
            );
          })}
        </div>

        {tab === "home" && <Home me={me} events={events} />}
        {tab === "worklog" && <WorkLog me={me} complaints={complaints} setComplaints={setComplaints}
          handovers={handovers} setHandovers={setHandovers} requests={requests} setRequests={setRequests} />}
        {tab === "search" && <RecordSearch complaints={complaints} />}
        {tab === "lost" && <Lost me={me} rows={lost} setRows={setLost} />}
        {tab === "events" && <Events me={me} rows={events} setRows={setEvents} />}
        {tab === "lessons" && <Registrations me={me} />}
        {tab === "worklogs" && <WorkLogs me={me} />}
        {tab === "officelog" && isOffice && <OfficeLog me={me} />}
        {tab === "link" && <LinkInfo />}
      </div>
    </div>
  );
}

/* ─────────────────────── 로그인 (아이디 입력형) */
function Login({ onLogin }) {
  const [id, setId] = useState("");
  const [pw, setPw] = useState("");
  const [err, setErr] = useState("");
  const submit = () => {
    const key = id.trim().toLowerCase();
    if (DEPT_OF[key] && pw === PW_OF[key]) { setErr(""); onLogin({ id: key, dept: DEPT_OF[key] }); }
    else setErr("아이디 또는 비밀번호가 올바르지 않습니다.");
  };
  return (
    <div className="min-h-screen bg-slate-100 flex items-center justify-center px-4">
      <div className="w-full max-w-sm">
        <div className="text-center mb-6">
          <div className="inline-flex items-center gap-2 text-xs font-mono tracking-widest text-emerald-700 mb-2"><span className="w-1.5 h-1.5 rounded-full bg-emerald-500" /> FOREON 커뮤니티</div>
          <h1 className="text-2xl font-bold">커뮤니티 운영</h1>
          <p className="text-sm text-slate-500 mt-1">아이디로 로그인하세요</p>
        </div>
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-6 space-y-4">
          <label className="block"><span className="text-xs font-medium text-slate-600 mb-1.5 block">아이디</span>
            <input value={id} onChange={(e) => setId(e.target.value)} placeholder="아이디 입력"
              className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100" /></label>
          <label className="block"><span className="text-xs font-medium text-slate-600 mb-1.5 block">비밀번호</span>
            <input type="password" value={pw} onChange={(e) => setPw(e.target.value)} onKeyDown={(e) => e.key === "Enter" && submit()} placeholder="비밀번호 입력"
              className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100" /></label>
          {err && <p className="text-xs text-red-500">{err}</p>}
          <button onClick={submit} className="w-full flex items-center justify-center gap-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2.5 transition"><LogIn size={16} /> 로그인</button>
          <p className="text-[11px] leading-relaxed text-slate-400 text-center">담당자에게 배정받은 계정으로 로그인하세요.</p>
        </div>
      </div>
    </div>
  );
}

/* ─────────────────────── 업무일지 */
function WorkLog({ me, complaints, setComplaints, handovers, setHandovers, requests, setRequests }) {
  const [sub, setSub] = useState("complaint");
  const subs = [
    { id: "complaint", label: "민원 작성", icon: MessageSquareWarning },
    { id: "handover", label: "인수인계", icon: ArrowLeftRight },
    { id: "request", label: "요청사항", icon: Inbox },
  ];
  return (
    <div>
      <div className="inline-flex gap-1 p-1 bg-slate-100 rounded-xl mb-4">
        {subs.map((s) => {
          const on = sub === s.id; const Icon = s.icon;
          return (
            <button key={s.id} onClick={() => setSub(s.id)}
              className={`flex items-center gap-1.5 rounded-lg px-3 py-1.5 text-sm font-medium transition ${on ? "bg-white shadow-sm text-emerald-700" : "text-slate-500"}`}>
              <Icon size={14} /> {s.label}
            </button>
          );
        })}
      </div>
      {sub === "complaint" && <Complaints me={me} rows={complaints} setRows={setComplaints} />}
      {sub === "handover" && <Handover me={me} rows={handovers} setRows={setHandovers} complaints={complaints} />}
      {sub === "request" && <Requests me={me} rows={requests} setRows={setRequests} />}
    </div>
  );
}

/* 민원 작성 — 모두 확인·작성·수정 가능 */
function Complaints({ me, rows, setRows }) {
  const [open, setOpen] = useState(false);
  const blank = { time: nowLocal(), cat: 민원카테고리[0], dong: "", ho: "", name: "", phone: "", content: "", action: "", owner: 담당부서[0], status: "진행 중" };
  const [v, setV] = useState(blank);
  const add = () => {
    if (!v.content.trim()) return;
    setRows([{ id: Date.now(), ...v, time: v.time.replace("T", " "), by: me.id, dept: me.dept, date: today(), hist: [], deleted: null }, ...rows]);
    setV(blank); setOpen(false);
  };
  const toggle = (id) => setRows(rows.map((r) => r.id === id ? { ...r, status: r.status === "완료" ? "진행 중" : "완료" } : r));
  const saveEdit = (id, patch) => setRows(rows.map((r) => r.id === id ? { ...r, ...patch, hist: [...(r.hist || []), { by: me.id, at: stamp() }] } : r));
  const del = (id) => { if (window.confirm("이 민원을 삭제할까요?")) setRows(rows.map((r) => r.id === id ? { ...r, deleted: { by: me.id, at: stamp() } } : r)); };
  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-base font-bold">민원 대장 <span className="text-slate-400 font-normal text-sm">({rows.filter((r) => !r.deleted).length})</span></h2>
        <button onClick={() => setOpen(!open)} className="flex items-center gap-1 rounded-lg bg-slate-800 hover:bg-slate-900 text-white text-xs font-medium px-3 py-2"><Plus size={14} /> 민원 등록</button>
      </div>
      {open && (
        <div className="bg-white rounded-xl border border-emerald-200 ring-2 ring-emerald-50 p-4 mb-4 space-y-3">
          <div className="flex items-center justify-between"><span className="text-sm font-semibold">새 민원 등록</span><button onClick={() => setOpen(false)} className="text-slate-400"><X size={16} /></button></div>
          <div className="grid grid-cols-2 gap-2">
            <L label="발생 시각"><input type="datetime-local" value={v.time} onChange={(e) => setV({ ...v, time: e.target.value })} className="in" /></L>
            <L label="구분"><select value={v.cat} onChange={(e) => setV({ ...v, cat: e.target.value })} className="in">{민원카테고리.map((c) => <option key={c}>{c}</option>)}</select></L>
          </div>
          <div className="grid grid-cols-4 gap-2">
            <L label="동"><input value={v.dong} onChange={(e) => setV({ ...v, dong: e.target.value })} className="in" /></L>
            <L label="호"><input value={v.ho} onChange={(e) => setV({ ...v, ho: e.target.value })} className="in" /></L>
            <L label="성함"><input value={v.name} onChange={(e) => setV({ ...v, name: e.target.value })} className="in" /></L>
            <L label="연락처"><input value={v.phone} onChange={(e) => setV({ ...v, phone: e.target.value })} className="in" /></L>
          </div>
          <L label="발생 내용 *"><textarea rows={2} value={v.content} onChange={(e) => setV({ ...v, content: e.target.value })} className="in resize-none" placeholder="상세 경위" /></L>
          <L label="조치 사항"><textarea rows={2} value={v.action} onChange={(e) => setV({ ...v, action: e.target.value })} className="in resize-none" placeholder="데스크 조치 내용" /></L>
          <div className="grid grid-cols-2 gap-2">
            <L label="담당 부서"><select value={v.owner} onChange={(e) => setV({ ...v, owner: e.target.value })} className="in">{담당부서.map((c) => <option key={c}>{c}</option>)}</select></L>
            <L label="처리 상태"><select value={v.status} onChange={(e) => setV({ ...v, status: e.target.value })} className="in"><option>진행 중</option><option>완료</option></select></L>
          </div>
          <button onClick={add} className="w-full rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2.5">등록</button>
        </div>
      )}
      <div className="space-y-2">{rows.map((r) => r.deleted
        ? <DeletedCard key={r.id} r={r} label="민원" />
        : <ComplaintCard key={r.id} r={r} onToggle={() => toggle(r.id)} onSave={(patch) => saveEdit(r.id, patch)} onDelete={() => del(r.id)} />)}</div>
      <style>{`.in{width:100%;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:13px;outline:none;background:#fff}
        .in:focus{border-color:#10b981;box-shadow:0 0 0 3px #d1fae5}`}</style>
    </div>
  );
}
function ComplaintCard({ r, onToggle, onSave, onDelete }) {
  const done = r.status === "완료";
  const editable = !!onSave;
  const [edit, setEdit] = useState(false);
  const [v, setV] = useState(r);
  const start = () => { setV(r); setEdit(true); };
  const save = () => { onSave({ cat: v.cat, dong: v.dong, ho: v.ho, name: v.name, phone: v.phone, content: v.content, action: v.action, owner: v.owner, status: v.status }); setEdit(false); };

  if (edit) {
    return (
      <div className="bg-white rounded-xl border border-emerald-200 ring-2 ring-emerald-50 p-4 space-y-2.5">
        <div className="flex items-center justify-between"><span className="text-sm font-semibold">민원 수정</span><button onClick={() => setEdit(false)} className="text-slate-400"><X size={16} /></button></div>
        <div className="grid grid-cols-2 gap-2">
          <select value={v.cat} onChange={(e) => setV({ ...v, cat: e.target.value })} className="ce">{민원카테고리.map((c) => <option key={c}>{c}</option>)}</select>
          <select value={v.status} onChange={(e) => setV({ ...v, status: e.target.value })} className="ce"><option>진행 중</option><option>완료</option></select>
        </div>
        <div className="grid grid-cols-4 gap-2">
          <input value={v.dong} onChange={(e) => setV({ ...v, dong: e.target.value })} placeholder="동" className="ce" />
          <input value={v.ho} onChange={(e) => setV({ ...v, ho: e.target.value })} placeholder="호" className="ce" />
          <input value={v.name} onChange={(e) => setV({ ...v, name: e.target.value })} placeholder="성함" className="ce" />
          <input value={v.phone} onChange={(e) => setV({ ...v, phone: e.target.value })} placeholder="연락처" className="ce" />
        </div>
        <textarea rows={2} value={v.content} onChange={(e) => setV({ ...v, content: e.target.value })} placeholder="발생 내용" className="ce resize-none" />
        <textarea rows={2} value={v.action} onChange={(e) => setV({ ...v, action: e.target.value })} placeholder="조치 사항" className="ce resize-none" />
        <select value={v.owner} onChange={(e) => setV({ ...v, owner: e.target.value })} className="ce">{담당부서.map((c) => <option key={c}>{c}</option>)}</select>
        <div className="flex gap-2">
          <button onClick={save} className="flex-1 rounded-lg bg-emerald-600 text-white text-sm font-medium py-2">저장</button>
          <button onClick={() => setEdit(false)} className="px-4 rounded-lg border border-slate-200 text-sm text-slate-500">취소</button>
        </div>
        <style>{`.ce{width:100%;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:13px;outline:none;background:#fff}.ce:focus{border-color:#10b981;box-shadow:0 0 0 3px #d1fae5}`}</style>
      </div>
    );
  }
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
      <div className="flex items-start justify-between gap-3">
        <div className="min-w-0 flex-1">
          <div className="flex items-center gap-2 flex-wrap">
            <span className="text-xs px-2 py-0.5 rounded bg-slate-100 text-slate-600">{r.cat}</span>
            {(r.dong || r.name) && <span className="text-xs text-slate-500">{r.dong}{r.ho ? `-${r.ho}` : ""} {r.name}</span>}
          </div>
          <p className="font-medium mt-1.5">{r.content}</p>
          {r.action && <p className="text-sm text-emerald-800 bg-emerald-50 border border-emerald-100 rounded-lg px-3 py-1.5 mt-2"><b>조치 · </b>{r.action}</p>}
          <p className="text-[11px] text-slate-400 mt-2 font-mono">{r.time} · {r.owner} · {r.by}</p>
          <AuditLine r={r} />
        </div>
        <div className="shrink-0 flex flex-col items-end gap-2">
          <button onClick={onToggle}
            className={`flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-medium ${done ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"}`}>
            {done ? <CheckCircle2 size={12} /> : <Clock size={12} />} {r.status}
          </button>
          {editable && (
            <div className="flex gap-1">
              <button onClick={start} className="text-slate-400 hover:text-emerald-600 p-1" title="수정"><Pencil size={14} /></button>
              <button onClick={onDelete} className="text-slate-400 hover:text-rose-600 p-1" title="삭제"><Trash2 size={14} /></button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

/* 인수인계 — 같은 부서만 (사무실은 전체) */
function Handover({ me, rows, setRows, complaints }) {
  const [text, setText] = useState("");
  const canSee = (dept) => me.dept === "사무실" || me.dept === dept;
  const visible = rows.filter((r) => canSee(r.dept));
  const add = () => {
    if (!text.trim()) return;
    setRows([{ id: Date.now(), type: "인계 메모", text, by: me.id, dept: me.dept, at: nowLocal().replace("T", " "), hist: [], deleted: null }, ...rows]);
    setText("");
  };
  const saveEdit = (id, newText) => setRows(rows.map((r) => r.id === id ? { ...r, text: newText, hist: [...(r.hist || []), { by: me.id, at: stamp() }] } : r));
  const del = (id) => { if (window.confirm("이 인계 메모를 삭제할까요?")) setRows(rows.map((r) => r.id === id ? { ...r, deleted: { by: me.id, at: stamp() } } : r)); };
  return (
    <div className="space-y-4">
      <div className="bg-white rounded-xl border border-slate-200 p-4">
        <p className="text-xs font-semibold text-slate-500 mb-2">인계 메모 — 오늘 처리 못한 건, 공지, 익일 예약/행사 안내 등을 자유롭게 남기세요</p>
        <textarea rows={3} value={text} onChange={(e) => setText(e.target.value)} placeholder="다음 근무자에게 전달할 내용을 적으세요"
          className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 resize-none" />
        <button onClick={add} className="mt-2 flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2"><Send size={14} /> 등록</button>
      </div>
      <div className="space-y-2">
        {visible.map((r) => r.deleted
          ? <DeletedCard key={r.id} r={r} label="인계 메모" />
          : <MemoCard key={r.id} r={r} onSave={(t) => saveEdit(r.id, t)} onDelete={() => del(r.id)} />)}
        {!visible.length && <p className="text-center text-slate-400 text-sm py-6">표시할 인수인계가 없습니다.</p>}
      </div>
    </div>
  );
}
function MemoCard({ r, onSave, onDelete }) {
  const [edit, setEdit] = useState(false);
  const [t, setT] = useState(r.text);
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm">
      <div className="flex items-center gap-2 mb-1">
        <span className="text-xs px-2 py-0.5 rounded bg-sky-100 text-sky-700">{r.type}</span>
        <span className="text-xs text-slate-400 font-mono">{r.dept}</span>
        <div className="ml-auto flex gap-1">
          <button onClick={() => { setT(r.text); setEdit(!edit); }} className="text-slate-400 hover:text-emerald-600 p-1" title="수정"><Pencil size={13} /></button>
          <button onClick={onDelete} className="text-slate-400 hover:text-rose-600 p-1" title="삭제"><Trash2 size={13} /></button>
        </div>
      </div>
      {edit ? (
        <div>
          <textarea rows={2} value={t} onChange={(e) => setT(e.target.value)} className="w-full rounded-lg border border-slate-200 p-2 text-sm outline-none focus:border-emerald-500 resize-none" />
          <div className="flex gap-2 mt-2">
            <button onClick={() => { if (t.trim()) { onSave(t); setEdit(false); } }} className="rounded-lg bg-emerald-600 text-white text-xs font-medium px-3 py-1.5">저장</button>
            <button onClick={() => setEdit(false)} className="rounded-lg border border-slate-200 text-slate-500 text-xs px-3 py-1.5">취소</button>
          </div>
        </div>
      ) : <p className="text-sm">{r.text}</p>}
      <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.at} · {r.by}</p>
      <AuditLine r={r} />
    </div>
  );
}

/* 요청사항 — 작성 시 '사무실만 확인 가능' 체크로 공개범위 선택. 사무실은 항상 전체 열람·답변 */
function Requests({ me, rows, setRows }) {
  const [text, setText] = useState("");
  const [officeOnly, setOfficeOnly] = useState(true);
  const office = me.dept === "사무실";
  // 열람 규칙: 사무실 = 전체 열람 / 작성자 본인 = 항상 열람 / 그 외 = officeOnly가 아닌 요청만 열람
  const visible = office ? rows : rows.filter((r) => r.by === me.id || !r.officeOnly);
  const add = () => {
    if (!text.trim()) return;
    setRows([{ id: Date.now(), text, officeOnly, by: me.id, dept: me.dept, date: today(), replies: [], hist: [], deleted: null }, ...rows]);
    setText(""); setOfficeOnly(true);
  };
  const reply = (id, msg) => setRows(rows.map((r) => r.id === id ? { ...r, replies: [...r.replies, { text: msg, by: me.id, date: today() }] } : r));
  const saveEdit = (id, newText) => setRows(rows.map((r) => r.id === id ? { ...r, text: newText, hist: [...(r.hist || []), { by: me.id, at: stamp() }] } : r));
  const del = (id) => { if (window.confirm("이 요청을 삭제할까요?")) setRows(rows.map((r) => r.id === id ? { ...r, deleted: { by: me.id, at: stamp() } } : r)); };
  return (
    <div>
      <div className="bg-white rounded-xl border border-slate-200 p-4 mb-4">
        <p className="text-sm font-semibold mb-1">요청 작성</p>
        <textarea rows={2} value={text} onChange={(e) => setText(e.target.value)} placeholder="요청 내용을 적으세요"
          className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 resize-none" />
        <label className="flex items-center gap-2 mt-2 text-sm text-slate-600 cursor-pointer select-none">
          <input type="checkbox" checked={officeOnly} onChange={(e) => setOfficeOnly(e.target.checked)}
            className="w-4 h-4 rounded border-slate-300 accent-emerald-600" />
          사무실만 확인 가능하도록 하기
        </label>
        <p className="text-[11px] text-slate-400 mt-1">{officeOnly ? "체크됨 — 사무실 직원과 작성자 본인만 볼 수 있어요." : "체크 해제 — 다른 직원들도 이 요청을 볼 수 있어요."}</p>
        <button onClick={add} className="mt-2 flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2"><Send size={14} /> 요청 등록</button>
      </div>
      {office && <p className="text-xs text-emerald-700 mb-2 flex items-center gap-1"><Inbox size={13} /> 사무실 권한 · 모든 요청을 열람하고 답변할 수 있습니다.</p>}
      <div className="space-y-2">
        {visible.map((r) => r.deleted
          ? <DeletedCard key={r.id} r={r} label="요청" />
          : <RequestCard key={r.id} r={r} canReply={office} canEdit={r.by === me.id || office} onReply={(m) => reply(r.id, m)} onSave={(t) => saveEdit(r.id, t)} onDelete={() => del(r.id)} />)}
        {!visible.length && <p className="text-center text-slate-400 text-sm py-6">요청이 없습니다.</p>}
      </div>
    </div>
  );
}
function RequestCard({ r, canReply, canEdit, onReply, onSave, onDelete }) {
  const [msg, setMsg] = useState("");
  const [edit, setEdit] = useState(false);
  const [t, setT] = useState(r.text);
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
      <div className="flex items-center gap-2 mb-1">
        {r.officeOnly && <span className="text-[11px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">🔒 사무실만</span>}
        {canEdit && (
          <div className="ml-auto flex gap-1">
            <button onClick={() => { setT(r.text); setEdit(!edit); }} className="text-slate-400 hover:text-emerald-600 p-1" title="수정"><Pencil size={13} /></button>
            <button onClick={onDelete} className="text-slate-400 hover:text-rose-600 p-1" title="삭제"><Trash2 size={13} /></button>
          </div>
        )}
      </div>
      {edit ? (
        <div>
          <textarea rows={2} value={t} onChange={(e) => setT(e.target.value)} className="w-full rounded-lg border border-slate-200 p-2 text-sm outline-none focus:border-emerald-500 resize-none" />
          <div className="flex gap-2 mt-2">
            <button onClick={() => { if (t.trim()) { onSave(t); setEdit(false); } }} className="rounded-lg bg-emerald-600 text-white text-xs font-medium px-3 py-1.5">저장</button>
            <button onClick={() => setEdit(false)} className="rounded-lg border border-slate-200 text-slate-500 text-xs px-3 py-1.5">취소</button>
          </div>
        </div>
      ) : <p className="text-sm">{r.text}</p>}
      <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.date} · {r.by} ({r.dept})</p>
      <AuditLine r={r} />
      {r.replies.map((rp, i) => (
        <div key={i} className="mt-2 ml-3 pl-3 border-l-2 border-emerald-200">
          <p className="text-sm text-emerald-800">↳ {rp.text}</p>
          <p className="text-[11px] text-slate-400 font-mono">{rp.date} · {rp.by} (사무실)</p>
        </div>
      ))}
      {canReply && (
        <div className="flex gap-2 mt-3">
          <input value={msg} onChange={(e) => setMsg(e.target.value)} placeholder="답변 입력"
            className="flex-1 rounded-lg border border-slate-200 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
          <button onClick={() => { if (msg.trim()) { onReply(msg); setMsg(""); } }} className="rounded-lg bg-slate-800 text-white text-xs font-medium px-3">답변</button>
        </div>
      )}
    </div>
  );
}

/* ─────────────────────── 기록 검색 */
function RecordSearch({ complaints }) {
  const [q, setQ] = useState("");
  const [f, setF] = useState("전체");
  const list = useMemo(() => complaints.filter((r) => {
    if (r.deleted) return false;
    const catOk = f === "전체" || r.cat === f || r.status === f;
    const qOk = !q.trim() || (`${r.content} ${r.action} ${r.name} ${r.dong}`).toLowerCase().includes(q.toLowerCase());
    return catOk && qOk;
  }), [complaints, q, f]);
  return (
    <div>
      <div className="relative mb-3">
        <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
        <input value={q} onChange={(e) => setQ(e.target.value)} placeholder="비슷한 과거 사례 검색 (예: 온수, 환불, 골프장…)"
          className="w-full rounded-xl border border-slate-200 pl-9 pr-3 py-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100" />
      </div>
      <div className="flex gap-1.5 mb-3 flex-wrap">
        {["전체", ...민원카테고리, "진행 중", "완료"].map((c) => (
          <button key={c} onClick={() => setF(c)} className={`text-xs font-medium px-3 py-1.5 rounded-full ${f === c ? "bg-emerald-600 text-white" : "bg-white border border-slate-200 text-slate-500"}`}>{c}</button>
        ))}
      </div>
      <p className="text-xs text-slate-400 mb-2">{list.length}건</p>
      <div className="space-y-2">
        {list.length ? list.map((r) => <ComplaintCard key={r.id} r={r} onToggle={() => {}} />)
          : <div className="text-center text-slate-400 text-sm py-10">검색 결과가 없습니다.</div>}
      </div>
    </div>
  );
}

/* ─────────────────────── 분실물 */
function Lost({ me, rows, setRows }) {
  const [open, setOpen] = useState(false);
  const [v, setV] = useState({});
  const [editId, setEditId] = useState(null);
  const [editV, setEditV] = useState({});

  const add = () => {
    if (!(v.item || "").trim()) return;
    setRows([{ id: Date.now(), item: v.item, place: v.place || "", found: today(), by: me.id, status: "보관중" }, ...rows]);
    setV({}); setOpen(false);
  };
  const toggle = (id) => setRows(rows.map((r) => r.id === id ? { ...r, status: r.status === "보관중" ? "반환완료" : "보관중" } : r));
  const startEdit = (r) => { setEditId(r.id); setEditV({ item: r.item, place: r.place }); };
  const saveEdit = (id) => {
    if (!editV.item.trim()) return;
    setRows(rows.map((r) => r.id === id ? { ...r, item: editV.item, place: editV.place } : r));
    setEditId(null);
  };
  const del = (id) => { if (askYesNo("이 분실물 기록을 삭제할까요?")) setRows(rows.filter((r) => r.id !== id)); };

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-base font-bold">분실물 <span className="text-slate-400 font-normal text-sm">({rows.length})</span></h2>
        <button onClick={() => setOpen(true)} className="flex items-center gap-1 rounded-lg bg-slate-800 text-white text-xs font-medium px-3 py-2"><Plus size={14} /> 등록</button>
      </div>
      {open && (
        <div className="bg-white rounded-xl border border-emerald-200 ring-2 ring-emerald-50 p-4 mb-3 space-y-2.5">
          <div className="flex items-center justify-between"><span className="text-sm font-semibold">분실물 등록</span><button onClick={() => setOpen(false)} className="text-slate-400"><X size={16} /></button></div>
          <input placeholder="물품명" value={v.item || ""} onChange={(e) => setV({ ...v, item: e.target.value })} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" />
          <input placeholder="습득 장소" value={v.place || ""} onChange={(e) => setV({ ...v, place: e.target.value })} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" />
          <button onClick={add} className="w-full rounded-lg bg-emerald-600 text-white text-sm font-medium py-2">등록</button>
        </div>
      )}
      <div className="space-y-2">
        {rows.map((r) => (
          <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm">
            {editId === r.id ? (
              <div className="space-y-2">
                <input value={editV.item} onChange={(e) => setEditV({ ...editV, item: e.target.value })} className="w-full rounded-lg border border-emerald-300 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
                <input value={editV.place} onChange={(e) => setEditV({ ...editV, place: e.target.value })} placeholder="습득 장소" className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
                <div className="flex gap-2">
                  <button onClick={() => saveEdit(r.id)} className="flex-1 rounded-lg bg-emerald-600 text-white text-xs font-medium py-1.5">저장</button>
                  <button onClick={() => setEditId(null)} className="flex-1 rounded-lg border border-slate-200 text-slate-500 text-xs py-1.5">취소</button>
                </div>
              </div>
            ) : (
              <div className="flex items-start justify-between gap-3">
                <div className="min-w-0">
                  <p className="font-medium truncate">{r.item}</p>
                  <p className="text-xs text-slate-500 mt-0.5">{r.place} · 습득 <span className="font-mono">{r.found}</span></p>
                  <p className="text-xs text-emerald-600 mt-1">👤 {r.by}</p>
                </div>
                <div className="flex items-center gap-1 shrink-0">
                  <button onClick={() => toggle(r.id)} className={`rounded-full px-2.5 py-1 text-xs font-medium ${r.status === "보관중" ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"}`}>{r.status}</button>
                  <button onClick={() => startEdit(r)} className="text-slate-400 hover:text-emerald-600 p-1" title="수정"><Pencil size={14} /></button>
                  <button onClick={() => del(r.id)} className="text-slate-400 hover:text-rose-600 p-1" title="삭제"><Trash2 size={14} /></button>
                </div>
              </div>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────── 주요 행사·일정 */
function Events({ me, rows, setRows }) {
  const office = me.dept === "사무실";
  const [open, setOpen] = useState(false);
  const [v, setV] = useState({ title: "", date: today(), text: "", file: null });
  const [editId, setEditId] = useState(null);
  const [editV, setEditV] = useState({});

  const onFile = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => setV((s) => ({ ...s, file: { name: f.name, type: f.type, url: ev.target.result } }));
    reader.readAsDataURL(f);
  };
  const onEditFile = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => setEditV((s) => ({ ...s, file: { name: f.name, type: f.type, url: ev.target.result } }));
    reader.readAsDataURL(f);
  };
  const add = () => { if (!v.title.trim()) return; setRows([{ id: Date.now(), ...v, by: me.id }, ...rows]); setV({ title: "", date: today(), text: "", file: null }); setOpen(false); };
  const startEdit = (r) => { setEditId(r.id); setEditV({ title: r.title, date: r.date, text: r.text, file: r.file || null }); };
  const saveEdit = (id) => {
    if (!editV.title.trim()) return;
    setRows(rows.map((r) => r.id === id ? { ...r, ...editV } : r));
    setEditId(null);
  };
  const del = (id) => { if (askYesNo("이 행사·일정을 삭제할까요?")) setRows(rows.filter((r) => r.id !== id)); };

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-base font-bold">주요 행사·일정 <span className="text-slate-400 font-normal text-sm">({rows.length})</span></h2>
        {office
          ? <button onClick={() => setOpen(true)} className="flex items-center gap-1 rounded-lg bg-slate-800 text-white text-xs font-medium px-3 py-2"><Plus size={14} /> 행사 등록</button>
          : <span className="text-xs text-slate-400">사무실만 등록 가능</span>}
      </div>
      {open && office && (
        <div className="bg-white rounded-xl border border-emerald-200 ring-2 ring-emerald-50 p-4 mb-4 space-y-2.5">
          <div className="flex items-center justify-between"><span className="text-sm font-semibold">행사·일정 등록</span><button onClick={() => setOpen(false)} className="text-slate-400"><X size={16} /></button></div>
          <input placeholder="제목" value={v.title} onChange={(e) => setV({ ...v, title: e.target.value })} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" />
          <input type="date" value={v.date} onChange={(e) => setV({ ...v, date: e.target.value })} className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" />
          <textarea rows={3} placeholder="내용" value={v.text} onChange={(e) => setV({ ...v, text: e.target.value })} className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 resize-none" />
          <label className="flex items-center gap-2 text-sm text-slate-500 cursor-pointer">
            <span className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 hover:bg-slate-50"><Upload size={14} /> 사진·PDF 첨부</span>
            <input type="file" accept=".jpg,.jpeg,.png,.pdf" onChange={onFile} className="hidden" />
            {v.file && <span className="text-xs text-emerald-600 truncate">{v.file.name}</span>}
          </label>
          <button onClick={add} className="w-full rounded-lg bg-emerald-600 text-white text-sm font-medium py-2">등록</button>
        </div>
      )}
      <div className="space-y-3">
        {rows.map((r) => (
          <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
            {editId === r.id ? (
              <div className="space-y-2">
                <input value={editV.title} onChange={(e) => setEditV({ ...editV, title: e.target.value })} className="w-full rounded-lg border border-emerald-300 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
                <input type="date" value={editV.date} onChange={(e) => setEditV({ ...editV, date: e.target.value })} className="w-full rounded-lg border border-slate-200 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
                <textarea rows={3} value={editV.text} onChange={(e) => setEditV({ ...editV, text: e.target.value })} className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 resize-none" />
                <label className="flex items-center gap-2 text-sm text-slate-500 cursor-pointer">
                  <span className="inline-flex items-center gap-1.5 rounded-lg border border-slate-200 px-3 py-1.5 hover:bg-slate-50"><Upload size={14} /> 첨부 교체</span>
                  <input type="file" accept=".jpg,.jpeg,.png,.pdf" onChange={onEditFile} className="hidden" />
                  {editV.file && <span className="text-xs text-emerald-600 truncate">{editV.file.name}</span>}
                </label>
                <div className="flex gap-2">
                  <button onClick={() => saveEdit(r.id)} className="flex-1 rounded-lg bg-emerald-600 text-white text-xs font-medium py-1.5">저장</button>
                  <button onClick={() => setEditId(null)} className="flex-1 rounded-lg border border-slate-200 text-slate-500 text-xs py-1.5">취소</button>
                </div>
              </div>
            ) : (
              <>
                <div className="flex items-center gap-2 mb-1">
                  <CalendarDays size={15} className="text-emerald-600" />
                  <span className="font-semibold">{r.title}</span>
                  <span className="text-xs text-slate-400 font-mono ml-auto">{r.date}</span>
                  {office && (
                    <>
                      <button onClick={() => startEdit(r)} className="text-slate-400 hover:text-emerald-600 p-1" title="수정"><Pencil size={14} /></button>
                      <button onClick={() => del(r.id)} className="text-slate-400 hover:text-rose-600 p-1" title="삭제"><Trash2 size={14} /></button>
                    </>
                  )}
                </div>
                {r.text && <p className="text-sm text-slate-600 mt-1 whitespace-pre-wrap">{r.text}</p>}
                {r.file && (r.file.type.includes("image")
                  ? <img src={r.file.url} alt="" className="mt-2 rounded-lg border border-slate-200 max-h-56" />
                  : <a href={r.file.url} download={r.file.name} className="mt-2 inline-flex items-center gap-1.5 text-sm text-sky-600 bg-sky-50 rounded-lg px-3 py-1.5"><FileText size={14} /> {r.file.name}</a>)}
                <p className="text-[11px] text-slate-400 mt-2 font-mono">등록 · {r.by}</p>
              </>
            )}
          </div>
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────── 홈 (로그인 첫 화면) */
function Home({ me, events }) {
  const isAI = me.id === AI_ADMIN;
  const [stats, setStats] = useState([]);
  const [loading, setLoading] = useState(true);
  const [job, setJob] = useState(null);
  const [view, setView] = useState(null);      // null | "under" | "over" | "note"
  const [noteId, setNoteId] = useState(null);
  const [noteDraft, setNoteDraft] = useState("");
  const [adding, setAdding] = useState(false);
  const [pick, setPick] = useState("");        // 특이사항 등록용 강좌 검색어
  const [picked, setPicked] = useState(null);
  const [newNote, setNewNote] = useState("");

  const load = React.useCallback(async () => {
    const { data } = await supabase.from("course_stats").select("*").order("course");
    setStats(data || []);
    setLoading(false);
  }, []);
  React.useEffect(() => { load(); }, [load]);

  React.useEffect(() => {
    if (!job || job.status === "done" || job.status === "error") return;
    const t = setInterval(async () => {
      const { data } = await supabase.from("roster_jobs").select("*").eq("id", job.id).maybeSingle();
      if (data) { setJob(data); if (data.status === "done") load(); }
    }, 3000);
    return () => clearInterval(t);
  }, [job, load]);

  const requestUpdate = async () => {
    const { data, error } = await supabase.from("roster_jobs")
      .insert({ status: "pending", by: me.id, at: stamp() }).select().maybeSingle();
    if (error) { alert("요청 실패: " + error.message); return; }
    setJob(data);
  };

  const saveNote = async (course, text) => {
    const { error } = await supabase.from("course_stats").update({ note: text }).eq("course", course);
    if (error) { alert("저장 실패: " + error.message); return false; }
    setStats((s) => s.map((x) => x.course === course ? { ...x, note: text } : x));
    return true;
  };

  const under = stats.filter(isUnder);
  const over = stats.filter(isOver);
  const noted = stats.filter((s) => (s.note || "").trim());
  const synced = stats.find((s) => s.synced_at)?.synced_at;
  const upcoming = [...(events || [])]
    .filter((e) => e.date >= today()).sort((a, b) => a.date.localeCompare(b.date)).slice(0, 5);

  const cards = [
    { id: "under", label: "정원 미달 강좌", value: under.length, color: "emerald" },
    { id: "over", label: "정원 초과 강좌", value: over.length, color: "rose" },
    { id: "note", label: "강습 특이사항", value: noted.length, color: "amber" },
  ];
  const shown = view === "under" ? under : view === "over" ? over : view === "note" ? noted : [];

  // 특이사항 등록용 강좌 검색 (수집된 강좌 우선, 없으면 전체 강좌 목록)
  const pool = [...new Set([...stats.map((s) => s.course), ...ALL_COURSES.map((c) => c.full)])];
  const found = pick.trim()
    ? pool.filter((c) => c.replace(/\s/g, "").includes(pick.replace(/\s/g, ""))).slice(0, 8)
    : [];

  const addNote = async () => {
    if (!picked || !newNote.trim()) return;
    const exists = stats.some((s) => s.course === picked);
    if (exists) {
      if (!(await saveNote(picked, newNote.trim()))) return;
    } else {
      const { error } = await supabase.from("course_stats")
        .upsert({ course: picked, note: newNote.trim(), enrolled: 0, capacity: 0 });
      if (error) { alert("저장 실패: " + error.message); return; }
      await load();
    }
    setAdding(false); setPick(""); setPicked(null); setNewNote("");
    setView("note");
  };

  return (
    <div className="space-y-5">
      <div className="flex items-start justify-between gap-2">
        <div>
          <h2 className="text-lg font-bold">안녕하세요, {me.id}님</h2>
          <p className="text-sm text-slate-500">{today()} · {me.dept}</p>
        </div>
        {isAI && (
          <button onClick={requestUpdate} disabled={job && (job.status === "pending" || job.status === "working")}
            className="shrink-0 flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-medium px-3 py-2">
            {job && (job.status === "pending" || job.status === "working")
              ? <><Loader2 size={13} className="animate-spin" /> 가져오는 중…</> : "수강인원 업데이트"}
          </button>
        )}
      </div>

      {/* 요약 카드 — 누르면 목록이 펼쳐집니다 */}
      <div className="grid grid-cols-3 gap-2">
        {cards.map((c) => {
          const on = view === c.id;
          const tone = {
            emerald: on ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-200 hover:border-emerald-300",
            rose: on ? "bg-rose-600 border-rose-600 text-white" : "bg-white border-slate-200 hover:border-rose-300",
            amber: on ? "bg-amber-500 border-amber-500 text-white" : "bg-white border-slate-200 hover:border-amber-300",
          }[c.color];
          return (
            <button key={c.id} onClick={() => setView(on ? null : c.id)}
              className={`rounded-xl border p-3 text-center transition ${tone}`}>
              <p className={`text-[11px] leading-tight ${on ? "opacity-90" : "text-slate-500"}`}>{c.label}</p>
              <p className="text-2xl font-bold mt-0.5">{c.value}</p>
            </button>
          );
        })}
      </div>
      <p className="text-[11px] text-slate-400 -mt-3">
        {synced ? `마지막 업데이트 ${synced} · 이후 등록·취소는 자동 반영` : "아직 업데이트한 적이 없습니다"}
        {!isAI && " (업데이트는 foreon4 계정)"}
      </p>

      {/* 선택한 카드의 목록 */}
      {view && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4">
          <div className="flex items-center justify-between mb-2">
            <h3 className="font-bold text-sm">{cards.find((c) => c.id === view).label}</h3>
            <div className="flex items-center gap-2">
              {view === "note" && (
                <button onClick={() => setAdding(!adding)} className="flex items-center gap-1 rounded-lg bg-slate-800 text-white text-xs px-2.5 py-1.5">
                  <Plus size={12} /> 등록
                </button>
              )}
              <button onClick={() => setView(null)} className="text-slate-400"><X size={16} /></button>
            </div>
          </div>

          {/* 특이사항 등록 — 강좌 검색 후 선택 */}
          {view === "note" && adding && (
            <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 mb-3 space-y-2">
              <input value={pick} onChange={(e) => { setPick(e.target.value); setPicked(null); }}
                placeholder="강좌명 검색 (예: 요가 수금, 성인수영 화목)" autoFocus
                className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-amber-500" />
              {!picked && found.map((c) => (
                <button key={c} onClick={() => { setPicked(c); setPick(c); }}
                  className="block w-full text-left text-sm bg-white border border-slate-200 rounded-lg px-3 py-2 hover:border-amber-400">{c}</button>
              ))}
              {!picked && pick.trim() && !found.length && <p className="text-xs text-slate-400">검색 결과가 없습니다.</p>}
              {picked && (
                <>
                  <p className="text-xs text-amber-800">선택: <b>{picked}</b></p>
                  <textarea rows={2} value={newNote} onChange={(e) => setNewNote(e.target.value)}
                    placeholder="특이사항 내용" className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-amber-500 resize-none" />
                  <div className="flex gap-2">
                    <button onClick={addNote} className="flex-1 rounded-lg bg-amber-500 text-white text-xs font-medium py-2">등록</button>
                    <button onClick={() => { setAdding(false); setPick(""); setPicked(null); setNewNote(""); }}
                      className="flex-1 rounded-lg border border-slate-200 text-slate-500 text-xs py-2">취소</button>
                  </div>
                </>
              )}
            </div>
          )}

          {loading ? <p className="text-sm text-slate-400 py-6 text-center">불러오는 중…</p>
            : !shown.length ? <p className="text-sm text-slate-400 py-6 text-center">해당하는 강좌가 없습니다.</p>
              : (
                <div className="space-y-1.5 max-h-[26rem] overflow-y-auto">
                  {shown.map((s) => (
                    <div key={s.course} className="border border-slate-100 rounded-lg px-3 py-2">
                      <div className="flex items-start justify-between gap-2">
                        <span className="text-sm leading-tight">{s.course}</span>
                        <span className={`shrink-0 text-xs font-semibold px-2 py-0.5 rounded-full ${
                          isOver(s) ? "bg-rose-100 text-rose-700"
                            : isUnder(s) ? "bg-amber-100 text-amber-700"
                              : "bg-emerald-50 text-emerald-700"}`}>
                          {s.enrolled}/{capacityOf(s.course)}
                        </span>
                      </div>
                      {noteId === s.course ? (
                        <div className="flex gap-1 mt-1.5">
                          <input value={noteDraft} onChange={(e) => setNoteDraft(e.target.value)} autoFocus
                            className="flex-1 rounded border border-amber-300 px-2 py-1 text-xs outline-none" />
                          <button onClick={async () => { if (await saveNote(s.course, noteDraft)) setNoteId(null); }}
                            className="text-emerald-600 text-xs font-medium px-1">저장</button>
                          <button onClick={() => setNoteId(null)} className="text-slate-400 text-xs px-1">취소</button>
                        </div>
                      ) : (
                        <button onClick={() => { setNoteId(s.course); setNoteDraft(s.note || ""); }}
                          className="mt-1 text-xs text-left w-full">
                          {s.note ? <span className="text-amber-700">📌 {s.note}</span> : <span className="text-slate-300">+ 특이사항 입력</span>}
                        </button>
                      )}
                    </div>
                  ))}
                </div>
              )}
        </div>
      )}

      {/* 다가오는 행사 */}
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5">
        <h3 className="font-bold flex items-center gap-1.5 mb-3"><CalendarDays size={17} className="text-emerald-600" /> 다가오는 행사·일정</h3>
        {upcoming.length ? (
          <div className="space-y-2">
            {upcoming.map((e) => (
              <div key={e.id} className="flex items-start gap-3 border-b border-slate-50 pb-2 last:border-0">
                <span className="text-xs font-mono bg-emerald-50 text-emerald-700 rounded px-2 py-1 shrink-0">{e.date?.slice(5)}</span>
                <div className="min-w-0">
                  <p className="text-sm font-medium">{e.title}</p>
                  {e.text && <p className="text-xs text-slate-500 truncate">{e.text}</p>}
                </div>
              </div>
            ))}
          </div>
        ) : <p className="text-sm text-slate-400 py-4 text-center">예정된 행사가 없습니다.</p>}
      </div>
    </div>
  );
}

/* 일일업무일지 자동입력 — 커뮤니티이력(A) 내용을 업무일지(B)에 채워 넣기 */
const fileToB64 = (file) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = (e) => res(e.target.result.split(",")[1]);
  r.onerror = () => rej(new Error("파일을 읽지 못했습니다"));
  r.readAsDataURL(file);
});

function DailyExcel({ me }) {
  const [a, setA] = useState(null);
  const [b, setB] = useState(null);
  const [day, setDay] = useState(today());
  const [job, setJob] = useState(null);
  const [busy, setBusy] = useState(false);
  const [open, setOpen] = useState(true);
  const [notes, setNotes] = useState("");
  const [saved, setSaved] = useState(null);     // { by, at }
  const [dirty, setDirty] = useState(false);
  const [savingNote, setSavingNote] = useState(false);

  // 날짜를 바꾸면 그 날짜에 저장된 업무내용을 불러온다 (오전반이 적어둔 내용)
  React.useEffect(() => {
    let alive = true;
    supabase.from("daily_notes").select("*").eq("day", day).maybeSingle()
      .then(({ data }) => {
        if (!alive) return;
        setNotes(data?.text || "");
        setSaved(data ? { by: data.by, at: data.at } : null);
        setDirty(false);
      });
    return () => { alive = false; };
  }, [day]);

  const saveNotes = async () => {
    setSavingNote(true);
    const at = stamp();
    const { error } = await supabase.from("daily_notes")
      .upsert({ day, text: notes, by: me.id, at });
    setSavingNote(false);
    if (error) { alert("저장 실패: " + error.message); return; }
    setSaved({ by: me.id, at });
    setDirty(false);
  };

  React.useEffect(() => {
    if (!job || job.status === "done" || job.status === "error") return;
    const t = setInterval(async () => {
      const { data } = await supabase.from("excel_jobs").select("*").eq("id", job.id).maybeSingle();
      if (data) setJob(data);
    }, 3000);
    return () => clearInterval(t);
  }, [job]);

  const run = async () => {
    if (!a || !b) { alert("커뮤니티이력(A)과 업무일지(B) 파일을 모두 선택해주세요."); return; }
    if (dirty && !askYesNo("저장하지 않은 업무내용이 있습니다. 지금 화면의 내용으로 진행할까요?")) return;
    setBusy(true);
    try {
      const [fa, fb] = await Promise.all([fileToB64(a), fileToB64(b)]);
      const { data, error } = await supabase.from("excel_jobs").insert({
        day, name_a: a.name, name_b: b.name, file_a: fa, file_b: fb,
        notes,
        status: "pending", by: me.id, at: stamp(),
      }).select().maybeSingle();
      if (error) throw new Error(error.message);
      setJob(data);
    } catch (e) {
      alert("요청 실패: " + e.message);
    }
    setBusy(false);
  };

  const download = () => {
    const bin = atob(job.result_file);
    const buf = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) buf[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    const link = document.createElement("a");
    link.href = url; link.download = job.result_name || "업무일지.xlsx";
    link.click(); URL.revokeObjectURL(url);
  };

  const running = job && (job.status === "pending" || job.status === "working");

  return (
    <div className="bg-white rounded-xl border border-slate-200 mb-4 overflow-hidden">
      <button onClick={() => setOpen(!open)} className="w-full flex items-center justify-between px-4 py-3 hover:bg-slate-50">
        <span className="text-sm font-semibold flex items-center gap-1.5"><FileText size={15} className="text-emerald-600" /> 일일업무일지 자동입력</span>
        <span className="text-xs text-slate-400">{open ? "닫기" : "열기"}</span>
      </button>
      {open && (
        <div className="px-4 pb-4 space-y-3 border-t border-slate-100 pt-3">
          <p className="text-[11px] text-slate-500 leading-relaxed">
            바이비 <b>커뮤니티이력</b> 엑셀(A)과 <b>월별 업무일지</b> 엑셀(B)을 올리면, 해당 날짜의 수입현황이 B에 채워진 새 파일을 받습니다. 원본은 변경되지 않습니다.
          </p>
          <label className="block">
            <span className="text-xs font-medium text-slate-500 block mb-1">① 커뮤니티이력 엑셀 (A)</span>
            <input type="file" accept=".xlsx,.xls" onChange={(e) => setA(e.target.files?.[0] || null)}
              className="w-full text-xs file:mr-2 file:rounded-lg file:border-0 file:bg-slate-100 file:px-3 file:py-1.5 file:text-xs" />
          </label>
          <label className="block">
            <span className="text-xs font-medium text-slate-500 block mb-1">② 업무일지 엑셀 (B) — 내용이 채워질 파일</span>
            <input type="file" accept=".xlsx,.xlsm" onChange={(e) => setB(e.target.files?.[0] || null)}
              className="w-full text-xs file:mr-2 file:rounded-lg file:border-0 file:bg-slate-100 file:px-3 file:py-1.5 file:text-xs" />
          </label>
          <label className="block">
            <span className="text-xs font-medium text-slate-500 block mb-1">③ 처리할 날짜</span>
            <input type="date" value={day} onChange={(e) => setDay(e.target.value)}
              className="rounded-lg border border-slate-200 px-3 py-1.5 text-sm outline-none focus:border-emerald-500" />
          </label>
          <div className="bg-slate-50 border border-slate-200 rounded-xl p-3">
            <span className="text-xs font-medium text-slate-600 block mb-1">④ 업무내용 (엑셀 B77~B87 칸에 그대로 입력)</span>
            <p className="text-[11px] text-slate-400 mb-1.5">오전반이 적고 저장해두면, 오후반이 같은 날짜를 열었을 때 그대로 이어받습니다.</p>
            <textarea rows={6} value={notes}
              onChange={(e) => { setNotes(e.target.value); setDirty(true); }}
              placeholder={"- 업무내용1\n- 업무내용2\n- 업무내용3"}
              className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 resize-none font-mono" />
            <div className="flex items-center justify-between mt-1.5 gap-2">
              <p className="text-[11px] text-slate-400">
                한 줄이 한 칸 · {notes.split("\n").filter((l) => l.trim()).length}줄 (최대 11줄)
                {saved && <span className="text-emerald-600"> · 저장됨 {saved.at} ({saved.by})</span>}
                {dirty && <span className="text-amber-600"> · 저장 안 됨</span>}
              </p>
              <button onClick={saveNotes} disabled={savingNote || !dirty}
                className="shrink-0 rounded-lg bg-slate-800 hover:bg-slate-900 disabled:opacity-40 text-white text-xs font-medium px-3 py-1.5">
                {savingNote ? "저장 중…" : "업무내용 저장"}
              </button>
            </div>
          </div>
          <button onClick={run} disabled={busy || running}
            className="w-full rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-sm font-medium py-2.5 flex items-center justify-center gap-1.5">
            {busy || running ? <><Loader2 size={15} className="animate-spin" /> 처리 중… (PC 작업자 실행)</> : <><PlayCircle size={15} /> 자동 입력 실행</>}
          </button>

          {job && job.status === "error" && (
            <p className="text-xs text-rose-600 bg-rose-50 rounded-lg p-2.5 whitespace-pre-wrap">실패: {job.error}</p>
          )}
          {job && job.status === "done" && (
            <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-3 space-y-2">
              <p className="text-sm text-emerald-800">✔ 완료 — {job.result_name}</p>
              <button onClick={download} className="w-full rounded-lg bg-emerald-600 text-white text-sm font-medium py-2">파일 다운로드</button>
              {job.report && <details><summary className="text-xs text-emerald-700 cursor-pointer">검증 리포트 보기</summary>
                <pre className="text-[10px] text-slate-600 whitespace-pre-wrap mt-1 max-h-48 overflow-auto">{job.report}</pre></details>}
            </div>
          )}
        </div>
      )}
    </div>
  );
}

/* ─────────────────────── 사무실 업무일지 (사무실 계정 전용) */
function OfficeLog({ me }) {
  return (
    <div>
      <h2 className="text-base font-bold mb-3">사무실 업무일지</h2>
      <DailyExcel me={me} />
    </div>
  );
}

/* ─────────────────────── 강습 등록·취소 */
/* work_logs 에 작업 기록 남기기 (강습·독서실·골프백 공통) */
async function logWork(row) {
  try { await supabase.from("work_logs").insert(row); } catch (e) { /* 기록 실패는 조용히 무시 */ }
}

/* ─────────────────────── 등록·취소 (강습 / 독서실 / 골프백) */
function Registrations({ me }) {
  const [sub, setSub] = useState("강습");
  const subs = ["강습", "독서실", "골프백", "이월"];
  return (
    <div>
      <div className="inline-flex gap-1 p-1 bg-slate-100 rounded-xl mb-4">
        {subs.map((s) => (
          <button key={s} onClick={() => setSub(s)}
            className={`rounded-lg px-4 py-1.5 text-sm font-medium transition ${sub === s ? "bg-white shadow-sm text-emerald-700" : "text-slate-500"}`}>
            {s}
          </button>
        ))}
      </div>
      {sub === "강습" && <LessonForm me={me} />}
      {sub === "독서실" && <ReservationForm me={me} kind="독서실" options={[["개인", "독서실_개인"], ["일반", "독서실_일반"]]} seatLabel="좌석 번호" seatPh="예: 02" />}
      {sub === "골프백" && <ReservationForm me={me} kind="골프백" options={[["상단", "골프백_상단"], ["하단", "골프백_하단"]]} seatLabel="락커 번호" seatPh="예: 007" />}
      {sub === "이월" && <TransferForm me={me} />}
    </div>
  );
}

/* ─────────────────────── 이월 (지난달 수강생 → 다음달 창구) */
const TRANSFER_GROUPS = [
  "2단지 요가", "2단지 줌바", "2단지 타바타", "2단지 방송댄스", "2단지 근력",
  "수영", "아쿠아", "2단지 축구", "2단지 농구",
];

function TransferForm({ me }) {
  const [group, setGroup] = useState("");
  const [fromYm, setFromYm] = useState(ymNow());
  const [toYm, setToYm] = useState(ymList()[1]);
  const [excludes, setExcludes] = useState([]);
  const [file, setFile] = useState(null);
  const [safe, setSafe] = useState(true);
  const [saving, setSaving] = useState(false);
  const [job, setJob] = useState(null);
  const [err, setErr] = useState("");
  const [counts, setCounts] = useState({});
  const [confirming, setConfirming] = useState(false);

  // 강좌별 현재 수강 인원 (홈 화면과 같은 자료)
  const [statsMsg, setStatsMsg] = useState("인원 불러오는 중…");
  const keyOf = (t) => String(t).replace(/[\s\-()/·.,]/g, "");
  const loadCounts = React.useCallback(async () => {
    setStatsMsg("인원 불러오는 중…");
    const { data, error } = await supabase.from("course_stats").select("course,enrolled");
    if (error) { setStatsMsg("인원을 불러오지 못했습니다: " + error.message); return; }
    const map = {};
    for (const r of data || []) map[keyOf(r.course)] = r.enrolled;
    setCounts(map);
    setStatsMsg(Object.keys(map).length
      ? ""
      : "인원 숫자가 없습니다 — 홈 화면에서 [수강인원 업데이트] 를 먼저 눌러주세요.");
  }, []);
  React.useEffect(() => { loadCounts(); }, [loadCounts]);
  const countOf = (full) => counts[keyOf(full)];

  const isSwim = group === "수영" || group === "아쿠아";
  const flat = (t) => String(t).replace(/[\s\-()/·.,]/g, "");
  const words = group.trim().split(/\s+/).map(flat).filter(Boolean);
  const targets = words.length
    ? ALL_COURSES.filter((c) => words.every((w) => flat(c.full).includes(w)))
    : [];
  const willRun = targets.filter((c) => !excludes.includes(c.full));

  React.useEffect(() => {
    if (!job || job.status === "done" || job.status === "error") return;
    const t = setInterval(async () => {
      if (job.id) {
        const { data } = await supabase.from("transfer_jobs").select("*").eq("id", job.id).maybeSingle();
        if (data) setJob(data);
      } else {
        // id 를 못 받은 경우: 내가 보낸 가장 최근 요청을 따라간다
        const { data } = await supabase.from("transfer_jobs")
          .select("*").eq("by", me.id).order("id", { ascending: false }).limit(1);
        if (data && data[0]) setJob(data[0]);
      }
    }, 3000);
    return () => clearInterval(t);
  }, [job, me.id]);

  const toggle = (full) =>
    setExcludes((x) => x.includes(full) ? x.filter((v) => v !== full) : [...x, full]);

  const ask = () => {
    setErr("");
    if (!group) { setErr("강좌 그룹을 먼저 선택해주세요."); return; }
    if (!willRun.length) { setErr("이월할 강좌가 없습니다."); return; }
    if (fromYm === toYm) { setErr("가져올 달과 넣을 달이 같습니다."); return; }
    setConfirming(true);
  };

  const run = async () => {
    setErr("");
    setConfirming(false);
    setSaving(true);
    try {
      let b64 = null;
      if (file) b64 = await fileToB64(file);
      const base = {
        group_key: group, from_ym: fromYm, to_ym: toYm,
        excludes, change_file: b64, change_name: file?.name || null,
        safe, status: "pending", by: me.id, at: stamp(),
      };
      const targets = willRun.map((c) => c.full);
      let { data, error } = await supabase.from("transfer_jobs")
        .insert({ ...base, targets }).select().maybeSingle();

      // targets 칸이 아직 없는 계정이면 그 칸만 빼고 다시 저장
      if (error && /targets/i.test(error.message || "")) {
        ({ data, error } = await supabase.from("transfer_jobs")
          .insert(base).select().maybeSingle());
        if (!error) setErr("참고: 대상 강좌 저장 칸이 없어 PC가 직접 검색합니다 (7차 SQL 실행 권장).");
      }
      if (error) throw new Error(error.message);

      // 저장은 됐는데 응답이 비는 경우(읽기 권한 등) → 방금 넣은 요청을 다시 찾아본다
      if (!data) {
        const { data: found } = await supabase.from("transfer_jobs")
          .select("*").eq("by", me.id).eq("at", base.at)
          .order("id", { ascending: false }).limit(1);
        data = (found && found[0]) || null;
      }
      if (data) setJob(data);
      else setJob({ id: null, status: "working", progress: "요청 전송됨 — PC 작업자가 처리 중입니다" });
    } catch (e) {
      setErr("요청 실패: " + (e.message || String(e)));
    }
    setSaving(false);
  };

  const running = job && (job.status === "pending" || job.status === "working");

  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
      <div>
        <span className="text-xs font-medium text-slate-500 block mb-1.5">강좌 그룹 선택</span>
        <div className="flex flex-wrap gap-1.5">
          {TRANSFER_GROUPS.map((g) => (
            <button key={g} type="button" onClick={() => { setGroup(g); setExcludes([]); if (g !== "수영" && g !== "아쿠아") setFile(null); }}
              className={`text-sm px-3 py-1.5 rounded-full border transition ${group === g ? "bg-emerald-600 border-emerald-600 text-white font-medium" : "bg-white border-slate-200 text-slate-600 hover:border-emerald-300"}`}>
              {g}
            </button>
          ))}
        </div>
      </div>

      <div className="grid grid-cols-2 gap-2">
        <label className="block">
          <span className="text-xs font-medium text-slate-500 block mb-1">가져올 달 (기존 수강생)</span>
          <select value={fromYm} onChange={(e) => setFromYm(e.target.value)}
            className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500 bg-white">
            {ymList(6).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </label>
        <label className="block">
          <span className="text-xs font-medium text-slate-500 block mb-1">넣을 달 (이월 대상)</span>
          <select value={toYm} onChange={(e) => setToYm(e.target.value)}
            className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500 bg-white">
            {ymList(6).map((y) => <option key={y} value={y}>{y}</option>)}
          </select>
        </label>
      </div>

      {isSwim && (
        <label className="block bg-amber-50 border border-amber-200 rounded-lg p-3">
          <span className="text-xs font-medium text-amber-900 block mb-1">반변경 명단 엑셀 (선택)</span>
          <p className="text-[11px] text-amber-700 mb-1.5">강사명 · 시간 · 변경내용(기존 → 변경) · 회원명 · 동·호수 순서의 파일</p>
          <input type="file" accept=".xlsx,.xls" onChange={(e) => setFile(e.target.files?.[0] || null)}
            className="w-full text-xs file:mr-2 file:rounded-lg file:border-0 file:bg-white file:px-3 file:py-1.5 file:text-xs" />
          {file && <span className="text-xs text-emerald-700 block mt-1">✔ {file.name}</span>}
        </label>
      )}

      {group && (
        <div>
          <div className="flex items-center justify-between mb-1.5">
            <span className="text-xs font-medium text-slate-500">
              대상 강좌 {willRun.length}개 {excludes.length > 0 && <span className="text-rose-500">(제외 {excludes.length}개)</span>}
            </span>
            <div className="flex items-center gap-2">
              <button type="button" onClick={loadCounts} className="text-[11px] text-slate-500 underline">인원 새로고침</button>
              {targets.length > 0 && (
                <button type="button" onClick={() => setExcludes(excludes.length ? [] : targets.map((c) => c.full))}
                  className="text-[11px] text-slate-500 underline">
                  {excludes.length ? "제외 해제" : "전체 제외"}
                </button>
              )}
            </div>
          </div>
          {statsMsg && (
            <p className={`text-[11px] mb-1.5 rounded p-2 ${statsMsg.startsWith("인원 불러오는") ? "text-slate-400" : "bg-amber-50 text-amber-700 border border-amber-200"}`}>
              {statsMsg}
            </p>
          )}
          {targets.length ? (
            <div className="border border-slate-200 rounded-lg divide-y divide-slate-100 max-h-64 overflow-y-auto">
              {targets.map((c) => {
                const off = excludes.includes(c.full);
                return (
                  <button key={c.full} type="button" onClick={() => toggle(c.full)}
                    className={`flex items-center gap-2 w-full text-left px-3 py-2 text-sm ${off ? "bg-slate-50 text-slate-400 line-through" : "hover:bg-emerald-50"}`}>
                    <span className={`w-4 h-4 rounded border flex items-center justify-center shrink-0 ${off ? "border-slate-300" : "bg-emerald-600 border-emerald-600"}`}>
                      {!off && <CheckCircle2 size={12} className="text-white" />}
                    </span>
                    <span className="flex-1">{c.full}</span>
                    {countOf(c.full) !== undefined && (
                      <span className={`shrink-0 text-xs font-semibold px-2 py-0.5 rounded-full ${off ? "bg-slate-100 text-slate-400" : countOf(c.full) === 0 ? "bg-slate-100 text-slate-500" : "bg-emerald-50 text-emerald-700"}`}>
                        {countOf(c.full)}/{capacityOf(c.full)}
                      </span>
                    )}
                  </button>
                );
              })}
            </div>
          ) : <p className="text-xs text-slate-400">해당하는 강좌가 없습니다.</p>}
          <p className="text-[11px] text-slate-400 mt-1.5">수강생이 0명인 강좌는 자동으로 건너뜁니다. 반변경자는 기존 반에서 <b>제외만</b> 되며, 새 반 등록은 직접 해주세요.</p>

        </div>
      )}

      <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer select-none">
        <input type="checkbox" checked={safe} onChange={(e) => setSafe(e.target.checked)}
          className="w-4 h-4 rounded border-slate-300 accent-emerald-600" />
        안전모드 (강좌마다 [예약 생성] 은 PC에서 직접 클릭)
      </label>

      {confirming ? (
        <div className="bg-emerald-50 border border-emerald-300 rounded-xl p-3 space-y-2">
          <p className="text-sm text-emerald-900">
            <b>{willRun.length}개</b> 강좌를 <b>{fromYm} → {toYm}</b> 로 이월합니다. 진행할까요?
          </p>
          <div className="flex gap-2">
            <button onClick={run} className="flex-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2">
              네, 진행합니다
            </button>
            <button onClick={() => setConfirming(false)} className="flex-1 rounded-lg border border-slate-300 text-slate-600 text-sm py-2">
              취소
            </button>
          </div>
        </div>
      ) : (
        <button onClick={ask} disabled={saving || running || !willRun.length}
          className="w-full rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-sm font-medium py-2.5 flex items-center justify-center gap-1.5">
          {saving || running ? <><Loader2 size={15} className="animate-spin" /> 이월 진행 중…</> : <><PlayCircle size={15} /> 이월 실행 ({willRun.length}개)</>}
        </button>
      )}

      {err && (
        <p className={`text-xs rounded-lg p-2.5 border ${err.startsWith("참고") ? "bg-amber-50 text-amber-800 border-amber-200" : "bg-rose-50 text-rose-700 border-rose-200"}`}>
          {err}
        </p>
      )}

      {job && (
        <div className={`rounded-lg p-3 text-sm ${job.status === "error" ? "bg-rose-50 border border-rose-200 text-rose-700" : "bg-slate-50 border border-slate-200"}`}>
          {job.progress && <p className="font-medium">{job.progress}</p>}
          {job.result && <p className="text-xs mt-1 whitespace-pre-wrap">{job.result}</p>}
          {Array.isArray(job.detail) && job.detail.length > 0 && (
            <details className="mt-2">
              <summary className="text-xs cursor-pointer text-slate-500">강좌별 결과 보기</summary>
              <div className="mt-1 space-y-0.5">
                {job.detail.map((d, i) => (
                  <div key={i}>
                    <p className="text-[11px] text-slate-600">
                      {d.ok ? "✔" : "✖"} {d.course} — {d.msg}
                    </p>
                    {d.manual && d.manual.length > 0 && (
                      <p className="text-[11px] text-amber-700 pl-3">↳ 수동 등록 필요: {d.manual.join(", ")}</p>
                    )}
                  </div>
                ))}
              </div>
            </details>
          )}
        </div>
      )}
    </div>
  );
}

/* 강습 — 기존 자동화(lesson_jobs) 실행 + work_logs 기록 */
function LessonForm({ me }) {
  const isAI = me.id === AI_ADMIN;
  const [startYm, setStartYm] = useState(ymNow());
  const [q, setQ] = useState("");
  const [picked, setPicked] = useState(null);

  /* 강좌명 검색 — 띄어쓰기 무시, 여러 단어 모두 포함하는 강좌를 찾는다 */
  const flatten = (t) => String(t).replace(/[\s\-()/·.,]/g, "");
  const words = q.trim().split(/\s+/).map(flatten).filter(Boolean);
  const hits = words.length
    ? ALL_COURSES.filter((c) => {
        const flat = flatten(c.full);
        return words.every((w) => flat.includes(w));
      }).slice(0, 30)
    : [];
  const course = picked ? picked.full : "";
  const cat = picked ? picked.cat : "";

  const [action, setAction] = useState("등록");
  const [dong, setDong] = useState(""); const [ho, setHo] = useState(""); const [name, setName] = useState("");
  const [amount, setAmount] = useState(""); const [closed, setClosed] = useState(""); const [safe, setSafe] = useState(true);
  const [running, setRunning] = useState(false); const [done, setDone] = useState(null);

  const run = () => {
    if (!course) { alert("강좌를 먼저 선택하세요."); return; }
    if (!name.trim()) { alert("회원 이름을 입력하세요."); return; }
    if (action === "등록" && (!dong.trim() || !ho.trim())) { alert("등록 시 동·호수는 필수입니다."); return; }
    setRunning(true);
    const at = nowLocal().replace("T", " ");
    const job = { action, cat, course, member: name, dong, ho, amount, closed, safe, start_ym: startYm, status: "pending", by: me.id, at };
    supabase.from("lesson_jobs").insert(job).then(({ error }) => {
      setRunning(false);
      if (error) { alert("요청 전송 실패: " + error.message); return; }
      logWork({ kind: "강습", action, target: course, seat: cat, member: name, dong, ho, amount, timing: closed ? `휴강 ${closed}` : "", by: me.id, at });
      setDone(`${action} 요청 전송 완료 — PC 파이썬이 처리합니다.`);
      setDong(""); setHo(""); setName(""); setAmount(""); setClosed("");
      setTimeout(() => setDone(null), 4000);
    });
  };

  return (
    <div>
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
        <ActionToggle action={action} setAction={setAction} />
        {isAI && (
          <div className="bg-amber-50 border border-amber-200 rounded-lg p-3">
            <PickRow label="수강 시작 연월 (어느 달 창구에 넣을지)" options={ymList()} value={startYm} onPick={setStartYm} render={ymLabel} />
            {startYm !== ymNow() && <p className="text-[11px] text-amber-700 mt-2">선택한 달 기준으로 차감·환불액이 계산됩니다.</p>}
          </div>
        )}

        <div>
          <span className="text-xs font-medium text-slate-500 block mb-1.5">강좌 검색</span>
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
            <input value={q} onChange={(e) => { setQ(e.target.value); setPicked(null); }}
              placeholder="예: 2단지 요가 / 수영 화목 07 / 요가 박상희"
              className="w-full rounded-lg border border-slate-200 pl-9 pr-8 py-2.5 text-sm outline-none focus:border-emerald-500" />
            {q && <button onClick={() => { setQ(""); setPicked(null); }}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-slate-400"><X size={15} /></button>}
          </div>

          {picked ? (
            <div className="mt-2 bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2.5 flex items-start justify-between gap-2">
              <span className="text-sm text-emerald-800">✔ <b>{course}</b></span>
              <button onClick={() => setPicked(null)} className="text-xs text-emerald-700 shrink-0 underline">변경</button>
            </div>
          ) : q.trim() ? (
            hits.length ? (
              <div className="mt-2 border border-slate-200 rounded-lg divide-y divide-slate-100 max-h-72 overflow-y-auto">
                <p className="text-[11px] text-slate-400 px-3 py-1.5 bg-slate-50">{hits.length}개 검색됨 — 강좌를 누르세요</p>
                {hits.map((c) => (
                  <button key={c.full} onClick={() => { setPicked(c); setQ(c.full); }}
                    className="block w-full text-left px-3 py-2 text-sm hover:bg-emerald-50">
                    {c.full}
                    <span className="block text-[10px] text-slate-400">{c.floor} · {c.cat}</span>
                  </button>
                ))}
              </div>
            ) : <p className="text-xs text-slate-400 mt-2">검색 결과가 없습니다. 단어를 줄여보세요 (예: "요가").</p>
          ) : <p className="text-[11px] text-slate-400 mt-1.5">강좌명 일부만 입력해도 됩니다. 여러 단어를 띄어 쓰면 모두 포함된 강좌를 찾습니다.</p>}
        </div>

        <DongHoName {...{ action, dong, setDong, ho, setHo, name, setName }} />
        <div className="grid grid-cols-2 gap-2">
          <Fld label="수동 차감/환불액 (선택)"><input value={amount} onChange={(e) => setAmount(e.target.value)} placeholder="비워두면 자동 계산 (권장)" className="rin" /></Fld>
          <Fld label="휴강일 (선택)"><input value={closed} onChange={(e) => setClosed(e.target.value)} placeholder="예: 8/12, 8/19" className="rin" /></Fld>
        </div>
        <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer select-none">
          <input type="checkbox" checked={safe} onChange={(e) => setSafe(e.target.checked)} className="w-4 h-4 rounded border-slate-300 accent-emerald-600" />
          안전모드 (최종 버튼은 PC에서 직접 클릭)
        </label>
        <RunButton action={action} running={running} onClick={run} label="실행" />
        {done && <p className="text-sm text-emerald-700 text-center">✔ {done}</p>}
      </div>
      <style>{`.rin{width:100%;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:13px;outline:none}.rin:focus{border-color:#10b981;box-shadow:0 0 0 3px #d1fae5}`}</style>
    </div>
  );
}

/* 독서실·골프백 — 금액 자동계산 + 기록 (실제 BYB 처리는 직원이 계산금액으로 수행 / 자동화는 추후) */
function ReservationForm({ me, kind, options, seatLabel, seatPh }) {
  const [opt, setOpt] = useState(options[0][1]);
  const [action, setAction] = useState("등록");
  const [dateStr, setDateStr] = useState(today());
  const [dong, setDong] = useState(""); const [ho, setHo] = useState(""); const [name, setName] = useState(""); const [seat, setSeat] = useState("");
  const [amount, setAmount] = useState(""); const [done, setDone] = useState(null); const [saving, setSaving] = useState(false);
  const [timing, setTiming] = useState("endOfToday"); const [reason, setReason] = useState("");
  const [safe, setSafe] = useState(true);
  const [reception, setReception] = useState(""); const [editRec, setEditRec] = useState(false); const [recDraft, setRecDraft] = useState("");
  const isAI = me.id === AI_ADMIN;
  const recKey = `reception_${kind}`;
  // 독서실·골프백 모두 PC 파이썬(v3) 자동 실행 지원
  const canAutomate = kind === "독서실" || kind === "골프백";

  React.useEffect(() => {
    supabase.from("settings").select("value").eq("key", recKey).maybeSingle()
      .then(({ data }) => { if (data) setReception(data.value); });
  }, [recKey]);

  const day = dateStr ? new Date(dateStr).getDate() : null;
  const auto = calcAmount(opt, action, day);
  const optLabel = options.find((o) => o[1] === opt)[0];
  const targetName = kind === "독서실" ? `${optLabel}독서실` : `${optLabel}`;

  const saveReception = () => {
    supabase.from("settings").upsert({ key: recKey, value: recDraft }).then(({ error }) => {
      if (!error) { setReception(recDraft); setEditRec(false); }
      else alert("저장 실패: " + error.message);
    });
  };

  const run = async () => {
    if (!name.trim() || !dong.trim() || !ho.trim()) { alert("동·호수·이름을 입력하세요."); return; }
    if (action === "등록" && !seat.trim()) { alert(`등록 시 ${seatLabel}는 필수입니다.`); return; }
    const amt = amount.trim() || (auto != null ? String(auto) : "");
    const productKey = kind === "독서실" ? `${optLabel}독서실` : `골프백(${optLabel})`;
    setSaving(true);
    const at = nowLocal().replace("T", " ");
    // 기록 저장
    await logWork({ kind, action, target: productKey, seat, member: name, dong, ho, amount: amt, timing: `${action === "등록" ? "등록일" : "해지일"} ${dateStr}${action === "취소" ? " · " + (timing === "endOfToday" ? "오늘24시" : "즉시") : ""}`, by: me.id, at });
    if (canAutomate) {
      const { error } = await supabase.from("lesson_jobs").insert({
        kind, action, course: productKey, member: name, dong, ho, seat, amount: amt,
        timing: action === "취소" ? timing : "", reason: reason || "관리사무소 요청",
        safe, status: "pending", by: me.id, at,
      });
      setSaving(false);
      if (error) { alert("실행 요청 실패: " + error.message); return; }
      setDone(`${productKey} ${action} 요청 전송 완료 (금액 ${amt ? Number(amt).toLocaleString() + "원" : "-"}${safe ? ", 안전모드" : ""}).`);
    } else {
      setSaving(false);
      setDone(`${productKey} ${action} 기록 저장 완료 (금액 ${amt ? Number(amt).toLocaleString() + "원" : "-"})`);
    }
    setDong(""); setHo(""); setName(""); setSeat(""); setAmount(""); setReason("");
    setTimeout(() => setDone(null), 6000);
  };

  return (
    <div>
      <div className="bg-white rounded-xl border border-slate-200 p-3 mb-4 flex items-center gap-2 flex-wrap">
        <span className="text-xs font-semibold text-slate-500">접수기간</span>
        {!editRec ? (
          <>
            <span className="text-sm font-medium">{reception || "미설정"}</span>
            {isAI && <button onClick={() => { setRecDraft(reception); setEditRec(true); }} className="ml-auto text-xs text-emerald-600 flex items-center gap-1"><Pencil size={12} /> 수정</button>}
          </>
        ) : (
          <>
            <input value={recDraft} onChange={(e) => setRecDraft(e.target.value)} className="flex-1 min-w-[180px] rounded-lg border border-slate-200 px-2 py-1 text-sm outline-none focus:border-emerald-500" placeholder="2026-06-15 ~ 2026-06-30" />
            <button onClick={saveReception} className="rounded-lg bg-emerald-600 text-white text-xs px-3 py-1.5">저장</button>
            <button onClick={() => setEditRec(false)} className="rounded-lg border border-slate-200 text-slate-500 text-xs px-3 py-1.5">취소</button>
          </>
        )}
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
        <div>
          <span className="text-xs font-medium text-slate-500 block mb-1.5">{kind === "독서실" ? "독서실 종류" : "락커 위치"}</span>
          <div className="inline-flex p-1 bg-slate-100 rounded-lg">
            {options.map(([label, key]) => (
              <button key={key} onClick={() => setOpt(key)}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition ${opt === key ? "bg-slate-800 text-white" : "text-slate-500"}`}>{label}</button>
            ))}
          </div>
        </div>

        <ActionToggle action={action} setAction={setAction} />

        <div className="grid grid-cols-2 gap-2">
          <Fld label={action === "등록" ? "등록일" : "해지일"}>
            <input type="date" value={dateStr} onChange={(e) => setDateStr(e.target.value)} className="rin" />
          </Fld>
          <Fld label={seatLabel + " *"}>
            <input value={seat} onChange={(e) => setSeat(e.target.value)} placeholder={seatPh} className="rin" />
          </Fld>
        </div>

        <DongHoName action="등록" dong={dong} setDong={setDong} ho={ho} setHo={setHo} name={name} setName={setName} />

        {action === "취소" && (
          <div className="grid grid-cols-2 gap-2">
            <Fld label="해지 시점">
              <select value={timing} onChange={(e) => setTiming(e.target.value)} className="rin">
                <option value="endOfToday">오늘 24시까지 사용</option>
                <option value="now">지금부터 사용 중지</option>
              </select>
            </Fld>
            <Fld label="사유 (선택)">
              <input value={reason} onChange={(e) => setReason(e.target.value)} placeholder="관리사무소 요청" className="rin" />
            </Fld>
          </div>
        )}

        <div className="rounded-xl bg-slate-50 border border-slate-200 p-3">
          <div className="flex items-center justify-between">
            <span className="text-xs text-slate-500">{action === "등록" ? "차감" : "환불"} 금액 (자동)</span>
            <span className="text-lg font-bold text-emerald-700">{auto != null ? auto.toLocaleString() + "원" : "-"}</span>
          </div>
          <p className="text-[11px] text-slate-400 mt-1">{targetName} · {day}일 기준 (한 달 30일, 31일은 30일로 계산)</p>
          <label className="flex items-center gap-2 mt-2 text-xs text-slate-500">
            직접 수정:
            <input value={amount} onChange={(e) => setAmount(e.target.value)} placeholder={auto != null ? auto.toString() : "금액"} className="flex-1 rounded-lg border border-slate-200 px-2 py-1 text-sm outline-none focus:border-emerald-500" />
          </label>
        </div>

        {canAutomate && (
          <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer select-none">
            <input type="checkbox" checked={safe} onChange={(e) => setSafe(e.target.checked)} className="w-4 h-4 rounded border-slate-300 accent-emerald-600" />
            안전모드 (최종 버튼은 PC에서 직접 클릭)
          </label>
        )}
        <RunButton action={action} running={saving} onClick={run} label={canAutomate ? "실행" : "기록 저장"} />
        {done && <p className="text-sm text-emerald-700 text-center">✔ {done}</p>}
      </div>
      <style>{`.rin{width:100%;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:13px;outline:none}.rin:focus{border-color:#10b981;box-shadow:0 0 0 3px #d1fae5}`}</style>
    </div>
  );
}

/* 공통 소품 */
function ActionToggle({ action, setAction }) {
  return (
    <div>
      <span className="text-xs font-medium text-slate-500 block mb-1.5">처리 구분</span>
      <div className="inline-flex p-1 bg-slate-100 rounded-lg">
        {["등록", "취소"].map((a) => (
          <button key={a} onClick={() => setAction(a)}
            className={`px-4 py-1.5 rounded-md text-sm font-medium transition ${action === a ? (a === "등록" ? "bg-emerald-600 text-white" : "bg-rose-500 text-white") : "text-slate-500"}`}>
            {a === "등록" ? "등록(차감)" : "취소(환불)"}
          </button>
        ))}
      </div>
    </div>
  );
}
function DongHoName({ action, dong, setDong, ho, setHo, name, setName }) {
  return (
    <div className="grid grid-cols-3 gap-2">
      <Fld label={`동 ${action === "등록" ? "*" : ""}`}><input value={dong} onChange={(e) => setDong(e.target.value)} placeholder="208" className="rin" /></Fld>
      <Fld label={`호수 ${action === "등록" ? "*" : ""}`}><input value={ho} onChange={(e) => setHo(e.target.value)} placeholder="1504" className="rin" /></Fld>
      <Fld label="이름 *"><input value={name} onChange={(e) => setName(e.target.value)} placeholder="회원 성함" className="rin" /></Fld>
    </div>
  );
}
function Fld({ label, children }) {
  return <label className="block"><span className="text-xs font-medium text-slate-500 block mb-1.5">{label}</span>{children}</label>;
}
function RunButton({ action, running, onClick, label }) {
  return (
    <button onClick={onClick} disabled={running}
      className={`w-full flex items-center justify-center gap-2 rounded-lg text-white text-sm font-medium py-2.5 transition disabled:opacity-50
        ${action === "등록" ? "bg-emerald-600 hover:bg-emerald-700" : "bg-rose-600 hover:bg-rose-700"}`}>
      {running ? <><Loader2 size={16} className="animate-spin" /> 처리 중…</> : <><PlayCircle size={16} /> {label} ({action})</>}
    </button>
  );
}

/* ─────────────────────── 작업 기록 + 엑셀 다운로드 */
function WorkLogs({ me }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [filter, setFilter] = useState("전체");

  const load = () => {
    setLoading(true);
    supabase.from("work_logs").select("*").order("id", { ascending: false }).limit(500)
      .then(({ data }) => { setRows(data || []); setLoading(false); });
  };
  React.useEffect(load, []);

  const list = rows.filter((r) => filter === "전체" || r.kind === filter);

  const exportExcel = () => {
    const data = list.map((r) => ({
      "종류": r.kind, "구분": r.action, "대상": r.target, "좌석/락커/분류": r.seat,
      "동": r.dong, "호": r.ho, "이름": r.member, "금액": r.amount, "비고": r.timing,
      "처리자": r.by, "시각": r.at,
    }));
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "작업기록");
    XLSX.writeFile(wb, `작업기록_${today()}.xlsx`);
  };

  const del = async (id) => {
    if (!askYesNo("이 작업 기록을 삭제할까요?")) return;
    const { error } = await supabase.from("work_logs").delete().eq("id", id);
    if (error) { alert("삭제 실패: " + error.message); return; }
    setRows(rows.filter((r) => r.id !== id));
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-base font-bold">작업 기록 <span className="text-slate-400 font-normal text-sm">({list.length})</span></h2>
        <div className="flex gap-2">
          <button onClick={load} className="rounded-lg border border-slate-200 text-slate-500 text-xs px-3 py-2">새로고침</button>
          <button onClick={exportExcel} disabled={!list.length} className="flex items-center gap-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-xs font-medium px-3 py-2"><Upload size={13} /> 엑셀 다운로드</button>
        </div>
      </div>
      <div className="flex gap-1.5 mb-3">
        {["전체", "강습", "독서실", "골프백"].map((k) => (
          <button key={k} onClick={() => setFilter(k)} className={`text-xs font-medium px-3 py-1.5 rounded-full ${filter === k ? "bg-emerald-600 text-white" : "bg-white border border-slate-200 text-slate-500"}`}>{k}</button>
        ))}
      </div>
      {loading ? <p className="text-center text-slate-400 text-sm py-10"><Loader2 size={16} className="animate-spin inline" /> 불러오는 중…</p>
        : !list.length ? <p className="text-center text-slate-400 text-sm py-10">기록이 없습니다.</p>
        : <div className="space-y-2">
            {list.map((r) => (
              <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm">
                <div className="flex items-start justify-between gap-2">
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className="text-xs px-2 py-0.5 rounded bg-slate-100 text-slate-600">{r.kind}</span>
                    <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${r.action === "등록" ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700"}`}>{r.action}</span>
                    {r.amount && <span className="text-xs text-slate-500">{Number(r.amount).toLocaleString()}원</span>}
                  </div>
                  <button onClick={() => del(r.id)} className="text-slate-300 hover:text-rose-500 p-1 shrink-0" title="삭제"><Trash2 size={13} /></button>
                </div>
                <p className="text-sm font-medium mt-1">{r.target}{r.seat ? ` · ${r.seat}` : ""}</p>
                <p className="text-xs text-slate-500 mt-0.5">{r.dong}동 {r.ho}호 · {r.member}{r.timing ? ` · ${r.timing}` : ""}</p>
                <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.at} · {r.by}</p>
              </div>
            ))}
          </div>}
    </div>
  );
}


/* ─────────────────────── 외부 연동 안내 */
function LinkInfo() {
  return (
    <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5">
      <div className="flex items-center gap-2 mb-2"><ArrowLeftRight size={18} className="text-emerald-600" /><h2 className="font-bold">외부 홈페이지 연동 (강습 취소 등록)</h2></div>
      <p className="text-sm text-slate-600 leading-relaxed">
        결론부터: <b className="text-emerald-700">조건부로 가능합니다.</b> 다만 이 웹페이지가 직접 다른 사이트(BYB 등)를 클릭할 수는 없고, 방식이 세 가지로 나뉩니다.
      </p>
      <div className="mt-3 space-y-2 text-sm text-slate-600">
        <div className="rounded-lg bg-emerald-50 border border-emerald-100 p-3"><b>① 공식 API가 있는 경우 (최선)</b><br />BYB가 예약/취소 API를 제공하면 버튼 한 번으로 깔끔하게 연동됩니다.</div>
        <div className="rounded-lg bg-slate-50 border border-slate-200 p-3"><b>② 자동화 스크립트 연결 (현실적)</b><br />이미 쓰시는 Selenium/Playwright 스크립트를 서버에서 돌려 취소를 대신 처리. 가능하지만 사이트가 바뀌면 손봐야 합니다.</div>
        <div className="rounded-lg bg-slate-50 border border-slate-200 p-3"><b>③ 브라우저에서 바로 조작</b><br />보안(CORS) 때문에 <b>불가능</b>합니다.</div>
      </div>
    </div>
  );
}

/* ─────────────────────── 공통 */
function L({ label, children }) { return <label className="block"><span className="text-[11px] font-medium text-slate-500 mb-1 block">{label}</span>{children}</label>; }
function Panel({ title, children }) { return <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 mb-4"><p className="text-sm font-semibold text-slate-700 mb-3">{title}</p>{children}</div>; }
function Stat({ label, value, accent }) {
  return <div className={`rounded-xl border p-3 ${accent ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-200"}`}>
    <p className={`text-xs ${accent ? "text-emerald-100" : "text-slate-500"}`}>{label}</p><p className="text-base font-bold mt-0.5">{value}</p></div>;
}
