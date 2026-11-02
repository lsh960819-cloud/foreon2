import React, { useState, useMemo } from "react";
import {
  LogIn, LogOut, Search, PackageSearch, ClipboardList, CalendarDays, Home as HomeIcon,
  Plus, X, User, Loader2, MessageSquareWarning, ArrowLeftRight, Inbox, Send, Upload,
  CheckCircle2, Clock, FileText, GraduationCap, PlayCircle, Pencil, Trash2, Copy,
  FolderOpen, Download, RefreshCw, CalendarCheck, Menu, KeyRound, Monitor, AlertTriangle, Sparkles
} from "lucide-react";
import * as XLSX from "xlsx";
import { ResponsiveContainer, LineChart, Line, BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip } from "recharts";
import { supabase } from "./supabaseClient.js";

/* ═══════════════════════════════════════════════════════════
   계정: foreon1 ~ foreon9 (골프장·피트니스 계정은 폐지)
   부서 배정 (필요하면 여기 숫자만 바꾸면 됩니다)
   ═══════════════════════════════════════════════════════════ */
const DEPT_OF = {
  foreon1: "사무실", foreon2: "사무실", foreon3: "사무실", foreon4: "사무실", foreon5: "사무실",
  foreon6: "데스크", foreon8: "데스크", foreon9: "데스크",
};
/* 개인별 비밀번호 — 사무실: Office{번호}!26 (패턴형) / 데스크: 무작위(직원끼리 유추 방지) */
const PW_OF = {
  foreon1: "Office1!26", foreon2: "Office2!26", foreon3: "Office3!26", foreon4: "Office4!26", foreon5: "Office5!26",
  foreon6: "xQ4#nZt8Bw", foreon8: "Gk6&wTf1Qz", foreon9: "Rn2@Xy7Lqe",
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
  return <p className="text-[11px] text-amber-600 mt-1">✎ {last.at} {last.by} 수정{r.hist.length > 1 ? ` (총 ${r.hist.length}회)` : ""}</p>;
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
    if (!PW_OF[s.id] || Date.now() - s.at > SESSION_HOURS * 3600 * 1000) {   // 삭제된 계정·만료 세션은 로그아웃
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
  const [nav, setNav] = useState(false);

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

  const tabs = [
    { id: "home", label: "홈", icon: HomeIcon },
    { id: "attendance", label: "출석부 생성", icon: CalendarCheck },
    { id: "apply", label: "신청서 생성", icon: FileText },
    { id: "reading", label: "독서실 신청서·대기자", icon: FileText },
    { id: "lessonbyb", label: "강습 등록·취소", icon: GraduationCap },
    { id: "stats", label: "월간 이용 통계", icon: Search },
    { id: "saunakey", label: "사우나 핸드키", icon: KeyRound },
    { id: "screen", label: "스크린천 주문", icon: Monitor },
    { id: "cleaning", label: "시설 청소 관리", icon: Sparkles },
    { id: "link", label: "외부 연동", icon: ArrowLeftRight },
  ];

  const cur = tabs.find((t) => t.id === tab) || tabs[0];
  const go = (id) => { setTab(id); setNav(false); };

  return (
    <div className="min-h-screen bg-byb-bg text-byb-text">
      {/* 상단 바 */}
      <header className="bg-white border-b border-byb-line sticky top-0 z-30">
        <div className="h-[50px] px-3 flex items-center justify-between">
          <div className="flex items-center gap-3">
            <button onClick={() => setNav((v) => !v)} className="md:hidden w-8 h-8 flex items-center justify-center border border-byb-line text-slate-500">
              <Menu size={16} />
            </button>
            <span className="text-[22px] font-bold tracking-tight text-emerald-600 leading-none">클럽포레온</span>
            <span className="hidden sm:block w-px h-4 bg-slate-300" />
            <span className="hidden sm:block text-[13px] font-medium text-slate-600">커뮤니티 운영관리시스템</span>
          </div>
          <div className="flex items-center gap-2">
            <span className="flex items-center gap-1.5 h-8 px-3 border border-byb-line text-[13px]">
              <User size={13} className="text-slate-500" /><b className="font-medium">{me.id}</b>
              <span className="text-xs text-byb-sub">{me.dept}</span>
            </span>
            <button onClick={logout} title="로그아웃" className="w-8 h-8 flex items-center justify-center border border-byb-line text-slate-500 hover:text-slate-700"><LogOut size={14} /></button>
          </div>
        </div>
      </header>

      <div className="flex">
        {/* 왼쪽 메뉴 */}
        <aside className={`${nav ? "block" : "hidden"} md:block fixed md:sticky top-[50px] left-0 z-20 w-[200px] shrink-0 h-[calc(100vh-50px)] overflow-y-auto bg-white border-r border-byb-line`}>
          <nav className="py-3">
            {tabs.map((t) => {
              const on = tab === t.id; const Icon = t.icon;
              return (
                <button key={t.id} onClick={() => go(t.id)}
                  className={`w-full flex items-center gap-2.5 pl-5 pr-3 h-11 text-[13px] text-left transition border-r-2
                    ${on ? "bg-emerald-50 text-emerald-600 font-medium border-emerald-600" : "text-slate-700 border-transparent hover:bg-slate-50"}`}>
                  <Icon size={15} className={on ? "text-emerald-600" : "text-slate-500"} /> {t.label}
                </button>
              );
            })}
          </nav>
        </aside>
        {nav && <div onClick={() => setNav(false)} className="md:hidden fixed inset-0 top-[50px] z-10 bg-black/20" />}

        {/* 본문 */}
        <main className="flex-1 min-w-0 px-4 md:px-6 pt-6 pb-20">
          <div className="max-w-5xl">
            <h1 className="text-[22px] font-bold text-emerald-600 mb-5">{cur.label}</h1>
            {tab === "home" && <Home me={me} events={events} />}
            {tab === "attendance" && <Attendance key="att" me={me} />}
            {tab === "apply" && <><Attendance key="apply" mode="apply" me={me} /><SwimClassChange me={me} /></>}
            {tab === "reading" && <ReadingRoom />}
            {tab === "stats" && <Attendance key="stats" mode="stats" me={me} />}
            {tab === "lessonbyb" && <Attendance key="lesson" mode="lesson" me={me} />}
            {tab === "link" && <LinkInfo />}
            {tab === "saunakey" && <SaunaKeys />}
            {tab === "screen" && <ScreenCloth />}
            {tab === "cleaning" && <Cleaning me={me} />}
          </div>
        </main>
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
    <div className="relative min-h-screen overflow-hidden bg-slate-800 flex items-center justify-center px-4">
      {/* 배경 영상: 올림픽파크포레온 소개 영상 앞 15초를 0.5배속(30초)으로, 소리 없이 반복 */}
      <video className="absolute inset-0 w-full h-full object-cover" src="/login-bg.mp4" poster="/login-bg.jpg"
        autoPlay muted loop playsInline preload="auto" />
      <div className="absolute inset-0 bg-black/45" />
      <div className="relative w-full max-w-sm">
        <div className="text-center mb-6 text-white">
          <div className="inline-flex items-center gap-2 text-xs tracking-widest text-white/80 mb-2"><span className="w-1.5 h-1.5 rounded-full bg-emerald-400" /> FOREON 커뮤니티</div>
          <h1 className="text-2xl font-bold drop-shadow">커뮤니티 운영</h1>
          <p className="text-sm text-white/80 mt-1">아이디로 로그인하세요</p>
        </div>
        <div className="bg-white/95 border border-byb-line p-6 space-y-4">
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
          <p className="text-[11px] text-slate-400 mt-2">{r.time} · {r.owner} · {r.by}</p>
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
        <span className="text-xs text-slate-400">{r.dept}</span>
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
      <p className="text-[11px] text-slate-400 mt-1.5">{r.at} · {r.by}</p>
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
      <p className="text-[11px] text-slate-400 mt-1.5">{r.date} · {r.by} ({r.dept})</p>
      <AuditLine r={r} />
      {r.replies.map((rp, i) => (
        <div key={i} className="mt-2 ml-3 pl-3 border-l-2 border-emerald-200">
          <p className="text-sm text-emerald-800">↳ {rp.text}</p>
          <p className="text-[11px] text-slate-400">{rp.date} · {rp.by} (사무실)</p>
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
                  <span className="text-xs text-slate-400 ml-auto">{r.date}</span>
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
                <p className="text-[11px] text-slate-400 mt-2">등록 · {r.by}</p>
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

  const [hist, setHist] = useState([]);        // 수강인원 일별 스냅샷 (추이 그래프용)
  const [histErr, setHistErr] = useState("");

  const load = React.useCallback(async () => {
    const { data } = await supabase.from("course_stats").select("*").order("course");
    setStats(data || []);
    setLoading(false);
    // 오늘 날짜 스냅샷 저장(하루 1줄, 덮어쓰기) → 추이 불러오기
    const real = (data || []).filter((s) => s.synced_at);
    if (real.length) {
      const snap = Object.fromEntries(real.map((s) => [s.course, s.enrolled || 0]));
      await supabase.from("course_stats_history")
        .upsert({ snap_date: kstDay(), data: snap, at: stamp() }, { onConflict: "snap_date" });
    }
    const from = kstDay(-400);
    const h = await supabase.from("course_stats_history").select("snap_date,data")
      .gte("snap_date", from).order("snap_date");
    if (h.error) setHistErr(h.error.message);
    else { setHist(h.data || []); setHistErr(""); }
  }, []);
  React.useEffect(() => { load(); }, [load]);

  /* 수강인원 업데이트 — 바이비 이번 달(KST) 창구별 확정 수강생 수를 직접 세어 course_stats 를 통째로 갱신
     (예전 전산PC 워커 방식은 PC 프로그램이 꺼져 있으면 무한 대기 → 홈페이지 API 로 직접 처리) */
  const [bybLogin, setBybLogin] = useState(null);   // null | {step:"login"|"otp", user, pw, code, pending}
  const [updErr, setUpdErr] = useState("");
  const ymKst = kstDay().slice(0, 7);
  const curY = +ymKst.slice(0, 4), curM = +ymKst.slice(5, 7);

  const runUpdate = async () => {
    if (!BYB_AUTH) { setBybLogin({ step: "login", user: (() => { try { return localStorage.getItem("byb_user") || ""; } catch { return ""; } })(), pw: "", code: "" }); return; }
    setUpdErr(""); setJob({ status: "working", label: "강좌 목록 불러오는 중", done: 0, total: 0 });
    let cur = BYB_AUTH;
    const call = async (body) => {
      const j = await attCall({ ...body, auth: cur });
      if (j.auth) { cur = j.auth; BYB_AUTH = j.auth; }
      return j;
    };
    try {
      const { windows = [] } = await call({ action: "windows", year: curY, month: curM });
      const ws = windows.filter((w) => !String(w.name).startsWith("플그"));
      if (!ws.length) throw new Error("바이비에서 2단지 강좌 창구를 찾지 못했습니다.");
      const parts = [];
      for (let i = 0; i < ws.length; i += 15) parts.push(ws.slice(i, i + 15).map((w) => ({ id: w.id, name: w.name })));
      const counts = {};
      let done = 0, idx = 0;
      setJob({ status: "working", label: `${curM}월 창구 인원 세는 중`, done: 0, total: ws.length });
      const worker = async () => {
        while (idx < parts.length) {
          const p = parts[idx++];
          const j = await call({ action: "stats_count", year: curY, month: curM, windows: p });
          Object.assign(counts, j.counts || {});
          done += p.length;
          setJob({ status: "working", label: `${curM}월 창구 인원 세는 중`, done, total: ws.length });
        }
      };
      await Promise.all([worker(), worker(), worker()]);

      // 같은 이름 창구는 합산. 0명 창구는 앱 강좌 목록(COURSES)에 있는 것만 남김 (지난 특강·옛 창구 제외)
      const ck = (t) => String(t).replace(/[^0-9A-Za-z가-힣]/g, "");
      const known = new Set(ALL_COURSES.map((c) => ck(c.full)));
      const at = stamp(), byName = new Map();
      for (const w of ws) {
        const n = counts[w.id] || 0;
        if (!n && !known.has(ck(w.label || w.name))) continue;
        const name = w.label || w.name;
        byName.set(name, (byName.get(name) || 0) + n);
      }
      const rowsNew = [...byName.entries()].map(([course, enrolled]) => {
        const p = parseCourse(course);
        return { course, enrolled, capacity: capacityOf(course), cat: CAT_LABEL[p.cat] || p.cat,
          floor: /수영|아쿠아/.test(course) ? "B2" : "B1", synced_at: at };
      });
      const up = await supabase.from("course_stats").upsert(rowsNew, { onConflict: "course" });
      if (up.error) throw new Error("저장 실패: " + up.error.message);
      // 이번 달 바이비에 없는 예전 강좌: 특이사항이 있으면 인원만 비우고, 없으면 삭제
      const stale = stats.filter((s) => !byName.has(s.course));
      const keep = stale.filter((s) => (s.note || "").trim()).map((s) => s.course);
      const drop = stale.filter((s) => !(s.note || "").trim()).map((s) => s.course);
      if (keep.length) await supabase.from("course_stats").update({ enrolled: 0, synced_at: null }).in("course", keep);
      if (drop.length) await supabase.from("course_stats").delete().in("course", drop);
      setJob({ status: "done" });
      await load();
    } catch (e) {
      setJob(null);
      setUpdErr(e.message);
      if (/만료|로그인 정보|401|토큰/.test(e.message)) BYB_AUTH = null;
    }
  };

  const bybStep = async () => {
    const b = bybLogin; setUpdErr("");
    try {
      if (b.step === "login") {
        const j = await attCall({ action: "login", user: b.user, password: b.pw });
        try { localStorage.setItem("byb_user", b.user); } catch {}
        if (j.auth) { BYB_AUTH = j.auth; setBybLogin(null); runUpdate(); }
        else setBybLogin({ ...b, pw: "", step: "otp", pending: j.pending });
      } else {
        const j = await attCall({ action: "verify", pending: b.pending, code: b.code });
        BYB_AUTH = j.auth; setBybLogin(null); runUpdate();
      }
    } catch (e) { setUpdErr(e.message); }
  };

  const saveNote = async (course, text) => {
    const { error } = await supabase.from("course_stats").update({ note: text }).eq("course", course);
    if (error) { alert("저장 실패: " + error.message); return false; }
    setStats((s) => s.map((x) => x.course === course ? { ...x, note: text } : x));
    return true;
  };

  const live = stats.filter((s) => s.synced_at);
  const under = live.filter(isUnder);
  const over = live.filter(isOver);
  const noted = stats.filter((s) => (s.note || "").trim());
  const synced = live.reduce((a, s) => (s.synced_at > a ? s.synced_at : a), "");
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
          <p className="text-sm text-slate-500">{kstDay()} · {me.dept}</p>
        </div>
        {isAI && (
          <button onClick={runUpdate} disabled={job?.status === "working"}
            className="shrink-0 flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-medium px-3 py-2">
            {job?.status === "working"
              ? <><Loader2 size={13} className="animate-spin" /> {job.label}{job.total ? ` ${job.done}/${job.total}` : "…"}</>
              : `수강인원 업데이트 (${curM}월 창구)`}
          </button>
        )}
      </div>

      {/* 바이비 로그인 (업데이트에 필요, 새로고침 전까지 유지) */}
      {isAI && bybLogin && (
        <div className="bg-white border border-byb-line p-4 space-y-2">
          <div className="flex items-center justify-between">
            <p className="text-sm font-bold">바이비 로그인 <span className="font-normal text-slate-400 text-xs">· 인원을 바이비에서 직접 가져옵니다</span></p>
            <button onClick={() => setBybLogin(null)} className="text-slate-400"><X size={16} /></button>
          </div>
          {bybLogin.step === "login" ? (
            <div className="flex flex-wrap gap-2">
              <input value={bybLogin.user} onChange={(e) => setBybLogin({ ...bybLogin, user: e.target.value })} placeholder="바이비 아이디" className={`${attInput} flex-1 min-w-[10rem]`} />
              <input type="password" value={bybLogin.pw} onChange={(e) => setBybLogin({ ...bybLogin, pw: e.target.value })}
                onKeyDown={(e) => e.key === "Enter" && bybStep()} placeholder="비밀번호" className={`${attInput} flex-1 min-w-[10rem]`} />
              <button onClick={bybStep} className="rounded-lg bg-emerald-600 text-white text-xs font-medium px-4">로그인</button>
            </div>
          ) : (
            <div className="flex flex-wrap gap-2">
              <input value={bybLogin.code} onChange={(e) => setBybLogin({ ...bybLogin, code: e.target.value })}
                onKeyDown={(e) => e.key === "Enter" && bybStep()} placeholder="휴대폰 인증번호" autoFocus className={`${attInput} flex-1`} />
              <button onClick={bybStep} className="rounded-lg bg-emerald-600 text-white text-xs font-medium px-4">확인 후 업데이트</button>
            </div>
          )}
        </div>
      )}
      {updErr && <p className="text-xs text-rose-600 bg-rose-50 border border-rose-200 px-3 py-2">업데이트 실패: {updErr}</p>}

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
        {synced ? `마지막 업데이트 ${synced} · 바이비 창구별 확정 수강생 기준` : "아직 업데이트한 적이 없습니다"}
        {!isAI && " (업데이트는 foreon4 계정)"}
      </p>

      {/* 선택한 카드의 목록 */}
      {view && (
        <div className="bg-white border border-byb-line p-4">
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

      {/* 강좌별 수강 현황 — 종목 → 요일/시간/강사/반 → 강좌 순으로 눌러 들어가기 */}
      <EnrollBoard stats={stats} hist={hist} histErr={histErr} loading={loading} />

      {/* 다가오는 행사 */}
      <div className="bg-white border border-byb-line p-5">
        <h3 className="font-bold flex items-center gap-1.5 mb-3"><CalendarDays size={17} className="text-emerald-600" /> 다가오는 행사·일정</h3>
        {upcoming.length ? (
          <div className="space-y-2">
            {upcoming.map((e) => (
              <div key={e.id} className="flex items-start gap-3 border-b border-slate-50 pb-2 last:border-0">
                <span className="text-xs bg-emerald-50 text-emerald-700 rounded px-2 py-1 shrink-0">{e.date?.slice(5)}</span>
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

/* ─────────────────────── 강좌별 수강 현황 (홈) ───────────────────────
   종목(줌바·요가·수영…) → 요일/시간/반/강사 → 개별 강좌 순으로 눌러 들어가며
   현재 인원·정원·충원율과 추이(course_stats_history 일별 스냅샷)를 보여줍니다. */
const kstDay = (offset = 0) => new Date(Date.now() + 9 * 3600e3 + offset * 86400e3).toISOString().slice(0, 10);
const CAT_LABEL = { 매트: "매트필라테스", 방송: "방송댄스", 축구: "축구교실", 농구: "농구교실" };
const SERIES = ["#2a78d6", "#eb6834", "#1baf7a", "#eda100", "#e87ba4", "#008300"];
const DIMS = [{ id: "days", label: "요일별" }, { id: "time", label: "시간별" }, { id: "lvl", label: "반별" }, { id: "teacher", label: "강사별" }];
const ord = (arr, v) => { const i = arr.indexOf(v); return i < 0 ? 99 : i; };
const dimSort = (d) => d === "days" ? (a, b) => ord(DAY_ORDER, a) - ord(DAY_ORDER, b)
  : d === "time" ? (a, b) => timeMin(a) - timeMin(b)
    : (a, b) => a.localeCompare(b, "ko", { numeric: true });
const dimText = (d, k) => !k ? "(구분 없음)" : d === "days" ? dayLabel(k) : k;
const shortName = (full) => String(full).replace(/^\s*\d단지\s*/, "");
const toneOf = (en, cap) => !cap ? "none" : en >= cap ? "over" : en < Math.ceil(cap * UNDER_RATIO) ? "under" : "ok";

const _parsed = new Map();
function parseCourse(full) {
  const s = String(full || "");
  if (_parsed.has(s)) return _parsed.get(s);
  const inner = (s.match(/\(([^)]*)\)/) || [, ""])[1];
  const parts = inner.split("/").map((p) => p.trim());
  const dtp = parts.find((p) => /^(화목|수금|월|화|수|목|금|토|일)/.test(p)) || "";
  const dm = dtp.match(/^(화목|수금|월|화|수|목|금|토|일)(.*)$/);
  const head = s.replace(/^\s*\d단지\s*/, "").split("(")[0].trim();
  let sport = head, lvl = "";
  const di = head.indexOf("-");
  if (di > -1) { lvl = head.slice(di + 1).trim(); sport = head.slice(0, di).trim(); }
  if (parts.length >= 3) lvl = parts[0];
  if (/수영/.test(sport)) lvl = `${sport.replace("수영", "")} ${lvl}`.trim();   // 성인 기초, 어린이 상급…
  const cat = LESSON_CATS.find((k) => s.includes(k === "아쿠아로빅" ? "아쿠아" : k)) || sport || "기타";
  const r = {
    full: s, cat, lvl, days: dm ? dm[1] : "", time: dm ? dm[2].trim() : "",
    teacher: parts.length > 1 ? parts[parts.length - 1] : "",
  };
  _parsed.set(s, r);
  return r;
}

function FillRow({ label, sub, en, cap, under, over, single, active, onClick }) {
  const pct = cap ? Math.round((en / cap) * 100) : 0;
  const tone = toneOf(en, cap);
  const bar = { over: "bg-rose-500", under: "bg-amber-400", ok: "bg-emerald-600", none: "bg-slate-300" }[tone];
  return (
    <button onClick={onClick}
      className={`w-full text-left px-3 py-2 border-b border-byb-line last:border-0 transition ${active ? "bg-emerald-50" : "hover:bg-slate-50"}`}>
      <div className="flex items-baseline justify-between gap-2">
        <span className="min-w-0 truncate">
          <span className="text-[13px] font-medium text-slate-800">{label}</span>
          <span className="text-[11px] text-slate-400 ml-1.5">{sub}</span>
        </span>
        <span className="shrink-0 text-[13px] tabular-nums">
          <b className="text-slate-900 text-[15px]">{en}</b><span className="text-slate-400"> / {cap}명</span>
          <span className="inline-block w-11 text-right text-slate-600 font-medium">{pct}%</span>
        </span>
      </div>
      <div className="flex items-center gap-2 mt-1">
        <div className="relative flex-1 h-2 bg-slate-100 overflow-hidden rounded-sm">
          <div className={`h-full ${bar}`} style={{ width: `${Math.min(100, pct)}%` }} />
          <div className="absolute top-0 bottom-0 w-px bg-white" style={{ left: `${UNDER_RATIO * 100}%` }} />
        </div>
        <span className="shrink-0 text-[11px] space-x-2 w-[7.5rem] text-right">
          {single
            ? tone === "over" ? <span className="text-rose-600 font-medium">▲ 정원 초과</span>
              : tone === "under" ? <span className="text-amber-600 font-medium">▼ 미달 (여석 {cap - en})</span>
                : <span className="text-slate-400">여석 {Math.max(0, cap - en)}</span>
            : <>
              {over > 0 && <span className="text-rose-600 font-medium">▲ 초과 {over}</span>}
              {under > 0 && <span className="text-amber-600 font-medium">▼ 미달 {under}</span>}
            </>}
        </span>
      </div>
    </button>
  );
}

function TrendTip({ active, payload, label }) {
  if (!active || !payload?.length) return null;
  const sum = payload.reduce((a, p) => a + (p.value || 0), 0);
  return (
    <div className="bg-white border border-byb-line px-3 py-2 text-xs shadow min-w-[9rem]">
      <p className="font-bold text-slate-800 mb-1">{label}</p>
      {payload.map((p) => (
        <p key={p.dataKey} className="flex items-center gap-1.5 py-0.5">
          <span className="inline-block w-3 h-0.5" style={{ background: p.color }} />
          <span className="text-slate-500 truncate max-w-[10rem]">{p.name}</span>
          <b className="ml-auto pl-3 text-slate-800">{p.value}명</b>
        </p>
      ))}
      {payload.length > 1 && (
        <p className="flex border-t border-byb-line mt-1 pt-1"><span className="text-slate-500">합계</span><b className="ml-auto text-slate-800">{sum}명</b></p>
      )}
    </div>
  );
}

function EnrollBoard({ stats, hist, histErr, loading }) {
  const [cat, setCat] = useState(null);       // 종목
  const [dim, setDim] = useState("days");     // 2단계 분류 기준
  const [key, setKey] = useState(null);       // 2단계 선택값 (예: 화목)
  const [course, setCourse] = useState(null); // 개별 강좌
  const [period, setPeriod] = useState("day"); // day | month

  const rows = useMemo(() => stats
    .filter((s) => s.synced_at || s.enrolled)   // 특이사항만 적어둔 빈 줄 제외
    .map((s) => ({ ...parseCourse(s.course), enrolled: s.enrolled || 0, cap: capacityOf(s.course) })), [stats]);

  const agg = (list) => ({
    en: list.reduce((a, r) => a + r.enrolled, 0),
    cap: list.reduce((a, r) => a + r.cap, 0),
    n: list.length,
    under: list.filter((r) => toneOf(r.enrolled, r.cap) === "under").length,
    over: list.filter((r) => toneOf(r.enrolled, r.cap) === "over").length,
  });
  const groupBy = (list, fn, sort) => {
    const m = new Map();
    for (const r of list) { const k = fn(r); if (!m.has(k)) m.set(k, []); m.get(k).push(r); }
    return [...m.entries()].sort((a, b) => sort(a[0], b[0])).map(([k, l]) => ({ k, list: l, ...agg(l) }));
  };

  const inCat = cat ? rows.filter((r) => r.cat === cat) : rows;
  const dims = cat ? DIMS.filter((d) => new Set(inCat.map((r) => r[d.id] || "")).size > 1) : [];
  const curDim = dims.some((d) => d.id === dim) ? dim : (dims[0]?.id || "days");
  const inKey = key === null ? inCat : inCat.filter((r) => (r[curDim] || "") === key);
  const showCourses = !!cat && (key !== null || !dims.length);
  const catName = cat ? (CAT_LABEL[cat] || cat) : "";

  const items = !cat
    ? groupBy(rows, (r) => r.cat, (a, b) => ord(LESSON_CATS, a) - ord(LESSON_CATS, b))
      .map((g) => ({ ...g, label: CAT_LABEL[g.k] || g.k, sub: `강좌 ${g.n}개` }))
    : !showCourses
      ? groupBy(inCat, (r) => r[curDim] || "", dimSort(curDim))
        .map((g) => ({ ...g, label: dimText(curDim, g.k), sub: `강좌 ${g.n}개` }))
      : [...inKey].sort((a, b) => ord(DAY_ORDER, a.days) - ord(DAY_ORDER, b.days) || timeMin(a.time) - timeMin(b.time)
        || a.full.localeCompare(b.full, "ko"))
        .map((r) => ({ k: r.full, list: [r], ...agg([r]), single: true, label: shortName(r.full),
          sub: [dayLabel(r.days), r.time, r.lvl, r.teacher].filter(Boolean).join(" · ") }));

  const pickItem = (it) => {
    if (!cat) { setCat(it.k); setKey(null); setCourse(null); }
    else if (!showCourses) { setKey(it.k); setCourse(null); }
    else setCourse(course === it.k ? null : it.k);
  };

  const scope = course ? inKey.filter((r) => r.full === course) : inKey;
  const sum = agg(scope);
  const scopeName = course ? shortName(course)
    : key !== null ? `${catName} ${dimText(curDim, key)}` : cat ? catName : "전체 강좌";

  /* 추이 — 이름으로 분류하므로 강좌명이 바뀐 과거 기록도 종목·요일별 합계에 포함됩니다 */
  const inScope = (p) => (!cat || p.cat === cat) && (key === null || (p[curDim] || "") === key);
  let series;
  if (course) series = [{ name: shortName(course), test: (p) => p.full === course }];
  else if (!cat) series = [{ name: "전체 수강생", test: () => true }];
  else {
    const kids = items.map((it) => ({
      name: it.label,
      test: showCourses ? (p) => p.full === it.k : (p) => p.cat === cat && (p[curDim] || "") === it.k,
    }));
    series = kids.length >= 2 && kids.length <= SERIES.length ? kids : [{ name: scopeName, test: inScope }];
  }
  const tooMany = !course && cat && items.length > SERIES.length;

  const snaps = period === "day"
    ? hist.filter((h) => h.snap_date >= kstDay(-60))
    : [...hist.reduce((m, h) => m.set(h.snap_date.slice(0, 7), h), new Map()).values()];  // 월말(마지막 기록) 기준
  const points = snaps.map((h) => {
    const pt = { d: period === "day" ? h.snap_date.slice(5).replace("-", "/") : `${h.snap_date.slice(2, 4)}.${h.snap_date.slice(5, 7)}` };
    series.forEach((s, i) => { pt[`s${i}`] = 0; });
    for (const [name, n] of Object.entries(h.data || {})) {
      const p = parseCourse(name);
      series.forEach((s, i) => { if (s.test(p)) pt[`s${i}`] += Number(n) || 0; });
    }
    pt.total = series.reduce((a, _, i) => a + pt[`s${i}`], 0);
    return pt;
  });
  const delta = points.length >= 2 ? points.at(-1).total - points.at(-2).total : null;
  const noTable = /course_stats_history|does not exist|schema cache/i.test(histErr);

  const crumbs = [
    { label: "전체", go: () => { setCat(null); setKey(null); setCourse(null); } },
    cat && { label: catName, go: () => { setKey(null); setCourse(null); } },
    key !== null && { label: dimText(curDim, key), go: () => setCourse(null) },
    course && { label: shortName(course), go: () => {} },
  ].filter(Boolean);

  const pct = sum.cap ? Math.round((sum.en / sum.cap) * 100) : 0;
  const kpis = [
    { label: "현재 수강생", value: `${sum.en.toLocaleString()}명`,
      foot: delta === null ? "추이 기록 대기" : `${period === "day" ? "직전 기록" : "전월"} 대비 ${delta > 0 ? "+" : ""}${delta}명` },
    { label: "정원", value: `${sum.cap.toLocaleString()}명`, foot: `여석 ${Math.max(0, sum.cap - sum.en)}명` },
    { label: "충원율", value: `${pct}%`, foot: `미달 기준 ${UNDER_RATIO * 100}%` },
    { label: "강좌 수", value: `${sum.n}개`, foot: `미달 ${sum.under} · 초과 ${sum.over}` },
  ];

  return (
    <div className="bg-white border border-byb-line">
      <div className="px-4 py-3 border-b border-byb-line flex flex-wrap items-center justify-between gap-2">
        <h3 className="font-bold text-[14px] text-slate-800 flex items-center gap-1.5">
          <GraduationCap size={16} className="text-emerald-600" /> 강좌별 수강 현황
        </h3>
        <nav className="flex flex-wrap items-center gap-1 text-[12px]">
          {crumbs.map((c, i) => (
            <React.Fragment key={i}>
              {i > 0 && <span className="text-slate-300">›</span>}
              <button onClick={c.go} disabled={i === crumbs.length - 1}
                className={i === crumbs.length - 1 ? "font-bold text-slate-800" : "text-emerald-700 hover:underline"}>{c.label}</button>
            </React.Fragment>
          ))}
        </nav>
      </div>

      {loading ? <p className="text-sm text-slate-400 py-10 text-center">불러오는 중…</p>
        : !rows.length ? <p className="text-sm text-slate-400 py-10 text-center">아직 수강인원 자료가 없습니다 — [수강인원 업데이트] 를 눌러주세요.</p>
          : (
            <div className="p-4 space-y-4">
              {/* 선택 범위 요약 */}
              <div>
                <p className="text-[12px] text-slate-500 mb-1.5"><b className="text-slate-800">{scopeName}</b> 기준</p>
                <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                  {kpis.map((k) => (
                    <div key={k.label} className="border border-byb-line px-3 py-2.5">
                      <p className="text-[11px] text-slate-500">{k.label}</p>
                      <p className="text-xl font-bold text-slate-900 mt-0.5 tabular-nums">{k.value}</p>
                      <p className="text-[11px] text-slate-400 mt-0.5">{k.foot}</p>
                    </div>
                  ))}
                </div>
              </div>

              <div className="grid lg:grid-cols-2 gap-4">
                {/* 목록 — 눌러서 한 단계씩 들어가기 */}
                <div className="border border-byb-line">
                  <div className="px-3 py-2 bg-byb-bg border-b border-byb-line flex flex-wrap items-center gap-1.5">
                    <span className="text-[12px] font-bold text-slate-700 mr-1">
                      {!cat ? "종목별" : showCourses ? "강좌별" : `${catName} · ${DIMS.find((d) => d.id === curDim)?.label}`}
                    </span>
                    {cat && key === null && dims.map((d) => (
                      <button key={d.id} onClick={() => setDim(d.id)}
                        className={`text-[11px] px-2 py-0.5 border ${curDim === d.id ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-byb-line text-slate-600 hover:border-emerald-300"}`}>{d.label}</button>
                    ))}
                    <span className="ml-auto text-[11px] text-slate-400">{showCourses ? "누르면 그래프에 표시" : "누르면 세부 보기"}</span>
                  </div>
                  <div className="max-h-[32rem] overflow-y-auto">
                    {items.map((it) => (
                      <FillRow key={it.k} label={it.label} sub={it.sub} en={it.en} cap={it.cap} under={it.under} over={it.over}
                        single={it.single} active={course === it.k} onClick={() => pickItem(it)} />
                    ))}
                  </div>
                </div>

                {/* 추이 그래프 */}
                <div className="border border-byb-line p-3 flex flex-col">
                  <div className="flex items-center justify-between gap-2 mb-2">
                    <p className="text-[12px] font-bold text-slate-700">수강생 추이 <span className="font-normal text-slate-400">· {scopeName}</span></p>
                    <div className="flex border border-byb-line text-[11px]">
                      {[["day", "일별"], ["month", "월별"]].map(([id, lb]) => (
                        <button key={id} onClick={() => setPeriod(id)}
                          className={`px-2.5 py-1 ${period === id ? "bg-emerald-600 text-white" : "bg-white text-slate-600"}`}>{lb}</button>
                      ))}
                    </div>
                  </div>
                  {series.length > 1 && (
                    <div className="flex flex-wrap gap-x-3 gap-y-1 mb-2">
                      {series.map((s, i) => (
                        <span key={i} className="flex items-center gap-1 text-[11px] text-slate-600">
                          <span className="inline-block w-3 h-[3px] rounded-full" style={{ background: SERIES[i] }} />{s.name}
                        </span>
                      ))}
                    </div>
                  )}
                  {noTable ? (
                    <p className="text-xs text-slate-500 bg-byb-bg p-3 my-auto leading-relaxed">
                      추이를 저장할 표(<b>course_stats_history</b>)가 아직 없습니다.<br />Supabase SQL Editor 에서 생성 SQL 을 한 번 실행해주세요.
                    </p>
                  ) : !points.length ? (
                    <p className="text-xs text-slate-400 text-center my-auto py-10">아직 쌓인 기록이 없습니다.</p>
                  ) : (
                    <>
                      <div className="h-56">
                        <ResponsiveContainer width="100%" height="100%">
                          <LineChart data={points} margin={{ top: 8, right: 12, left: 0, bottom: 0 }}>
                            <CartesianGrid vertical={false} stroke="#eef0f3" />
                            <XAxis dataKey="d" tick={{ fontSize: 11, fill: "#888" }} tickLine={false} axisLine={{ stroke: "#e3e6ea" }} minTickGap={16} />
                            <YAxis tick={{ fontSize: 11, fill: "#888" }} tickLine={false} axisLine={false} allowDecimals={false} width={36} />
                            <Tooltip content={<TrendTip />} cursor={{ stroke: "#c8ccd2", strokeDasharray: "3 3" }} />
                            {series.map((s, i) => (
                              <Line key={i} type="linear" dataKey={`s${i}`} name={s.name} stroke={SERIES[i]} strokeWidth={2}
                                dot={points.length <= 14 ? { r: 3, strokeWidth: 2, stroke: "#fff", fill: SERIES[i] } : false}
                                activeDot={{ r: 5, strokeWidth: 2, stroke: "#fff" }} isAnimationActive={false} />
                            ))}
                          </LineChart>
                        </ResponsiveContainer>
                      </div>
                      <p className="text-[11px] text-slate-400 mt-1">
                        {points.length < 2 ? "기록이 쌓이면 추이가 그려집니다 · " : ""}
                        홈 화면을 열 때마다 그날 인원이 자동 저장됩니다{period === "month" ? " (월별 = 그 달 마지막 기록)" : " (최근 60일)"}
                        {tooMany ? " · 항목이 많아 합계만 표시 — 목록을 눌러 좁혀보세요" : ""}
                      </p>
                    </>
                  )}
                </div>
              </div>
            </div>
          )}
    </div>
  );
}

/* 업로드 파일 → base64 */
const fileToB64 = (file) => new Promise((res, rej) => {
  const r = new FileReader();
  r.onload = (e) => res(e.target.result.split(",")[1]);
  r.onerror = () => rej(new Error("파일을 읽지 못했습니다"));
  r.readAsDataURL(file);
});


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


/* ─────────────────────── 자료실 (공유 폴더)
   Supabase Storage 의 'shared' 버킷을 씁니다.
   파일 이름에 올린 사람과 시각을 함께 담아, 표 하나 더 만들지 않고도
   누가 언제 올렸는지 보이도록 했습니다.                                */
const FILE_BUCKET = "shared";
const FILE_FOLDERS = ["공지·안내", "서식·양식", "사진", "기타"];

const encName = (folder, uploader, name) =>
  `${folder}/${Date.now()}__${uploader}__${name.replace(/[/\\]/g, "_")}`;

function decName(path) {
  const base = String(path).split("/").pop();
  const m = base.match(/^(\d+)__([^_]+)__(.+)$/);
  if (!m) return { name: base, by: "", at: "" };
  const d = new Date(Number(m[1]));
  const two = (n) => String(n).padStart(2, "0");
  return {
    name: m[3], by: m[2],
    at: `${d.getFullYear()}-${two(d.getMonth() + 1)}-${two(d.getDate())} ${two(d.getHours())}:${two(d.getMinutes())}`,
  };
}

const sizeText = (b) => {
  if (!b && b !== 0) return "";
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${Math.round(b / 1024)} KB`;
  return `${(b / 1024 / 1024).toFixed(1)} MB`;
};

function SharedFiles({ me }) {
  const [folder, setFolder] = useState(FILE_FOLDERS[0]);
  const [rows, setRows] = useState([]);
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const [err, setErr] = useState("");
  const inputRef = React.useRef(null);

  const load = React.useCallback(async (f) => {
    setBusy(true); setErr("");
    const { data, error } = await supabase.storage.from(FILE_BUCKET)
      .list(f, { limit: 200, sortBy: { column: "name", order: "desc" } });
    setBusy(false);
    if (error) {
      setErr(error.message.includes("not found")
        ? `저장 공간(${FILE_BUCKET})이 아직 없습니다. Supabase → Storage 에서 '${FILE_BUCKET}' 버킷을 만들어주세요.`
        : "목록을 불러오지 못했습니다: " + error.message);
      setRows([]); return;
    }
    setRows((data || []).filter((r) => r.name && !r.name.startsWith(".")));
  }, []);

  React.useEffect(() => { load(folder); }, [folder, load]);

  const upload = async (fileList) => {
    const files = [...(fileList || [])];
    if (!files.length) return;
    setBusy(true); setErr(""); setMsg("");
    let ok = 0;
    for (const f of files) {
      if (f.size > 50 * 1024 * 1024) { setErr(`${f.name} 은 50MB 를 넘어 올리지 못했습니다.`); continue; }
      const { error } = await supabase.storage.from(FILE_BUCKET)
        .upload(encName(folder, me.id, f.name), f, { upsert: false });
      if (error) setErr(`${f.name} 올리기 실패: ${error.message}`); else ok++;
    }
    setBusy(false);
    if (ok) setMsg(`${ok}개 파일을 올렸습니다.`);
    if (inputRef.current) inputRef.current.value = "";
    load(folder);
  };

  const open = async (row) => {
    const { data, error } = await supabase.storage.from(FILE_BUCKET)
      .createSignedUrl(`${folder}/${row.name}`, 60 * 10);
    if (error || !data) { setErr("내려받기 실패: " + (error?.message || "")); return; }
    window.open(data.signedUrl, "_blank");
  };

  const remove = async (row) => {
    const info = decName(row.name);
    if (!window.confirm(`'${info.name}' 을 지울까요?`)) return;
    setBusy(true);
    const { error } = await supabase.storage.from(FILE_BUCKET).remove([`${folder}/${row.name}`]);
    setBusy(false);
    if (error) setErr("삭제 실패: " + error.message); else load(folder);
  };

  return (
    <div className="bg-white border border-byb-line p-5 space-y-4">
      <div className="flex items-center justify-between">
        <h2 className="font-bold">자료실</h2>
        <button onClick={() => load(folder)} className="text-xs text-slate-500 flex items-center gap-1">
          <RefreshCw size={13} /> 새로고침
        </button>
      </div>

      <div className="flex flex-wrap gap-1.5">
        {FILE_FOLDERS.map((f) => (
          <button key={f} type="button" onClick={() => setFolder(f)}
            className={`text-sm px-3 py-1.5 rounded-full border transition ${folder === f ? "bg-emerald-600 border-emerald-600 text-white font-medium" : "bg-white border-slate-200 text-slate-600 hover:border-emerald-300"}`}>
            {f}
          </button>
        ))}
      </div>

      <label className="block border-2 border-dashed border-slate-200 rounded-xl p-4 text-center cursor-pointer hover:border-emerald-300"
        onDragOver={(e) => e.preventDefault()}
        onDrop={(e) => { e.preventDefault(); upload(e.dataTransfer.files); }}>
        <input ref={inputRef} type="file" multiple className="hidden"
          onChange={(e) => upload(e.target.files)} />
        <Upload size={18} className="mx-auto text-slate-400 mb-1" />
        <p className="text-sm text-slate-600">파일을 끌어다 놓거나 눌러서 올리기</p>
        <p className="text-[11px] text-slate-400 mt-0.5">여러 개 한꺼번에 가능 · 한 개당 50MB 까지</p>
      </label>

      {msg && <p className="text-xs text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg p-2">{msg}</p>}
      {err && <p className="text-xs text-rose-700 bg-rose-50 border border-rose-200 rounded-lg p-2">{err}</p>}

      {busy && <p className="text-xs text-slate-400 flex items-center gap-1"><Loader2 size={13} className="animate-spin" /> 처리 중…</p>}

      {rows.length ? (
        <div className="border border-slate-200 rounded-lg divide-y divide-slate-100">
          {rows.map((r) => {
            const info = decName(r.name);
            const mine = info.by === me.id || me.dept === "사무실";
            return (
              <div key={r.name} className="flex items-center gap-2 px-3 py-2.5">
                <FileText size={15} className="text-slate-400 shrink-0" />
                <div className="flex-1 min-w-0">
                  <p className="text-sm truncate">{info.name}</p>
                  <p className="text-[11px] text-slate-400">
                    {info.by && `${info.by} · `}{info.at}
                    {r.metadata?.size ? ` · ${sizeText(r.metadata.size)}` : ""}
                  </p>
                </div>
                <button onClick={() => open(r)} title="내려받기"
                  className="text-slate-400 hover:text-emerald-600 p-1.5"><Download size={15} /></button>
                {mine && (
                  <button onClick={() => remove(r)} title="삭제"
                    className="text-slate-300 hover:text-rose-600 p-1.5"><Trash2 size={15} /></button>
                )}
              </div>
            );
          })}
        </div>
      ) : (!busy && !err && <p className="text-xs text-slate-400">아직 올린 파일이 없습니다.</p>)}

      <p className="text-[11px] text-slate-400">
        올린 사람 본인과 사무실 계정만 삭제할 수 있습니다. 개인정보가 담긴 파일은 올리지 말아주세요.
      </p>
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
  const [stats, setStats] = useState([]);
  const keyOf = (t) => String(t).replace(/[^0-9A-Za-z가-힣]/g, "");
  const loadCounts = React.useCallback(async () => {
    setStatsMsg("인원 불러오는 중…");
    const { data, error } = await supabase.from("course_stats").select("course,enrolled,capacity");
    if (error) { setStatsMsg("인원을 불러오지 못했습니다: " + error.message); return; }
    const list = data || [];
    const map = {};
    for (const r of list) map[keyOf(r.course)] = r.enrolled;
    setStats(list);
    setCounts(map);
    setStatsMsg(list.length
      ? ""
      : "인원 숫자가 없습니다 — 홈 화면에서 [수강인원 업데이트] 를 먼저 눌러주세요.");
  }, []);
  React.useEffect(() => { loadCounts(); }, [loadCounts]);
  const countOf = (full) => counts[keyOf(full)];

  const isSwim = group === "수영" || group === "아쿠아";
  const flat = (t) => String(t).replace(/[\s\-()/·.,]/g, "");
  const words = group.trim().split(/\s+/).map(flat).filter(Boolean);

  /* 대상 강좌는 '수강인원 업데이트' 로 가져온 실제 강좌(course_stats) 를 먼저 쓴다.
     → 강좌명이 바뀌어도 인원이 항상 표시되고, 이월 검색어도 실제 이름과 일치한다.
     아직 한 번도 갱신하지 않았다면 앱에 적어둔 목록으로 대신한다. */
  const dbTargets = words.length
    ? stats.filter((s) => words.every((w) => flat(s.course).includes(w)))
        .map((s) => ({ full: s.course, enrolled: s.enrolled, capacity: s.capacity }))
        .sort((a, b) => a.full.localeCompare(b.full, "ko"))
    : [];
  const listTargets = words.length
    ? ALL_COURSES.filter((c) => words.every((w) => flat(c.full).includes(w)))
        .map((c) => ({ full: c.full, enrolled: countOf(c.full), capacity: undefined }))
    : [];
  const fromDb = dbTargets.length > 0;
  const targets = fromDb ? dbTargets : listTargets;
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
    }, 1500);
    return () => clearInterval(t);
  }, [job, me.id]);

  /* PC 작업자가 사람 손을 기다리는 중이면 progress 가 '⏸' 로 시작한다 */
  const waitText = job && typeof job.progress === "string" && job.progress.startsWith("⏸")
    ? job.progress.slice(1).trim() : "";

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
    <div className="bg-white border border-byb-line p-5 space-y-4">
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
                const n = c.enrolled;
                const cap = c.capacity || capacityOf(c.full);
                return (
                  <button key={c.full} type="button" onClick={() => toggle(c.full)}
                    className={`flex items-center gap-2 w-full text-left px-3 py-2 text-sm ${off ? "bg-slate-50 text-slate-400 line-through" : "hover:bg-emerald-50"}`}>
                    <span className={`w-4 h-4 rounded border flex items-center justify-center shrink-0 ${off ? "border-slate-300" : "bg-emerald-600 border-emerald-600"}`}>
                      {!off && <CheckCircle2 size={12} className="text-white" />}
                    </span>
                    <span className="flex-1">{c.full}</span>
                    <span className={`shrink-0 text-xs font-semibold px-2 py-0.5 rounded-full ${off ? "bg-slate-100 text-slate-400" : n === undefined ? "bg-slate-100 text-slate-400" : n === 0 ? "bg-slate-100 text-slate-500" : "bg-emerald-50 text-emerald-700"}`}>
                      {n === undefined ? "–" : `${n}/${cap}`}
                    </span>
                  </button>
                );
              })}
            </div>
          ) : <p className="text-xs text-slate-400">해당하는 강좌가 없습니다.</p>}
          {!fromDb && targets.length > 0 && (
            <p className="text-[11px] text-amber-700 bg-amber-50 border border-amber-200 rounded p-2 mt-1.5">
              앱에 적어둔 강좌 목록으로 보여주는 중입니다. 홈 화면에서 [수강인원 업데이트] 를 한 번 눌러주시면
              실제 강좌와 인원으로 바뀝니다.
            </p>
          )}
          <p className="text-[11px] text-slate-400 mt-1.5">수강생이 0명인 강좌는 자동으로 건너뜁니다. 반변경자는 기존 반에서 <b>제외만</b> 되며, 새 반 등록은 직접 해주세요.</p>

        </div>
      )}

      <label className="flex items-center gap-2 text-sm text-slate-600 cursor-pointer select-none">
        <input type="checkbox" checked={safe} onChange={(e) => setSafe(e.target.checked)}
          className="w-4 h-4 rounded border-slate-300 accent-emerald-600" />
        안전모드 (강좌마다 [예약 생성] 은 PC에서 직접 클릭)
      </label>
      <p className="text-[11px] text-slate-500 bg-slate-50 border border-slate-200 rounded-lg p-2 -mt-1">
        <b>'강좌 월 선택' 창의 월 선택과 [확인] 은 항상 PC에서 직접</b> 눌러주세요.
        그 외 강좌 검색 · 명단 전체 선택 · 반변경자 제외는 자동으로 진행됩니다.
        차례가 되면 이 화면에 알려드립니다.
      </p>

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

      {waitText && (
        <div className="rounded-xl border-2 border-amber-400 bg-amber-50 p-3 animate-pulse">
          <p className="text-xs font-bold text-amber-900 mb-0.5">⏸ 지금 PC 브라우저에서 직접 눌러주세요</p>
          <p className="text-sm text-amber-900">{waitText}</p>
          <p className="text-[11px] text-amber-700 mt-1">누르시면 나머지는 자동으로 이어집니다.</p>
        </div>
      )}

      {job && (
        <div className={`rounded-lg p-3 text-sm ${job.status === "error" ? "bg-rose-50 border border-rose-200 text-rose-700" : "bg-slate-50 border border-slate-200"}`}>
          {job.progress && !waitText && <p className="font-medium">{job.progress}</p>}
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
      <div className="bg-white border border-byb-line p-5 space-y-4">
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
  const [gender, setGender] = useState("남자");   // 독서실: 남자석/여자석 구분
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
    // 독서실은 남자석/여자석이 나뉘어 있어 성별을 함께 보냅니다 (예: "남자 13")
    const seatSend = (kind === "독서실" && optLabel === "일반" && seat.trim())
      ? `${gender} ${seat.trim()}` : seat;
    setSaving(true);
    const at = nowLocal().replace("T", " ");
    // 기록 저장
    await logWork({ kind, action, target: productKey, seat: seatSend, member: name, dong, ho, amount: amt, timing: `${action === "등록" ? "등록일" : "해지일"} ${dateStr}${action === "취소" ? " · " + (timing === "endOfToday" ? "오늘24시" : "즉시") : ""}`, by: me.id, at });
    if (canAutomate) {
      const { error } = await supabase.from("lesson_jobs").insert({
        kind, action, course: productKey, member: name, dong, ho, seat: seatSend, amount: amt,
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

      <div className="bg-white border border-byb-line p-5 space-y-4">
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
          {kind === "독서실" && optLabel === "일반" && (
            <Fld label="남자석 / 여자석 *">
              <div className="flex gap-1.5">
                {["남자", "여자"].map((g) => (
                  <button key={g} type="button" onClick={() => setGender(g)}
                    className={`flex-1 text-sm py-2 rounded-lg border transition ${gender === g ? "bg-emerald-600 border-emerald-600 text-white font-medium" : "bg-white border-slate-200 text-slate-600"}`}>
                    {g}
                  </button>
                ))}
              </div>
            </Fld>
          )}
          <Fld label={seatLabel + " *"}>
            <input value={seat} onChange={(e) => setSeat(e.target.value)} placeholder={seatPh} className="rin" />
            {kind === "골프백" && (
              <p className="text-[11px] text-slate-400 mt-1">홀수 번호는 상단, 짝수 번호는 하단으로 자동 선택됩니다.</p>
            )}
          </Fld>
        </div>

        <div className="space-y-2">
          <div className="flex gap-2 items-end">
            <L label="동"><input className={attInput} inputMode="numeric" value={dong} onChange={(e) => { setDong(e.target.value); setPeople(null); setPerson(null); }} /></L>
            <L label="호수"><input className={attInput} inputMode="numeric" value={ho} onChange={(e) => { setHo(e.target.value); setPeople(null); setPerson(null); }}
              onKeyDown={(e) => e.key === "Enter" && dong && ho && findPeople()} /></L>
            <button onClick={findPeople} disabled={!dong || !ho || !!busy} className="rounded-lg border border-emerald-600 text-emerald-700 text-sm px-3 py-2 whitespace-nowrap disabled:opacity-40">회원 찾기</button>
          </div>
          {people && (!people.length ? <p className="text-xs text-slate-400">{dong}동 {ho}호에 등록된 회원이 없어요.</p> :
            <div className="border border-slate-200 rounded-lg divide-y divide-slate-100 max-h-48 overflow-y-auto">
              {people.map((pp) => (
                <button key={pp.rauthId} onClick={() => setPerson(pp)} className={`w-full text-left px-3 py-2 text-sm ${person === pp ? "bg-emerald-50 text-emerald-800 font-medium" : "hover:bg-slate-50"}`}>
                  {pp.name} <span className="text-xs text-slate-400">{pp.birth ? pp.birth + "년생" : ""} {pp.phone4 ? "· ****" + pp.phone4 : ""}</span>
                </button>
              ))}
            </div>)}
        </div>

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
        {["전체", "강습", "출석부", "신청서", "독서실", "골프백"].map((k) => (
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
                <p className="text-[11px] text-slate-400 mt-1.5">{r.at} · {r.by}</p>
              </div>
            ))}
          </div>}
    </div>
  );
}


/* ─────────────────────── 외부 연동 안내 */
function LinkInfo() {
  return (
    <div className="bg-white border border-byb-line p-5">
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

/* ─────────────────────── 출석부 생성 (바이비 연동) */
const ATT_KINDS = ["수영", "농구", "축구", "GX"];
const attKindOf = (n) => (n.includes("수영") ? "수영" : n.includes("농구") ? "농구" : n.includes("축구") ? "축구" : "GX");
const attInput = "w-full rounded-lg border border-slate-200 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500";

async function attCall(body) {
  const r = await fetch("/api/attendance", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  let j = {};
  try { j = await r.json(); } catch {}
  if (!r.ok) throw new Error(j.error || `서버 오류 (${r.status})`);
  return j;
}

const APPLY_KEYWORDS = ["농구", "축구", "2단지 요가", "2단지 줌바", "2단지 매트", "2단지 방송", "2단지 근력", "수영", "아쿠아", "2단지 타바타"];
let BYB_AUTH = null; // 바이비 로그인 토큰: 탭을 옮겨도 유지, 새로고침/닫으면 사라짐

function Attendance({ mode = "attendance", me }) {
  const isApply = mode === "apply";
  const next = new Date(new Date().getFullYear(), new Date().getMonth() + 1, 1);
  const [step, setStep] = useState(() => (BYB_AUTH ? "ready" : "login")); // login → otp → ready
  const [user, setUser] = useState(() => { try { return localStorage.getItem("byb_user") || ""; } catch { return ""; } });
  const [pw, setPw] = useState("");
  const [code, setCode] = useState("");
  const [pending, setPending] = useState(null); // 인증 대기 정보 (메모리에만 보관)
  const [auth, setAuthState] = useState(BYB_AUTH); // 바이비 로그인 토큰 (메모리에만 보관)
  const setAuth = (a) => { BYB_AUTH = a; setAuthState(a); };
  const [kws, setKws] = useState(() => Object.fromEntries(APPLY_KEYWORDS.map((k) => [k, true])));
  const [year, setYear] = useState(next.getFullYear());
  const [month, setMonth] = useState(next.getMonth() + 1);
  const [holi, setHoli] = useState("");
  const [kinds, setKinds] = useState({ 수영: true, 농구: true, 축구: true, GX: true });
  const [skipEmpty, setSkipEmpty] = useState(true);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [prog, setProg] = useState(null);
  const [log, setLog] = useState([]);

  const guard = async (fn) => {
    setBusy(true); setErr("");
    try { await fn(); }
    catch (e) {
      setErr(e.message);
      if (/만료|로그인 정보/.test(e.message)) { setAuth(null); setStep("login"); }
    } finally { setBusy(false); }
  };

  const doLogin = () => guard(async () => {
    const j = await attCall({ action: "login", user, password: pw });
    try { localStorage.setItem("byb_user", user); } catch {}
    setPw("");
    if (j.auth) { setAuth(j.auth); setStep("ready"); }
    else { setPending(j.pending); setStep("otp"); }
  });

  const doVerify = () => guard(async () => {
    const j = await attCall({ action: "verify", pending, code });
    setAuth(j.auth); setPending(null); setCode(""); setStep("ready");
  });

  const doRun = () => guard(async () => {
    setLog([]);
    setProg({ label: "강좌 목록 불러오는 중", done: 0, total: 1 });
    let cur = auth; // 5분마다 서버가 토큰을 새로 받아오면 교체
    const withAuth = async (body) => {
      const j = await attCall({ ...body, auth: cur });
      if (j.auth) { cur = j.auth; setAuth(j.auth); }
      return j;
    };
    const keywords = APPLY_KEYWORDS.filter((k) => kws[k]);
    if (isApply && !keywords.length) throw new Error("검색할 강좌를 하나 이상 골라 주세요.");
    // 신청서: N월 신청서 = (N-1)월 창구 명단
    const src = isApply ? new Date(year, month - 2, 1) : new Date(year, month - 1, 1);
    const sy = src.getFullYear(), sm = src.getMonth() + 1;
    const { windows } = await withAuth(isApply ? { action: "windows", keywords, year: sy, month: sm } : { action: "windows", year: sy, month: sm });
    const list = isApply ? windows : windows.filter((w) => kinds[attKindOf(w.name)]);
    if (!list.length) throw new Error("선택한 종류의 강좌가 없어요.");
    const chunks = [];
    for (let i = 0; i < list.length; i += 10) chunks.push(list.slice(i, i + 10));
    const got = {};
    let done = 0, idx = 0;
    setProg({ label: "강좌별 명단 수집 중", done: 0, total: list.length });
    const worker = async () => {
      while (idx < chunks.length) {
        const c = chunks[idx++];
        const { members } = await withAuth({ action: "members", year: sy, month: sm, windows: c });
        Object.entries(members).forEach(([n, m]) => {
          const prev = got[n] || [];
          got[n] = prev.concat(m.filter((x) => !prev.some((p) => p.name === x.name && p.phone === x.phone)));
        });
        done += c.length;
        setProg({ label: "강좌별 명단 수집 중", done, total: list.length });
      }
    };
    await Promise.all([worker(), worker(), worker()]);
    const names = [...new Set(list.map((w) => w.name))];
    const data = names.map((n) => [n, got[n] || []]).filter(([, m]) => !(isApply && skipEmpty && !m.length));
    if (!data.length) throw new Error("명단이 있는 강좌가 없어요.");
    setProg({ label: "엑셀 만드는 중", done: list.length, total: list.length });
    const holidays = holi.split(/[,\s]+/).map(Number).filter((n) => n >= 1 && n <= 31);
    const r = await attCall(isApply ? { action: "apply_build", year, month, data } : { action: "build", year, month, holidays, skipEmpty, data });
    const bin = atob(r.zip);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    const url = URL.createObjectURL(new Blob([bytes], { type: "application/zip" }));
    const a = document.createElement("a");
    a.href = url; a.download = `${isApply ? "신청서" : "출석부"}_${year}년${String(month).padStart(2, "0")}월.zip`; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    const total = data.reduce((s, [, m]) => s + m.length, 0);
    logWork({ kind: isApply ? "신청서" : "출석부", action: "생성", target: `${year}년 ${month}월`,
      seat: isApply ? APPLY_KEYWORDS.filter((k) => kws[k]).join(", ") : ATT_KINDS.filter((k) => kinds[k]).join(", "),
      member: `강좌 ${data.length}개 · ${total}명`, by: me?.id || "?", at: nowLocal().replace("T", " ") });
    setLog([`강좌 ${data.length}개 · 회원 ${total}명`, ...(r.log || [])]);
    setProg(null);
  });

  const btn = "rounded-lg bg-emerald-600 text-white px-4 py-2 text-sm font-medium hover:bg-emerald-700 disabled:opacity-50 flex items-center gap-1.5";

  return (
    <div>
      <Panel title={mode === "stats" ? "월간 이용 통계 (1~4단지 강좌 수강생 · 시설 이용자)" : mode === "lesson" ? "강습 등록·취소 (바이비)" : isApply ? "신청서 생성 (전월 창구 명단 → 신청서 엑셀)" : "출석부 생성 (바이비 명단 → 엑셀)"}>
        {step === "login" && (
          <div className="grid sm:grid-cols-3 gap-2 items-end">
            <L label="바이비 아이디"><input className={attInput} value={user} onChange={(e) => setUser(e.target.value)} /></L>
            <L label="비밀번호"><input type="password" className={attInput} value={pw} onChange={(e) => setPw(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && user && pw && doLogin()} /></L>
            <button className={btn} disabled={busy || !user || !pw} onClick={doLogin}>
              {busy ? <Loader2 size={15} className="animate-spin" /> : <LogIn size={15} />} 바이비 로그인
            </button>
          </div>
        )}
        {step === "otp" && (
          <div className="grid sm:grid-cols-3 gap-2 items-end">
            <L label="휴대폰으로 받은 인증번호"><input className={attInput} value={code} inputMode="numeric" onChange={(e) => setCode(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && code && doVerify()} /></L>
            <button className={btn} disabled={busy || !code} onClick={doVerify}>
              {busy ? <Loader2 size={15} className="animate-spin" /> : <CheckCircle2 size={15} />} 인증 확인
            </button>
            <button className="text-sm text-slate-500 underline" onClick={() => { setStep("login"); setPending(null); }}>처음부터</button>
          </div>
        )}
        {step === "ready" && mode === "stats" && (
          <>
            <p className="text-xs text-emerald-700 flex items-center gap-1 mb-3"><CheckCircle2 size={14} /> 바이비 로그인됨 ({user})</p>
            <StatsPanel call={async (body) => { const j = await attCall({ ...body, auth: BYB_AUTH }); if (j.auth) setAuth(j.auth); return j; }} />
          </>
        )}
        {step === "ready" && mode === "lesson" && (
          <>
            <p className="text-xs text-emerald-700 flex items-center gap-1 mb-3"><CheckCircle2 size={14} /> 바이비 로그인됨 ({user})</p>
            <LessonPanel me={me} call={async (body) => { const j = await attCall({ ...body, auth: BYB_AUTH }); if (j.auth) setAuth(j.auth); return j; }} />
          </>
        )}
        {step === "ready" && mode !== "lesson" && mode !== "stats" && (
          <div className="space-y-3">
            <p className="text-xs text-emerald-700 flex items-center gap-1"><CheckCircle2 size={14} /> 바이비 로그인됨 ({user}) — 이 창을 닫으면 로그인 정보는 사라져요.</p>
            <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
              <L label="년도"><input type="number" className={attInput} value={year} onChange={(e) => setYear(+e.target.value)} /></L>
              <L label="월">
                <select className={attInput} value={month} onChange={(e) => setMonth(+e.target.value)}>
                  {Array.from({ length: 12 }, (_, i) => <option key={i} value={i + 1}>{i + 1}월</option>)}
                </select>
              </L>
              {isApply
                ? <div className="col-span-2 text-xs text-slate-500 self-center">명단 기준: <b>{new Date(year, month - 2, 1).getMonth() + 1}월 창구</b> (바이비 확정 순서대로)</div>
                : <div className="col-span-2"><L label="공휴일 (쉬는 날, 쉼표로)"><input className={attInput} placeholder="예: 3, 9" value={holi} onChange={(e) => setHoli(e.target.value)} /></L></div>}
            </div>
            {isApply && (
              <div className="flex flex-wrap gap-3 text-sm">
                {APPLY_KEYWORDS.map((k) => (
                  <label key={k} className="flex items-center gap-1.5">
                    <input type="checkbox" checked={kws[k]} onChange={(e) => setKws({ ...kws, [k]: e.target.checked })} /> {k}
                  </label>
                ))}
                <label className="flex items-center gap-1.5 text-slate-500">
                  <input type="checkbox" checked={skipEmpty} onChange={(e) => setSkipEmpty(e.target.checked)} /> 0명 강좌 제외
                </label>
              </div>
            )}
            {!isApply && <div className="flex flex-wrap gap-3 text-sm">
              {ATT_KINDS.map((k) => (
                <label key={k} className="flex items-center gap-1.5">
                  <input type="checkbox" checked={kinds[k]} onChange={(e) => setKinds({ ...kinds, [k]: e.target.checked })} /> {k}{k === "수영" ? " (강사별 파일)" : ""}
                </label>
              ))}
              <label className="flex items-center gap-1.5 text-slate-500">
                <input type="checkbox" checked={skipEmpty} onChange={(e) => setSkipEmpty(e.target.checked)} /> 0명 강좌 제외
              </label>
            </div>}
            <button className={btn} disabled={busy} onClick={doRun}>
              {busy ? <Loader2 size={15} className="animate-spin" /> : <Download size={15} />} {isApply ? "신청서 생성" : "출석부 생성"}
            </button>
          </div>
        )}
        {prog && (
          <div className="mt-3">
            <div className="flex justify-between text-xs text-slate-500 mb-1"><span>{prog.label}</span><span>{prog.done}/{prog.total}</span></div>
            <div className="h-2 rounded bg-slate-100 overflow-hidden"><div className="h-full bg-emerald-500 transition-all" style={{ width: `${(prog.done / Math.max(prog.total, 1)) * 100}%` }} /></div>
          </div>
        )}
        {err && <p className="mt-3 text-sm text-red-600 break-all">⚠ {err}</p>}
      </Panel>
      {log.length > 0 && (
        <Panel title="생성 결과">
          <pre className="text-xs text-slate-600 whitespace-pre-wrap max-h-72 overflow-auto">{log.join("\n")}</pre>
        </Panel>
      )}
    </div>
  );
}

/* ─────────────────────── 강습 등록·취소 (바이비 직접) — 금액 계산은 PC 자동화(v3)와 동일 */
const BYB_LCATS = {
  GX: ["2단지 요가", "2단지 줌바", "2단지 매트", "2단지 근력", "2단지 방송", "2단지 타바타"],
  농구: ["농구"], 축구: ["축구"], 수영: ["수영"], 아쿠아: ["아쿠아"],
};
const WD = { 일: 0, 월: 1, 화: 2, 수: 3, 목: 4, 금: 5, 토: 6 };
function lessonInfo(name) {
  const inner = (name.split("(").pop() || "");
  const d = inner.match(/([월화수목금토일]+)(\d{1,2})시(?:(\d{1,2})분)?/);
  const days = d ? [...d[1]] : [];
  const hh = d ? +d[2] : 0, mm = d && d[3] ? +d[3] : 0;
  const ball = /축구|농구/.test(name);
  const weekendOnly = days.length > 0 && days.every((x) => x === "토" || x === "일");
  const total = ball || weekendOnly ? 4 : 8;
  const price = ball ? 100000 : weekendOnly ? 30000 : 50000;
  return { days, hh, mm, total, price };
}
// 이번 달 수업일 (체크한 공휴일 제외, 기준 횟수까지만)
function classDays(info, y, m, holi) {
  const out = [];
  const last = new Date(y, m, 0).getDate();
  for (let d = 1; d <= last; d++) {
    if (holi.includes(d)) continue;
    const wd = new Date(y, m - 1, d).getDay();
    if (info.days.some((x) => WD[x] === wd)) out.push(d);
  }
  return out.slice(0, info.total);
}
function lessonCalc({ info, y, m, holi, action, when, regAt, price }) {
  const days = classDays(info, y, m, holi);
  const unit = price / info.total;
  const at = (d) => new Date(y, m - 1, d, info.hh, info.mm);
  if (action === "등록") {
    const missed = days.filter((d) => at(d) <= when).length;       // 이미 지난 수업
    const attendable = Math.max(days.length - missed, 0);
    const applied = Math.round(unit * Math.min(attendable, info.total));
    return { days, unit, count: attendable, amount: Math.min(Math.max(price - applied, 0), price),
      text: `참석 가능 ${attendable}회 × ${Math.round(unit).toLocaleString()}원 = 적용 ${applied.toLocaleString()}원 → 차감` };
  }
  if (price === 150000) return { days, unit, count: 0, amount: price, text: "특강(150,000원반) → 전액 환불" };
  const attended = days.filter((d) => at(d) <= when && (!regAt || at(d) > regAt)).length;  // 이미 들은 수업
  const refund = Math.max(Math.round(price - unit * Math.min(attended, info.total)), 0);
  return { days, unit, count: attended, amount: refund,
    text: `수강 ${attended}회 × ${Math.round(unit).toLocaleString()}원 사용 → 환불` };
}

function LessonPanel({ call, me }) {
  const now = new Date();
  const [cat, setCat] = useState("GX");
  const [ym, setYm] = useState({ y: now.getFullYear(), m: now.getMonth() + 1 });
  const [holi, setHoli] = useState([]);
  const [list, setList] = useState([]);        // [{name, ids:[]}]
  const [course, setCourse] = useState("");
  const [action, setAction] = useState("등록");
  const [members, setMembers] = useState(null);
  const [pick, setPick] = useState(null);      // 취소 대상
  const [dong, setDong] = useState(""); const [ho, setHo] = useState(""); const [name, setName] = useState("");
  const [whenStr, setWhenStr] = useState(() => { const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset()); return d.toISOString().slice(0, 16); });
  const [priceStr, setPriceStr] = useState("");
  const [manual, setManual] = useState("");
  const [busy, setBusy] = useState(""); const [err, setErr] = useState("");
  const [winId, setWinId] = useState("");
  const [q, setQ] = useState("");   // 취소 명단 이름 검색
  const [cq, setCq] = useState("");  // 강좌 검색 (띄어쓰기 무시, 여러 단어 모두 포함)
  const [holiMsg, setHoliMsg] = useState("");
  const holiKey = `lesson_holidays_${ym.y}-${String(ym.m).padStart(2, "0")}`;
  // 월별 공휴일: 한 번 체크하면 저장돼서 다음에도(다른 직원도) 그대로 사용
  React.useEffect(() => {
    let alive = true;
    const parse = (v) => String(v || "").split(",").map(Number).filter((n) => n >= 1 && n <= 31);
    let local = null;
    try { local = localStorage.getItem(holiKey); } catch {}
    setHoli(local != null ? parse(local) : []);
    setHoliMsg(local ? "저장된 공휴일을 불러왔어요" : "");
    // 공유 저장소(Supabase)에 있으면 그걸 우선 사용 — 연결이 안 되면 이 PC 저장값 사용
    supabase.from("settings").select("value").eq("key", holiKey).maybeSingle()
      .then(({ data }) => {
        if (!alive || !data || data.value == null) return;
        setHoli(parse(data.value));
        setHoliMsg("저장된 공휴일을 불러왔어요 (직원 공유)");
      }, () => {});
    return () => { alive = false; };
  }, [holiKey]);
  const toggleHoli = (d) => {
    const next = (holi.includes(d) ? holi.filter((x) => x !== d) : [...holi, d]).sort((a, b) => a - b);
    setHoli(next);
    const label = `${ym.m}월 공휴일 저장됨 (${next.length ? next.join(", ") + "일" : "없음"})`;
    try { localStorage.setItem(holiKey, next.join(",")); } catch {}
    setHoliMsg(label + " · 이 PC");
    supabase.from("settings").upsert({ key: holiKey, value: next.join(",") })
      .then(({ error }) => { if (!error) setHoliMsg(label + " · 직원 공유"); }, () => {});
  };
  const [people, setPeople] = useState(null); const [person, setPerson] = useState(null);
  const [confirm, setConfirm] = useState(false); const [reason, setReason] = useState(""); const [doneMsg, setDoneMsg] = useState("");

  const doCancel = async () => {
    setBusy("바이비에서 취소·환불 처리 중"); setErr(""); setDoneMsg("");
    try {
      await call({ action: "cancel", rid: pick.rid, amount, reason, course });
      setDoneMsg(`${pick.name} 취소 완료 · 환불 ${amount.toLocaleString()}원`);
      logWork({ kind: "강습", action: "취소", target: course, seat: `${ym.y}-${String(ym.m).padStart(2, "0")} 창구`, member: pick.name,
        dong: pick.dong, ho: pick.ho, amount: String(amount), timing: `환불 · 사유: ${reason}`, by: me?.id || "?", at: nowLocal().replace("T", " ") });
      setConfirm(false); setPick(null); setManual("");
      await loadMembers();
    } catch (e) { setErr(e.message); } finally { setBusy(""); }
  };

  const loadCourses = async () => {
    setBusy("강좌 불러오는 중"); setErr(""); setCourse(""); setCq(""); setMembers(null); setPick(null);
    try {
      const { windows } = await call({ action: "windows", keywords: BYB_LCATS[cat], year: ym.y, month: ym.m });
      const map = new Map();
      windows.forEach((w) => { if (!map.has(w.name)) map.set(w.name, []); map.get(w.name).push(w); });
      setList([...map].map(([n, ws]) => ({ name: n, ids: ws.map((w) => w.id), ws })).sort((a, b) => a.name.localeCompare(b.name, "ko")));
    } catch (e) { setErr(e.message); } finally { setBusy(""); }
  };
  React.useEffect(() => { loadCourses(); }, [cat, ym.y, ym.m]);  // eslint-disable-line

  const loadMembers = async (c = course) => {
    const item = list.find((x) => x.name === c);
    if (!item) return;
    setBusy("수강생 불러오는 중"); setErr(""); setPick(null); setConfirm(false); setQ("");
    try {
      const { members: r } = await call({ action: "members", meta: true, year: ym.y, month: ym.m,
        windows: item.ids.map((id) => ({ id, name: c })) });
      setMembers(r[c] || []);
    } catch (e) { setErr(e.message); } finally { setBusy(""); }
  };
  React.useEffect(() => { if (course && action === "취소") loadMembers(); }, [course, action, ym.y, ym.m]);  // eslint-disable-line

  const item = list.find((x) => x.name === course);
  const flat = (t) => String(t).replace(/[\s\-()/·.,]/g, "");
  const cwords = cq.trim().split(/\s+/).map(flat).filter(Boolean);
  const courseHits = list.filter((c) => cwords.every((w) => flat(c.name).includes(w)));
  const off = (ym.y - now.getFullYear()) * 12 + ym.m - (now.getMonth() + 1);
  const appKey = off < 0 ? "prevMonth" : off === 0 ? "thisMonth" : "nextMonth";
  const autoWin = item ? [...item.ws].sort((a, b) => ((b.apps || {})[appKey] || 0) - ((a.apps || {})[appKey] || 0) || (b.visible ? 1 : 0) - (a.visible ? 1 : 0))[0] : null;
  const win = item ? (item.ws.find((w) => w.id === winId) || autoWin) : null;
  const info0 = course ? lessonInfo(course) : null;
  const info = info0 && win && win.price ? { ...info0, price: win.price } : info0;
  const price = Number(priceStr.replace(/[^\d]/g, "")) || (info ? info.price : 0);

  const findPeople = async () => {
    setBusy("회원 찾는 중"); setErr(""); setPeople(null); setPerson(null);
    try { const { people: r } = await call({ action: "household", dong, ho }); setPeople(r); if (r.length === 1) setPerson(r[0]); }
    catch (e) { setErr(e.message); } finally { setBusy(""); }
  };
  const doRegister = async () => {
    setBusy("바이비에 등록 중"); setErr(""); setDoneMsg("");
    try {
      await call({ action: "register", windowId: win.id, rauthId: person.rauthId, uid: person.uid, course,
        month: `${ym.y}-${String(ym.m).padStart(2, "0")}`, deduct: amount, price });
      setDoneMsg(`${person.name} (${dong}동 ${ho}호) ${ym.m}월 등록 완료 · 차감 ${amount.toLocaleString()}원`);
      logWork({ kind: "강습", action: "등록", target: course, seat: `${ym.y}-${String(ym.m).padStart(2, "0")} 창구`, member: person.name,
        dong, ho, amount: String(amount), timing: `차감 · 강습료 ${price.toLocaleString()}원`, by: me?.id || "?", at: nowLocal().replace("T", " ") });
      setConfirm(false); setPerson(null); setPeople(null); setManual("");
    } catch (e) { setErr(e.message); } finally { setBusy(""); }
  };
  const when = new Date(whenStr);
  const regAt = pick && pick.at ? new Date(pick.at) : null;
  const calc = info ? lessonCalc({ info, y: ym.y, m: ym.m, holi, action, when, regAt, price }) : null;
  const amount = manual.trim() ? Number(manual.replace(/[^\d]/g, "")) : calc ? calc.amount : 0;

  const last = new Date(ym.y, ym.m, 0).getDate();
  const first = new Date(ym.y, ym.m - 1, 1).getDay();
  const cls = info ? new Set(classDays(info, ym.y, ym.m, [])) : new Set();
  const months = [-1, 0, 1, 2].map((k) => { const d = new Date(now.getFullYear(), now.getMonth() + k, 1); return { y: d.getFullYear(), m: d.getMonth() + 1 }; });
  const chip = (on) => `px-3 py-1.5 rounded-lg text-sm border ${on ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-200 text-slate-600"}`;

  return (
    <div className="space-y-4">
      <div>
        <p className="text-xs text-slate-500 mb-1">종목</p>
        <div className="flex flex-wrap gap-1.5">
          {Object.keys(BYB_LCATS).map((k) => <button key={k} className={chip(cat === k)} onClick={() => setCat(k)}>{k === "GX" ? "GX (요가·줌바·매트·근력·방송·타바타)" : k}</button>)}
        </div>
      </div>
      <div>
        <p className="text-xs text-slate-500 mb-1">작업할 창구 (월)</p>
        <div className="flex flex-wrap gap-1.5">
          {months.map((o) => <button key={o.y * 100 + o.m} className={chip(ym.y === o.y && ym.m === o.m)} onClick={() => setYm(o)}>{o.y}년 {o.m}월</button>)}
        </div>
      </div>
      <div>
        <p className="text-xs text-slate-500 mb-1">강좌 ({list.length}개)</p>
        {course ? (
          <div className="bg-emerald-50 border border-emerald-200 rounded-lg px-3 py-2 flex items-center justify-between gap-2">
            <span className="text-sm text-emerald-800">✔ <b>{course}</b></span>
            <button onClick={() => { setCourse(""); setCq(""); setConfirm(false); setMembers(null); }} className="text-xs text-emerald-700 underline shrink-0">변경</button>
          </div>
        ) : (
          <>
            <div className="relative">
              <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
              <input className={attInput + " pl-8"} placeholder="강좌 검색 (예: 화목 07 / 김소운 / 초급)" value={cq} onChange={(e) => setCq(e.target.value)} />
            </div>
            <div className="mt-1 border border-slate-200 rounded-lg max-h-60 overflow-y-auto divide-y divide-slate-100">
              {courseHits.length === 0 && <p className="text-xs text-slate-400 p-3">검색 결과가 없어요.</p>}
              {courseHits.map((c) => (
                <button key={c.name} onClick={() => { setCourse(c.name); setManual(""); setPriceStr(""); setWinId(""); setConfirm(false); }}
                  className="block w-full text-left px-3 py-2 text-sm hover:bg-emerald-50">{c.name}</button>
              ))}
            </div>
          </>
        )}
      </div>
      {item && item.ws.length > 1 && (
        <L label="같은 이름 창구가 여러 개예요 — 사용할 창구">
          <select className={attInput} value={win ? win.id : ""} onChange={(e) => setWinId(e.target.value)}>
            {item.ws.map((w, i) => <option key={w.id} value={w.id}>#{i + 1} {w.visible ? "공개" : "비공개"} · {(w.price || 0).toLocaleString()}원 · 지난달 {w.apps?.prevMonth ?? "-"} / 이번달 {w.apps?.thisMonth ?? "-"} / 다음달 {w.apps?.nextMonth ?? "-"}명{w.id === autoWin?.id ? " (자동 선택)" : ""}</option>)}
          </select>
        </L>
      )}
      <div>
        <p className="text-xs text-slate-500 mb-1">공휴일·휴강일 체크 ({ym.m}월) — <span className="text-emerald-700">초록 = 이 강좌 수업일</span></p>
        <div className="grid grid-cols-7 gap-1 max-w-sm text-center text-xs">
          {"일월화수목금토".split("").map((d) => <div key={d} className="text-slate-400">{d}</div>)}
          {Array.from({ length: first }, (_, i) => <div key={"b" + i} />)}
          {Array.from({ length: last }, (_, i) => i + 1).map((d) => {
            const off = holi.includes(d);
            return (
              <button key={d} onClick={() => toggleHoli(d)}
                className={`py-1.5 rounded border ${off ? "bg-rose-500 border-rose-500 text-white line-through" : cls.has(d) ? "bg-emerald-50 border-emerald-300 text-emerald-800 font-semibold" : "border-slate-200 text-slate-500"}`}>{d}</button>
            );
          })}
        </div>
        <p className="text-[11px] text-slate-400 mt-1">쉬는 날을 눌러 빨간색으로 표시하세요. 다시 누르면 해제됩니다. <b>월별로 자동 저장</b>돼서 다음에도 그대로 적용돼요.</p>
        {holiMsg && <p className="text-[11px] text-emerald-700 mt-0.5">✔ {holiMsg}</p>}
      </div>
      <ActionToggle action={action} setAction={(a) => { setAction(a); setPick(null); setManual(""); setConfirm(false); }} />
      {action === "등록" ? (
        <div className="space-y-2">
          <div className="flex gap-2 items-end">
            <L label="동"><input className={attInput} inputMode="numeric" value={dong} onChange={(e) => { setDong(e.target.value); setPeople(null); setPerson(null); }} /></L>
            <L label="호수"><input className={attInput} inputMode="numeric" value={ho} onChange={(e) => { setHo(e.target.value); setPeople(null); setPerson(null); }}
              onKeyDown={(e) => e.key === "Enter" && dong && ho && findPeople()} /></L>
            <button onClick={findPeople} disabled={!dong || !ho || !!busy} className="rounded-lg border border-emerald-600 text-emerald-700 text-sm px-3 py-2 whitespace-nowrap disabled:opacity-40">회원 찾기</button>
          </div>
          {people && (!people.length ? <p className="text-xs text-slate-400">{dong}동 {ho}호에 등록된 회원이 없어요.</p> :
            <div className="border border-slate-200 rounded-lg divide-y divide-slate-100 max-h-48 overflow-y-auto">
              {people.map((pp) => (
                <button key={pp.rauthId} onClick={() => setPerson(pp)} className={`w-full text-left px-3 py-2 text-sm ${person === pp ? "bg-emerald-50 text-emerald-800 font-medium" : "hover:bg-slate-50"}`}>
                  {pp.name} <span className="text-xs text-slate-400">{pp.birth ? pp.birth + "년생" : ""} {pp.phone4 ? "· ****" + pp.phone4 : ""}</span>
                </button>
              ))}
            </div>)}
        </div>
      ) : (
        <div>
          <p className="text-xs text-slate-500 mb-1">취소할 수강생 ({ym.m}월 확정 명단)</p>
          {!course ? <p className="text-xs text-slate-400">강좌를 먼저 고르세요.</p>
            : members == null ? <p className="text-xs text-slate-400">불러오는 중…</p>
            : !members.length ? <p className="text-xs text-slate-400">{ym.m}월 확정 수강생이 없어요.</p>
            : <>
              <div className="relative mb-1.5">
                <Search size={14} className="absolute left-2.5 top-1/2 -translate-y-1/2 text-slate-400" />
                <input className={attInput + " pl-8"} placeholder="이름으로 찾기 (예: 이승)" value={q} onChange={(e) => setQ(e.target.value)} />
              </div>
              <div className="border border-slate-200 rounded-lg max-h-60 overflow-y-auto divide-y divide-slate-100">
                {members.filter((mb) => !q.trim() || mb.name.startsWith(q.trim())).length === 0 && <p className="text-xs text-slate-400 p-3">'{q}'(으)로 시작하는 수강생이 없어요.</p>}
                {members.filter((mb) => !q.trim() || mb.name.startsWith(q.trim())).map((mb, i) => (
                  <button key={i} onClick={() => setPick(mb)}
                    className={`w-full text-left px-3 py-2 text-sm ${pick === mb ? "bg-rose-50 text-rose-800 font-medium" : "hover:bg-slate-50"}`}>
                    {mb.name} <span className="text-xs text-slate-400">{mb.dong}동 {mb.ho}호 · {mb.at ? mb.at.slice(0, 16).replace("T", " ") + " 등록" : ""}</span>
                  </button>
                ))}
              </div>
            </>}
        </div>
      )}
      {info && (
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
          <L label="처리 기준 일시"><input type="datetime-local" className={attInput} value={whenStr} onChange={(e) => setWhenStr(e.target.value)} /></L>
          <L label={`강습료 (기본 ${info.price.toLocaleString()}원)`}><input className={attInput} placeholder={String(info.price)} value={priceStr} onChange={(e) => setPriceStr(e.target.value)} /></L>
          <L label={`수동 ${action === "등록" ? "차감" : "환불"}액 (선택)`}><input className={attInput} placeholder="비우면 자동 계산" value={manual} onChange={(e) => setManual(e.target.value)} /></L>
        </div>
      )}
      {calc && (
        <div className={`rounded-xl border p-3 text-sm ${action === "등록" ? "bg-emerald-50 border-emerald-200" : "bg-rose-50 border-rose-200"}`}>
          <p className="text-xs text-slate-600">{ym.m}월 수업일 {calc.days.length}회 (기준 {info.total}회): {calc.days.map((d) => `${ym.m}/${d}`).join(", ") || "없음"}</p>
          <p className="text-xs text-slate-600 mt-0.5">{calc.text}</p>
          <p className="text-lg font-bold mt-1">{action === "등록" ? "차감" : "환불"} {amount.toLocaleString()}원{manual.trim() ? " (수동)" : ""}</p>
        </div>
      )}
      {action === "취소" ? (
        confirm ? (
          <div className="rounded-xl border-2 border-rose-300 bg-rose-50 p-3 space-y-2">
            <p className="text-sm text-rose-900"><b>{pick.name}</b> ({pick.dong}동 {pick.ho}호)<br />{course}<br />{ym.m}월 수강 취소 · <b>환불 {amount.toLocaleString()}원</b></p>
            <L label="사유 (바이비에 기록)"><input className={attInput} value={reason} onChange={(e) => setReason(e.target.value)} /></L>
            <div className="flex gap-2">
              <button onClick={doCancel} disabled={!!busy} className="flex-1 rounded-lg bg-rose-600 hover:bg-rose-700 text-white text-sm font-medium py-2">네, 바이비에서 취소·환불</button>
              <button onClick={() => setConfirm(false)} className="flex-1 rounded-lg border border-slate-300 text-slate-600 text-sm py-2">돌아가기</button>
            </div>
          </div>
        ) : (
          <button disabled={!pick || !!busy} onClick={() => { setReason(`관리사무소 요청 ${new Date().toLocaleString("ko-KR")}`); setConfirm(true); }}
            className="w-full rounded-lg bg-rose-600 hover:bg-rose-700 disabled:opacity-40 text-white text-sm font-medium py-2.5">
            바이비에 취소·환불 실행
          </button>
        )
      ) : (
        confirm ? (
          <div className="rounded-xl border-2 border-emerald-300 bg-emerald-50 p-3 space-y-2">
            <p className="text-sm text-emerald-900"><b>{person.name}</b> ({dong}동 {ho}호)<br />{course}<br />{ym.y}년 {ym.m}월 등록 · 강습료 {price.toLocaleString()}원 · <b>차감 {amount.toLocaleString()}원</b> (부과 {(price - amount).toLocaleString()}원)</p>
            <div className="flex gap-2">
              <button onClick={doRegister} disabled={!!busy} className="flex-1 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium py-2">네, 바이비에 등록</button>
              <button onClick={() => setConfirm(false)} className="flex-1 rounded-lg border border-slate-300 text-slate-600 text-sm py-2">돌아가기</button>
            </div>
          </div>
        ) : (
          <button disabled={!person || !win || !!busy || amount > price} onClick={() => setConfirm(true)}
            className="w-full rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-sm font-medium py-2.5">
            바이비에 등록 실행
          </button>
        )
      )}
      {doneMsg && <p className="text-sm text-emerald-700">✔ {doneMsg}</p>}
      {busy && <p className="text-xs text-slate-500 flex items-center gap-1"><Loader2 size={13} className="animate-spin" /> {busy}</p>}
      {err && <p className="text-sm text-red-600 break-all">⚠ {err}</p>}
    </div>
  );
}

/* ─────────────────────── 공통 */
function L({ label, children }) { return <label className="block"><span className="text-[11px] font-medium text-slate-500 mb-1 block">{label}</span>{children}</label>; }
function Panel({ title, children }) { return <div className="bg-white border border-byb-line mb-4"><p className="px-4 py-3 border-b border-byb-line text-[14px] font-bold text-slate-800">{title}</p><div className="p-4">{children}</div></div>; }
function Stat({ label, value, accent }) {
  return <div className={`rounded-xl border p-3 ${accent ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-slate-200"}`}>
    <p className={`text-xs ${accent ? "text-emerald-100" : "text-slate-500"}`}>{label}</p><p className="text-base font-bold mt-0.5">{value}</p></div>;
}

/* ═══════════ 수영 반변경 → 신청서 반영 ═══════════
   신청서(xlsx 또는 신청서 생성 zip) + 반변경 명단(양식) → 기존 반에서 빼고 위로 당김, 변경 반 맨 밑에 추가(비고 '반변경') */
const fileB64 = (file) => new Promise((ok, no) => {
  const r = new FileReader();
  r.onload = () => ok(String(r.result).split(",")[1]);
  r.onerror = () => no(new Error(`${file.name} 파일을 읽지 못했어요.`));
  r.readAsDataURL(file);
});

function SwimClassChange({ me }) {
  const [apply, setApply] = useState(null);
  const [list, setList] = useState(null);
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [res, setRes] = useState(null);

  const run = async () => {
    setErr(""); setRes(null); setBusy(true);
    try {
      const r = await fetch("/api/swapclass", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ apply: await fileB64(apply), apply_name: apply.name, list: await fileB64(list) }) });
      const j = await r.json().catch(() => ({ error: `서버 응답 오류 (${r.status})` }));
      if (!r.ok || j.error) throw new Error(j.error || `오류 ${r.status}`);
      const bytes = Uint8Array.from(atob(j.file), (c) => c.charCodeAt(0));
      const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = j.file_name; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      setRes(j);
      const ok = (j.results || []).filter((x) => x.ok).length;
      logWork({ kind: "신청서", action: "반변경", target: apply.name, seat: "수영",
        member: `${ok}명 적용 · ${(j.results || []).length - ok}명 미적용`, by: me?.id || "?", at: nowLocal().replace("T", " ") });
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };

  const pick = (label, f, set, hint) => (
    <label className="flex-1 flex flex-col items-center justify-center gap-1 rounded-xl border-2 border-dashed border-slate-200 py-4 px-2 text-sm text-slate-500 cursor-pointer hover:bg-slate-50 text-center">
      <span className="flex items-center gap-1.5"><Upload size={15} /> {label}</span>
      <span className={`text-xs truncate max-w-full ${f ? "text-emerald-700 font-medium" : "text-slate-400"}`}>{f ? f.name : hint}</span>
      <input type="file" accept={label.startsWith("①") ? ".xlsx,.zip" : ".xlsx"} className="hidden"
        onChange={(e) => { set(e.target.files[0] || null); setRes(null); e.target.value = ""; }} />
    </label>
  );

  const bad = res ? res.results.filter((x) => !x.ok) : [];
  return (
    <Panel title="수영 반변경 → 신청서 반영">
      <p className="text-xs text-slate-500 mb-2 leading-relaxed">
        기존 반에서는 이름을 지우고 아래 명단을 위로 당기고, 변경 반에는 맨 밑에 추가하면서 비고에 <b>반변경</b>을 적어요.<br />
        성함·동·호수가 신청서와 정확히 일치할 때만 옮기고, 애매한 건 바꾸지 않고 아래에 사유를 보여 줘요.
        <a href="/수영_반변경_양식.xlsx" download className="ml-1 text-emerald-700 underline">반변경 양식 받기</a>
      </p>
      <div className="flex gap-2">
        {pick("① 수영 신청서", apply, setApply, "xlsx 또는 신청서 생성 zip")}
        {pick("② 반변경 명단", list, setList, "양식에 적은 xlsx")}
      </div>
      {err && <p className="text-sm text-rose-600 mt-3">{err}</p>}
      <button onClick={run} disabled={busy || !apply || !list}
        className="mt-3 w-full flex items-center justify-center gap-2 rounded-xl bg-emerald-600 text-white font-medium py-3 disabled:opacity-40">
        {busy ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />} 반변경 반영한 신청서 만들기
      </button>
      {res && (
        <div className="mt-3 space-y-3 text-sm">
          <p className="font-medium text-slate-700">{res.log[0]}</p>
          {bad.length > 0 && (
            <div className="rounded-xl border border-amber-200 bg-amber-50 p-3">
              <p className="font-semibold text-amber-700 mb-1">미적용 {bad.length}명 — 양식을 고쳐서 다시 올려 주세요</p>
              {bad.map((x, i) => <p key={i} className="text-amber-700">· {x.row}행 {x.name}({x.dong}-{x.ho}): {x.msg}</p>)}
            </div>
          )}
          <div className="overflow-x-auto rounded-xl border border-slate-200">
            <table className="w-full text-xs">
              <thead className="bg-slate-50 text-slate-500"><tr>
                <th className="px-2 py-1.5 text-left">성함</th><th className="px-2 py-1.5">동-호</th>
                <th className="px-2 py-1.5 text-left">변경 전</th><th className="px-2 py-1.5 text-left">변경 후</th><th className="px-2 py-1.5">결과</th>
              </tr></thead>
              <tbody>{res.results.map((x, i) => (
                <tr key={i} className="border-t border-slate-100">
                  <td className="px-2 py-1.5">{x.name}</td><td className="px-2 py-1.5 text-center">{x.dong}-{x.ho}</td>
                  <td className="px-2 py-1.5">{x.from}</td><td className="px-2 py-1.5">{x.to}</td>
                  <td className={`px-2 py-1.5 text-center font-medium ${x.ok ? "text-emerald-600" : "text-amber-600"}`}>{x.ok ? "적용" : "미적용"}</td>
                </tr>))}</tbody>
            </table>
          </div>
          {res.over.length > 0 && (
            <div className="rounded-xl border border-slate-200 p-3">
              <p className="font-semibold text-slate-700 mb-1">정원(25명) 초과반 {res.over.length}개</p>
              {res.over.map((o, i) => <p key={i} className="text-slate-600">· {o.cls} <b>{o.n}명</b></p>)}
            </div>
          )}
        </div>
      )}
    </Panel>
  );
}

/* ═══════════ 독서실 신청서 · 대기자 생성 ═══════════
   ① 구독현황(당첨자, 좌석번호 있음) → 신청서 시트에 좌석번호 자리에 채움
   ② 추첨_지원자_리스트(당첨·탈락) → '탈락'만 대기순번 순서대로 대기자 시트에 채움 */
const RR_KINDS = ["개인", "남자일반", "여자일반"];
const rrNum = (s) => String(s ?? "").replace(/[^0-9]/g, "");
const rrPhone = (s) => { const d = rrNum(s); return d.length === 11 ? `${d.slice(0, 3)}-${d.slice(3, 7)}-${d.slice(7)}` : String(s ?? ""); };
const rrKindOf = (text, genders = []) => {
  if (/개인/.test(text)) return "개인";
  if (/남자|남성/.test(text)) return "남자일반";
  if (/여자|여성/.test(text)) return "여자일반";
  const g = new Set(genders.filter(Boolean));
  if (g.size === 1) return g.has("남") ? "남자일반" : "여자일반";
  return "";
};

async function rrParse(file) {
  const wb = XLSX.read(await file.arrayBuffer(), { type: "array" });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, defval: "", raw: false });
  const hi = aoa.findIndex((r) => r.includes("회원명"));
  if (hi < 0) throw new Error(`${file.name}: '회원명' 머리글을 못 찾았어요.`);
  const H = aoa[hi]; const col = (n) => H.indexOf(n);
  const body = aoa.slice(hi + 1).filter((r) => r[col("회원명")]);
  const title = String(aoa[0]?.[0] || "");
  const base = (r) => ({ dong: rrNum(r[col("동")]), ho: rrNum(r[col("호")]), name: String(r[col("회원명")]).trim() });
  if (col("시설/기기") >= 0) {   // 구독현황 → 신청서
    const rows = body.map((r) => ({ ...base(r), phone: rrPhone(r[col("연락처")]),
      rk: rrKindOf(`${r[col("그룹")] ?? ""} ${r[col("시설/기기")] ?? ""}`),
      seat: Number(String(r[col("시설/기기")]).split("/").pop().replace(/\D/g, "")) || 0,
      date: String(r[col("구독일시")]).slice(0, 10) }));
    const m = title.match(/(\d{4})-(\d{2})-\d{2}\s*~/);
    let kind = rrKindOf(file.name + title);
    if (!kind && /일반/.test(file.name + title)) kind = "일반(남녀)";
    return { type: "apply", rows, kind, period: m ? { y: +m[1], m: +m[2] } : null };
  }
  if (col("대기순번") >= 0) {     // 추첨 지원자 리스트 → 대기자
    const rows = body.filter((r) => String(r[col("상태")]).trim() === "탈락" && rrNum(r[col("대기순번")]))
      .map((r) => ({ ...base(r), phone: rrPhone(r[col("전화 번호")]), order: Number(rrNum(r[col("대기순번")])),
        date: String(r[col("추첨생성일시")]).slice(0, 10) }))
      .sort((a, b) => a.order - b.order);
    return { type: "wait", rows, kind: rrKindOf(file.name, body.map((r) => String(r[col("성별")]).trim())) };
  }
  throw new Error(`${file.name}: 구독현황/추첨 지원자 리스트 형식이 아니에요.`);
}

function ReadingRoom() {
  const now = new Date();
  let defStart = now.getMonth() + 2, defYear = now.getFullYear();   // 다음 달부터 시작하는 홀수 월 묶음
  if (defStart % 2 === 0) defStart += 1;
  if (defStart > 12) { defStart -= 12; defYear += 1; }
  const [year, setYear] = useState(defYear);
  const [start, setStart] = useState(defStart);
  const [files, setFiles] = useState([]);   // {name, type, kind, rows}
  const [busy, setBusy] = useState(false);
  const [err, setErr] = useState("");
  const [log, setLog] = useState([]);

  const add = async (list) => {
    setErr("");
    for (const f of list) {
      try {
        const p = await rrParse(f);
        if (p.period) { setYear(p.period.y); setStart(p.period.m); }
        setFiles((xs) => [...xs, { name: f.name, ...p }]);
      } catch (e) { setErr(e.message); }
    }
  };
  const setKind = (i, kind) => setFiles((xs) => xs.map((x, j) => (j === i ? { ...x, kind } : x)));

  const build = async () => {
    setErr(""); setLog([]);
    if (files.some((f) => !f.kind)) return setErr("구분(개인/남자일반/여자일반)을 모두 골라 주세요.");
    const sheets = {};
    for (const f of files) {
      for (const r of f.rows) {
        const k = f.kind === "일반(남녀)" ? r.rk : f.kind;
        if (!RR_KINDS.includes(k)) return setErr(`${f.name}: ${r.name}(${r.dong}동 ${r.ho}호) 님이 남자/여자 어느 독서실인지 파일에서 구분이 안 돼요.`);
        sheets[k] = sheets[k] || {};
        (sheets[k][f.type] = sheets[k][f.type] || []).push(r);
      }
    }
    setBusy(true);
    try {
      const r = await fetch("/api/reading", { method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ year, start, sheets }) });
      const j = await r.json().catch(() => ({ error: `서버 응답 오류 (${r.status})` }));
      if (!r.ok || j.error) throw new Error(j.error || `오류 ${r.status}`);
      const bytes = Uint8Array.from(atob(j.file), (c) => c.charCodeAt(0));
      const url = URL.createObjectURL(new Blob([bytes], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      const a = document.createElement("a");
      a.href = url; a.download = `2단지 독서실 ${start}~${start + 1}월 신청서·대기자.xlsx`; a.click();
      setTimeout(() => URL.revokeObjectURL(url), 5000);
      setLog(j.log || []);
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };

  const inp = "rounded-lg border border-slate-200 px-2.5 py-2 text-sm";
  return (
    <div>
      <Panel title="① 사용 기간">
        <div className="flex items-center gap-2 text-sm">
          <select value={year} onChange={(e) => setYear(+e.target.value)} className={inp}>
            {[now.getFullYear(), now.getFullYear() + 1].map((y) => <option key={y} value={y}>{y}년</option>)}
          </select>
          <select value={start} onChange={(e) => setStart(+e.target.value)} className={inp}>
            {[1, 3, 5, 7, 9, 11].map((m) => <option key={m} value={m}>{m}~{m + 1}월</option>)}
          </select>
          <span className="text-xs text-slate-400">구독현황 파일을 올리면 자동으로 맞춰져요</span>
        </div>
      </Panel>
      <Panel title="② 바이비 엑셀 올리기">
        <p className="text-xs text-slate-500 mb-2 leading-relaxed">
          · <b>구독현황</b>(당첨자·좌석번호) → 신청서에 좌석번호 자리대로<br />
          · <b>추첨 지원자 리스트</b>(당첨·탈락) → 탈락자만 대기순번 순서대로 대기자에<br />
          개인/남자일반/여자일반 파일을 여러 개 한 번에 올려도 돼요.
        </p>
        <label className="flex items-center justify-center gap-2 rounded-xl border-2 border-dashed border-slate-200 py-5 text-sm text-slate-500 cursor-pointer hover:bg-slate-50">
          <Upload size={16} /> 엑셀 파일 선택
          <input type="file" accept=".xlsx,.xls" multiple className="hidden"
            onChange={(e) => { add([...e.target.files]); e.target.value = ""; }} />
        </label>
        {files.length > 0 && (
          <div className="mt-3 space-y-2">
            {files.map((f, i) => (
              <div key={i} className="flex items-center gap-2 rounded-lg bg-slate-50 px-3 py-2 text-sm">
                <span className={`text-[11px] font-semibold rounded px-1.5 py-0.5 ${f.type === "apply" ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"}`}>
                  {f.type === "apply" ? "신청서" : "대기자"}
                </span>
                <span className="flex-1 truncate">{f.name}</span>
                <span className="text-xs text-slate-500">{f.rows.length}명</span>
                <select value={f.kind} onChange={(e) => setKind(i, e.target.value)} className={`${inp} py-1 ${f.kind ? "" : "border-rose-300"}`}>
                  <option value="">구분 선택</option>
                  {[...RR_KINDS, ...(f.type === "apply" ? ["일반(남녀)"] : [])].map((k) => <option key={k} value={k}>{k === "일반(남녀)" ? "일반(남녀 자동 분리)" : k}</option>)}
                </select>
                <button onClick={() => setFiles((xs) => xs.filter((_, j) => j !== i))} className="text-slate-400 hover:text-rose-500"><X size={15} /></button>
              </div>
            ))}
          </div>
        )}
      </Panel>
      {err && <p className="text-sm text-rose-600 mb-3">{err}</p>}
      <button onClick={build} disabled={busy || !files.length}
        className="w-full flex items-center justify-center gap-2 rounded-xl bg-emerald-600 text-white font-medium py-3 disabled:opacity-40">
        {busy ? <Loader2 size={16} className="animate-spin" /> : <Download size={16} />} 신청서·대기자 엑셀 만들기
      </button>
      {log.length > 0 && (
        <div className="mt-3 rounded-xl bg-white border border-slate-200 p-3 text-sm space-y-1">
          {log.map((l, i) => <p key={i} className={l.startsWith("⚠") ? "text-amber-600" : "text-slate-600"}>{l}</p>)}
        </div>
      )}
    </div>
  );
}

/* ═══════════ 월간 이용 통계 ═══════════
   강좌: 수강기간이 해당 월인 확정 수강생 (강좌명이 '플그'로 시작하면 제외)
   시설: 이용기간이 해당 월인 확정 예약 (체육관·스크린테니스·스크린골프·골프백보관함·실내골프장·탁구장·게스트·악기연습) */
function StatsPanel({ call }) {
  const prev = new Date(new Date().getFullYear(), new Date().getMonth() - 1, 1);
  const [year, setYear] = useState(prev.getFullYear());
  const [month, setMonth] = useState(prev.getMonth() + 1);
  const [busy, setBusy] = useState("");
  const [err, setErr] = useState("");
  const [res, setRes] = useState({});   // kind → [{title, n, people}]

  // 바이비가 '응답시간 초과'를 내면 한 번에 받는 건수를 줄여서(50→10) 같은 위치부터 이어서 조회
  const scan = async (action, onPart, label) => {
    let offset = 0, size = 100, total = null, fails = 0;   // 한 번에 100건씩 (요청 횟수 절반)
    for (let i = 0; i < 2000; i++) {
      let j;
      try { j = await call({ action, year, month, page: offset / size, size }); fails = 0; }
      catch (e) {
        if (!/시간 초과/.test(e.message)) throw e;
        if (size > 10) { size = 10; continue; }
        if (++fails >= 3) throw e;
        await new Promise((r) => setTimeout(r, 3000)); continue;
      }
      if (j.total != null) total = j.total;
      onPart(j);
      offset = j.next * size;
      setRes((r) => ({ ...r, [label]: { loading: `${Math.min(offset, total ?? offset).toLocaleString()}${total ? ` / ${total.toLocaleString()}` : ""}건 확인 중…`, list: [] } }));
      if (j.done) break;
    }
  };

  const run = async (kind) => {
    setErr(""); setBusy(kind);
    const acc = {}; const seen = new Set(); let total = 0;
    // 시설은 건수가 많아 하루씩 나눠 조회(바이비 응답시간 초과 방지), 같은 예약은 한 번만 셈
    const last = new Date(year, month, 0).getDate();
    const days = kind === "facility" ? Array.from({ length: last }, (_, i) => i + 1) : [null];
    try {
      if (kind === "klass") {   // 강좌: 바이비 [수강생 현황]과 같은 조회 → 안 되면 창구별 조회
        try {
          const acc2 = {};
          await scan("stats_v3", (j) => { for (const [t, n] of Object.entries(j.groups || {})) acc2[t] = (acc2[t] || 0) + n; }, "klass");
          const list = Object.entries(acc2).map(([title, n]) => ({ title, chan: "", n, people: n }))
            .sort((x, y) => x.title.localeCompare(y.title, "ko"));
          if (list.length) { setRes((r) => ({ ...r, klass: { list } })); return; }
        } catch (e) { /* 바이비 [수강생 현황] 조회 방식이 바뀌어 막히면 → 아래 창구별 조회로 계속 */ }
        const { windows = [] } = await call({ action: "stats_windows" });
        const list = [];
        for (let i = 0; i < windows.length; i += 15) {
          setRes((r) => ({ ...r, klass: { loading: `강좌 ${i}/${windows.length}개 확인 중…`, list: [] } }));
          const part = windows.slice(i, i + 15);
          const { counts = {} } = await call({ action: "stats_count", year, month, windows: part });
          part.forEach((w) => counts[w.id] && list.push({ title: w.name, chan: "", n: counts[w.id], people: counts[w.id] }));
        }
        list.sort((x, y) => x.title.localeCompare(y.title, "ko"));
        setRes((r) => ({ ...r, klass: { list } }));
        return;
      }
      try {   // 시설: 바이비 [예약자 현황]과 같은 조회 → 안 되면 하루씩 조회
        let got = false;
        await scan("stats_fac_v3", (j) => {
          for (const [t, g] of Object.entries(j.groups || {})) {
            const a = (acc[t] = acc[t] || { n: 0, uids: new Set(), chan: g.chan });
            a.n += g.n; g.uids.forEach((u) => a.uids.add(u)); got = true;
          }
        }, "facility");
        if (got) days.length = 0;
      } catch (e) { throw new Error(`시설 조회 실패: ${e.message}`); }
      for (const day of days) {
        let page = 0;
        for (let i = 0; i < 200; i++) {
          const j = await call({ action: "stats", kind, year, month, day, page, pages: 5 });
          for (const [t, g] of Object.entries(j.groups || {})) {
            const a = (acc[t] = acc[t] || { n: 0, uids: new Set(), chan: g.chan });
            const rids = g.rids || [];
            if (rids.length) { for (const r of rids) if (!seen.has(r)) { seen.add(r); a.n += 1; } }
            else a.n += g.n;
            g.uids.forEach((u) => a.uids.add(u));
          }
          total += j.rows || 0; page = j.next;
          setRes((r) => ({ ...r, [kind]: { loading: day ? `${month}월 ${day}일 확인 중… (${last}일까지)` : `${page * 100}건 확인 중…`, list: [] } }));
          if (j.done) break;
        }
      }
      const list = Object.entries(acc).map(([title, a]) => ({ title, chan: a.chan, n: a.n, people: a.uids.size }))
        .sort((x, y) => x.title.localeCompare(y.title, "ko"));
      setRes((r) => ({ ...r, [kind]: { list, total } }));
    } catch (e) { setErr(e.message); setRes((r) => ({ ...r, [kind]: null })); }
    finally { setBusy(""); }
  };

  const excel = () => {
    const wb = XLSX.utils.book_new();
    const add = (kind, name, head) => {
      const d = res[kind]?.list; if (!d?.length) return;
      const aoa = [head, ...d.map((x) => (kind === "klass" ? [x.title, x.people] : [x.title, x.chan, x.n, x.people]))];
      aoa.push(kind === "klass" ? ["합계", d.reduce((s, x) => s + x.people, 0)]
        : ["합계", "", d.reduce((s, x) => s + x.n, 0), d.reduce((s, x) => s + x.people, 0)]);
      const ws = XLSX.utils.aoa_to_sheet(aoa); ws["!cols"] = [{ wch: 48 }, { wch: 18 }, { wch: 10 }, { wch: 10 }];
      XLSX.utils.book_append_sheet(wb, ws, name);
    };
    add("klass", "강좌별 수강생", ["강좌명", "수강생 수"]);
    add("facility", "시설별 이용자", ["예약상품명", "채널명", "이용 건수", "이용자 수(중복 제외)"]);
    XLSX.writeFile(wb, `${year}년 ${month}월 이용 통계.xlsx`);
  };

  const box = "rounded-lg border border-slate-200 px-2.5 py-2 text-sm";
  const Table = ({ kind }) => {
    const r = res[kind]; if (!r) return null;
    if (r.loading) return <p className="text-sm text-slate-500 flex items-center gap-1"><Loader2 size={14} className="animate-spin" /> {r.loading}</p>;
    const sum = (k) => r.list.reduce((s, x) => s + x[k], 0);
    return (
      <div className="overflow-x-auto"><table className="w-full text-sm">
        <thead><tr className="text-left text-slate-500 border-b">
          <th className="py-1.5">{kind === "klass" ? "강좌명" : "예약상품명"}</th>
          {kind === "facility" && <th className="text-right">이용 건수</th>}
          <th className="text-right">{kind === "klass" ? "수강생" : "이용자(중복제외)"}</th></tr></thead>
        <tbody>
          {r.list.map((x) => <tr key={x.title} className="border-b border-slate-100"><td className="py-1">{x.title}</td>
            {kind === "facility" && <td className="text-right">{x.n.toLocaleString()}</td>}
            <td className="text-right">{x.people.toLocaleString()}</td></tr>)}
          <tr className="font-semibold"><td className="py-1.5">합계 ({r.list.length}개)</td>
            {kind === "facility" && <td className="text-right">{sum("n").toLocaleString()}</td>}
            <td className="text-right">{sum("people").toLocaleString()}</td></tr>
        </tbody></table></div>
    );
  };

  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center gap-2">
        <input type="number" className={`${box} w-24`} value={year} onChange={(e) => setYear(+e.target.value)} />
        <select className={box} value={month} onChange={(e) => setMonth(+e.target.value)}>
          {Array.from({ length: 12 }, (_, i) => <option key={i} value={i + 1}>{i + 1}월</option>)}
        </select>
        <button disabled={!!busy} onClick={() => run("klass")} className="rounded-lg bg-emerald-600 text-white text-sm px-3 py-2 disabled:opacity-40">
          {busy === "klass" ? "조회 중…" : "강좌별 수강생 조회"}</button>
        <button disabled={!!busy} onClick={() => run("facility")} className="rounded-lg bg-sky-600 text-white text-sm px-3 py-2 disabled:opacity-40">
          {busy === "facility" ? "조회 중…" : "시설별 이용자 조회"}</button>
        <button disabled={!!busy || !(res.klass?.list?.length || res.facility?.list?.length)} onClick={excel}
          className="rounded-lg border border-slate-200 text-sm px-3 py-2 flex items-center gap-1 disabled:opacity-40"><Download size={14} /> 엑셀</button>
      </div>
      <p className="text-xs text-slate-400">강좌: 수강기간 기준 확정 수강생, '플그'로 시작하는 강좌 제외 · 시설: 이용기간 기준 확정 예약(취소 제외), 체육관·스크린테니스·스크린골프·골프백보관함·실내골프장·탁구장·게스트·악기연습</p>
      {err && <p className="text-sm text-rose-600">{err}</p>}
      {res.klass && <div><p className="text-sm font-semibold mb-1">강좌별 수강생</p><Table kind="klass" /></div>}
      {res.facility && <div><p className="text-sm font-semibold mb-1">시설별 이용자</p><Table kind="facility" /></div>}
    </div>
  );
}

/* ═══════════ 사우나 핸드키 고장 취합 ═══════════
   바이비 2단지 남/여 사우나 1·2 의 락커(핸드키)별 메모를 읽어 고장건으로 분류·누적(sauna_key_issues) */
const KEY_KINDS = [
  ["분실", /분실|잃어|없어|미반납/],
  ["끈·고리 파손", /끈|고리|스프링|파손|끊|부러|깨/],
  ["고장", /고장|안\s*열|인식|작동|불량|안\s*됨|배터리|잠김|잠겨|오류|에러|먹통/],
];
const KIND_NAMES = [...KEY_KINDS.map((k) => k[0]), "기타"];
const pad2 = (n) => String(n).padStart(2, "0");
const lsGet = (k) => { try { return localStorage.getItem(k) || ""; } catch { return ""; } };

function parseKeyMemo(line, sauna, number) {
  const t = line.trim();
  const [nowY, nowM] = kstDay().split("-").map(Number);
  const dm = t.match(/^(\d{1,2})\s*[./월]\s*(\d{1,2})일?/);
  const hm = t.match(/(\d{3,4})\s*[-동]\s*(\d{3,4})호?/);
  let name = hm ? (t.slice(hm.index + hm[0].length).match(/^\s*([가-힣]{2,4})(?=\s|$)/) || [])[1] || "" : "";
  if (KEY_KINDS.some(([, re]) => re.test(name)) || name === "키") name = "";
  let date = "", ym = "";
  if (dm) {
    const m = +dm[1], d = +dm[2], y = m > nowM ? nowY - 1 : nowY;   // 이번 달보다 큰 월은 작년
    if (m >= 1 && m <= 12) { ym = `${y}-${pad2(m)}`; date = `${ym}-${pad2(d)}`; }
  }
  const kind = (KEY_KINDS.find(([, re]) => re.test(t)) || ["기타"])[0];
  return { id: `${sauna}|${number}|${t}`, sauna, number, line: t, date, ym, dongho: hm ? `${hm[1]}-${hm[2]}` : "", name, kind };
}
const saunaShort = (n) => String(n).replace(/^\s*\d단지\s*/, "").replace(/\s*사우나\s*/, " ").trim();

function BybLoginBox({ onAuth, onClose }) {
  const [f, setF] = useState({ step: "login", user: lsGet("byb_user"), pw: "", code: "", pending: null });
  const [err, setErr] = useState(""); const [busy, setBusy] = useState(false);
  const go = async () => {
    setBusy(true); setErr("");
    try {
      if (f.step === "login") {
        const j = await attCall({ action: "login", user: f.user, password: f.pw });
        try { localStorage.setItem("byb_user", f.user); } catch {}
        if (j.auth) { BYB_AUTH = j.auth; onAuth(); } else setF({ ...f, pw: "", step: "otp", pending: j.pending });
      } else {
        const j = await attCall({ action: "verify", pending: f.pending, code: f.code });
        BYB_AUTH = j.auth; onAuth();
      }
    } catch (e) { setErr(e.message); } finally { setBusy(false); }
  };
  const key = (e) => e.key === "Enter" && go();
  return (
    <div className="bg-white border border-byb-line p-4 space-y-2">
      <div className="flex items-center justify-between">
        <p className="text-sm font-bold">바이비 로그인</p>
        {onClose && <button onClick={onClose} className="text-slate-400"><X size={16} /></button>}
      </div>
      <div className="flex flex-wrap gap-2">
        {f.step === "login" ? <>
          <input value={f.user} onChange={(e) => setF({ ...f, user: e.target.value })} placeholder="바이비 아이디" className={`${attInput} flex-1 min-w-[10rem]`} />
          <input type="password" value={f.pw} onChange={(e) => setF({ ...f, pw: e.target.value })} onKeyDown={key} placeholder="비밀번호" className={`${attInput} flex-1 min-w-[10rem]`} />
        </> : <input value={f.code} onChange={(e) => setF({ ...f, code: e.target.value })} onKeyDown={key} placeholder="휴대폰 인증번호" autoFocus className={`${attInput} flex-1`} />}
        <button onClick={go} disabled={busy} className="rounded-lg bg-emerald-600 disabled:opacity-50 text-white text-xs font-medium px-4">
          {busy ? <Loader2 size={13} className="animate-spin" /> : f.step === "login" ? "로그인" : "확인"}</button>
      </div>
      {err && <p className="text-xs text-rose-600">{err}</p>}
    </div>
  );
}

function SaunaKeys() {
  const [db, setDb] = useState([]);            // 누적 기록
  const [dbErr, setDbErr] = useState("");
  const [live, setLive] = useState(null);      // 이번에 가져온 {rows, unavailable, lockers}
  const [login, setLogin] = useState(false);
  const [prog, setProg] = useState(null);
  const [err, setErr] = useState("");
  const [fm, setFm] = useState("");            // 월 필터
  const [fs, setFs] = useState("");            // 사우나 필터
  const [fk, setFk] = useState("");            // 증상 필터
  const [showGone, setShowGone] = useState(true);

  const loadDb = React.useCallback(async () => {
    const { data, error } = await supabase.from("sauna_key_issues").select("*");
    if (error) { setDbErr(error.message); return null; }
    setDbErr(""); setDb(data || []); return data || [];
  }, []);
  React.useEffect(() => { loadDb(); }, [loadDb]);

  const fetchByb = async () => {
    if (!BYB_AUTH) { setLogin(true); return; }
    setErr(""); setProg({ label: "사우나·핸드키 목록 불러오는 중", done: 0, total: 0 });
    let cur = BYB_AUTH;
    const call = async (body) => { const j = await attCall({ ...body, auth: cur }); if (j.auth) { cur = j.auth; BYB_AUTH = j.auth; } return j; };
    try {
      const { saunas = [] } = await call({ action: "sauna_list" });
      if (!saunas.length) throw new Error("바이비에서 2단지 사우나를 찾지 못했습니다.");
      const lockers = saunas.flatMap((s) => s.lockers.map((l) => ({ ...l, s: s.id, sauna: saunaShort(s.name) })));
      const parts = []; for (let i = 0; i < lockers.length; i += 80) parts.push(lockers.slice(i, i + 80));
      const memos = {}; let done = 0, idx = 0;
      setProg({ label: "핸드키 메모 읽는 중", done: 0, total: lockers.length });
      const worker = async () => {
        while (idx < parts.length) {
          const p = parts[idx++];
          const j = await call({ action: "sauna_memos", items: p.map((l) => ({ s: l.s, l: l.id })) });
          Object.assign(memos, j.memos || {});
          done += p.length; setProg({ label: "핸드키 메모 읽는 중", done, total: lockers.length });
        }
      };
      await Promise.all([worker(), worker(), worker()]);
      const rows = [];
      for (const l of lockers) {
        const memo = (memos[l.id] || "").trim();
        if (!memo) continue;
        memo.split(/\n|(?<=\S)\s+(?=\d{1,2}\/\d{1,2}\s)/).filter((x) => x.trim())
          .forEach((ln) => rows.push({ ...parseKeyMemo(ln, l.sauna, l.number), status: l.status, reason: l.reason }));
      }
      const uniq = [...new Map(rows.map((r) => [r.id, r])).values()];
      const unavailable = lockers.filter((l) => l.status === "UNAVAILABLE")
        .map((l) => ({ sauna: l.sauna, number: l.number, reason: l.reason }));
      setLive({ rows: uniq, unavailable, lockers: lockers.length, at: stamp() });

      // 누적 저장: 새 건은 추가, 이미 있는 건은 확인시각 갱신, 메모에서 사라진 건은 '삭제됨'(처리된 것으로 봄)
      if (!dbErr) {
        const prev = new Map(db.map((r) => [r.id, r])); const at = stamp();
        const up = uniq.map(({ status, reason, ...r }) => ({ ...r, first_seen: prev.get(r.id)?.first_seen || at, last_seen: at, gone: false }));
        for (let i = 0; i < up.length; i += 200) {
          const { error } = await supabase.from("sauna_key_issues").upsert(up.slice(i, i + 200));
          if (error) throw new Error("저장 실패: " + error.message);
        }
        const ids = new Set(uniq.map((r) => r.id));
        const gone = db.filter((r) => !r.gone && !ids.has(r.id)).map((r) => r.id);
        for (let i = 0; i < gone.length; i += 100)
          await supabase.from("sauna_key_issues").update({ gone: true, last_seen: at }).in("id", gone.slice(i, i + 100));
        await loadDb();
      }
    } catch (e) {
      setErr(e.message);
      if (/만료|로그인 정보/.test(e.message)) BYB_AUTH = null;
    } finally { setProg(null); }
  };

  const noTable = /sauna_key_issues|does not exist|schema cache/i.test(dbErr);
  const all = (noTable || dbErr) ? (live?.rows || []) : db;
  const months = [...new Set(all.map((r) => r.ym).filter(Boolean))].sort().reverse();
  const saunas = [...new Set(all.map((r) => r.sauna))].sort();
  const rows = all
    .filter((r) => (!fm || r.ym === fm) && (!fs || r.sauna === fs) && (!fk || r.kind === fk) && (showGone || !r.gone))
    .sort((a, b) => (b.date || "").localeCompare(a.date || "") || a.sauna.localeCompare(b.sauna) || a.number - b.number);
  const kindCount = Object.fromEntries(KIND_NAMES.map((k) => [k, rows.filter((r) => r.kind === k).length]));
  const repeat = new Map(); all.forEach((r) => { const k = `${r.sauna}|${r.number}`; repeat.set(k, (repeat.get(k) || 0) + 1); });

  // 월별 × 증상 (필터: 사우나·메모삭제 포함 여부만 적용, 최근 12개월)
  const chartBase = all.filter((r) => r.ym && (!fs || r.sauna === fs) && (showGone || !r.gone));
  const chart = [...new Set(chartBase.map((r) => r.ym))].sort().slice(-12).map((ym) => {
    const pt = { m: `${ym.slice(2, 4)}.${ym.slice(5)}` };
    KIND_NAMES.forEach((k) => { pt[k] = chartBase.filter((r) => r.ym === ym && r.kind === k).length; });
    return pt;
  });
  const thisYm = kstDay().slice(0, 7);

  const [xbusy, setXbusy] = useState(false);
  const exportXlsx = async () => {
    setXbusy(true); setErr("");
    try {
      const period = fm ? `${fm.slice(0, 4)}년 ${Number(fm.slice(5))}월` : "";
      const filter = [fs || "전체 사우나", fk || "전체 증상", showGone ? "처리 완료 포함" : "미처리만"].join(" · ");
      const j = await attCall({ action: "sauna_report", period, filter,
        rows: rows.map(({ date, ym, sauna, number, dongho, name, kind, line, gone }) => ({ date, ym, sauna, number, dongho, name, kind, line, gone: !!gone })) });
      const bin = Uint8Array.from(atob(j.xlsx), (c) => c.charCodeAt(0));
      const a = document.createElement("a");
      a.href = URL.createObjectURL(new Blob([bin], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
      a.download = `사우나_핸드키_고장_보고_${fm || "전체"}_${kstDay()}.xlsx`;
      a.click(); setTimeout(() => URL.revokeObjectURL(a.href), 2000);
    } catch (e) { setErr("엑셀 생성 실패: " + e.message); } finally { setXbusy(false); }
  };

  const sel = "rounded-lg border border-slate-200 px-2.5 py-1.5 text-xs bg-white";
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <h2 className="text-lg font-bold flex items-center gap-1.5"><KeyRound size={18} className="text-emerald-600" /> 사우나 핸드키 고장 취합</h2>
          <p className="text-xs text-slate-500">바이비 2단지 남자·여자 사우나 1·2 핸드키 메모 기준{live ? ` · 마지막 조회 ${live.at} (키 ${live.lockers}개)` : ""}</p>
        </div>
        <button onClick={fetchByb} disabled={!!prog}
          className="flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white text-xs font-medium px-3 py-2">
          {prog ? <><Loader2 size={13} className="animate-spin" /> {prog.label}{prog.total ? ` ${prog.done}/${prog.total}` : "…"}</>
            : <><RefreshCw size={13} /> 바이비에서 가져오기</>}
        </button>
      </div>
      {login && <BybLoginBox onClose={() => setLogin(false)} onAuth={() => { setLogin(false); fetchByb(); }} />}
      {err && <p className="text-xs text-rose-600 bg-rose-50 border border-rose-200 px-3 py-2">{err}</p>}
      {noTable && (
        <p className="text-xs text-slate-600 bg-amber-50 border border-amber-200 px-3 py-2">
          누적 저장용 표(<b>sauna_key_issues</b>)가 아직 없어 이번 조회 결과만 보여줍니다. Supabase SQL 을 한 번 실행하면 메모가 지워져도 기록이 남습니다.
        </p>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        {[
          ["이번 달 고장건", `${all.filter((r) => r.ym === thisYm).length}건`, `${Number(thisYm.slice(5))}월 메모 기준`],
          ["현재 메모 있는 건", `${all.filter((r) => !r.gone).length}건`, "아직 처리 안 된 것으로 봄"],
          ["메모 삭제된 건", `${all.filter((r) => r.gone).length}건`, "처리 완료 추정"],
          ["바이비 사용불가 키", live ? `${live.unavailable.length}개` : "-", live ? "점검 등으로 막아둔 키" : "가져오기 후 표시"],
        ].map(([l, v, f]) => (
          <div key={l} className="bg-white border border-byb-line px-3 py-2.5">
            <p className="text-[11px] text-slate-500">{l}</p><p className="text-xl font-bold text-slate-900 mt-0.5">{v}</p><p className="text-[11px] text-slate-400 mt-0.5">{f}</p>
          </div>
        ))}
      </div>

      {chart.length > 0 && (
        <div className="bg-white border border-byb-line p-4">
          <div className="flex flex-wrap items-center justify-between gap-2 mb-2">
            <p className="text-[13px] font-bold text-slate-800">월별·증상별 고장건 {fs && <span className="font-normal text-slate-400">· {fs}</span>}</p>
            <div className="flex flex-wrap gap-x-3 gap-y-1">
              {KIND_NAMES.map((k, i) => (
                <span key={k} className="flex items-center gap-1 text-[11px] text-slate-600">
                  <span className="inline-block w-2.5 h-2.5 rounded-sm" style={{ background: SERIES[i] }} />{k}
                </span>
              ))}
            </div>
          </div>
          <div className="h-52">
            <ResponsiveContainer width="100%" height="100%">
              <BarChart data={chart} margin={{ top: 8, right: 8, left: 0, bottom: 0 }} barCategoryGap="30%">
                <CartesianGrid vertical={false} stroke="#eef0f3" />
                <XAxis dataKey="m" tick={{ fontSize: 11, fill: "#888" }} tickLine={false} axisLine={{ stroke: "#e3e6ea" }} />
                <YAxis tick={{ fontSize: 11, fill: "#888" }} tickLine={false} axisLine={false} allowDecimals={false} width={30} />
                <Tooltip content={<TrendTip />} cursor={{ fill: "rgba(0,0,0,.03)" }} />
                {KIND_NAMES.map((k, i) => (
                  <Bar key={k} dataKey={k} name={k} stackId="a" fill={SERIES[i]} stroke="#fff" strokeWidth={1}
                    radius={i === KIND_NAMES.length - 1 ? [3, 3, 0, 0] : 0} isAnimationActive={false} />
                ))}
              </BarChart>
            </ResponsiveContainer>
          </div>
        </div>
      )}

      <div className="bg-white border border-byb-line">
        <div className="px-3 py-2.5 border-b border-byb-line flex flex-wrap items-center gap-2">
          <select value={fm} onChange={(e) => setFm(e.target.value)} className={sel}>
            <option value="">전체 월</option>{months.map((m) => <option key={m} value={m}>{m.slice(0, 4)}년 {Number(m.slice(5))}월</option>)}
          </select>
          <select value={fs} onChange={(e) => setFs(e.target.value)} className={sel}>
            <option value="">전체 사우나</option>{saunas.map((x) => <option key={x}>{x}</option>)}
          </select>
          <div className="flex flex-wrap gap-1">
            {["", ...KIND_NAMES].map((k) => (
              <button key={k || "all"} onClick={() => setFk(k)}
                className={`text-[11px] px-2 py-1 border ${fk === k ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-byb-line text-slate-600"}`}>
                {k || "전체"}{k ? ` ${kindCount[k]}` : ` ${rows.length}`}
              </button>
            ))}
          </div>
          <label className="flex items-center gap-1 text-[11px] text-slate-500"><input type="checkbox" checked={showGone} onChange={(e) => setShowGone(e.target.checked)} /> 메모 삭제된 건 포함</label>
          <button onClick={exportXlsx} disabled={!rows.length || xbusy}
            className="ml-auto flex items-center gap-1 rounded-lg bg-slate-800 disabled:opacity-40 text-white text-xs px-3 py-1.5">
            {xbusy ? <Loader2 size={12} className="animate-spin" /> : <Download size={12} />} 보고서 엑셀</button>
        </div>
        {!all.length ? (
          <p className="text-sm text-slate-400 py-10 text-center">[바이비에서 가져오기] 를 누르면 핸드키 메모를 읽어옵니다.</p>
        ) : (
          <div className="overflow-x-auto max-h-[34rem]">
            <table className="w-full text-[12px]">
              <thead className="bg-byb-bg text-slate-500 sticky top-0">
                <tr>{["접수일", "사우나", "키", "동-호", "이름", "증상", "메모 원문", "상태"].map((h) => <th key={h} className="text-left font-medium px-3 py-2 whitespace-nowrap">{h}</th>)}</tr>
              </thead>
              <tbody>
                {rows.map((r) => {
                  const rep = repeat.get(`${r.sauna}|${r.number}`) || 1;
                  return (
                    <tr key={r.id} className={`border-t border-byb-line ${r.gone ? "text-slate-400" : "text-slate-700"}`}>
                      <td className="px-3 py-2 whitespace-nowrap">{r.date ? r.date.slice(5).replace("-", "/") : "-"}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{r.sauna}</td>
                      <td className="px-3 py-2 whitespace-nowrap font-semibold">{r.number}번{rep > 1 && <span className="ml-1 text-[10px] text-rose-600 font-medium">반복 {rep}회</span>}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{r.dongho || "-"}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{r.name || "-"}</td>
                      <td className="px-3 py-2 whitespace-nowrap"><span className="inline-flex items-center gap-1"><span className="w-2 h-2 rounded-sm" style={{ background: SERIES[KIND_NAMES.indexOf(r.kind)] }} />{r.kind}</span></td>
                      <td className="px-3 py-2 min-w-[14rem]">{r.line}</td>
                      <td className="px-3 py-2 whitespace-nowrap">{r.gone ? "메모 삭제됨" : <span className="text-emerald-700 font-medium">메모 있음</span>}</td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  );
}

/* ═══════════ 스크린천 주문 관리 (GDR 타석 1~14 · 스크린룸 1~5) ═══════════
   주문 기록·타석별 사이즈/교체주기는 settings 표(key: screen_orders)에 JSON 으로 저장 → 별도 SQL 필요 없음
   교체주기 = 타석별 주문 간격 평균(그룹 평균을 한 번 더 섞어 이력이 적을 때 튀지 않게), 이력 없으면 그룹/전체 평균 */
const SC_KEY = "screen_orders";
// start = 운영 시작일(처음 설치한 천) → 첫 교체까지 걸린 기간도 주기 계산에 포함
const SC_GROUPS = [{ id: "gdr", label: "GDR 타석", n: 14 }, { id: "room", label: "스크린룸", n: 5, start: "2025-05-01" }];
const SC_START = { gdr: "", room: "2025-05-01" };
const SC_MIN_IV = 45; // 45일 안에 다시 주문한 건 불량·파손 재주문으로 보고 주기 계산에서 제외
const scName = (g, no) => (g === "gdr" ? `${no}번타석` : `스크린룸 ${no}번`);
const SC_UNITS = SC_GROUPS.flatMap((g) => Array.from({ length: g.n }, (_, i) => ({ key: `${g.id}-${i + 1}`, g: g.id, no: i + 1, name: scName(g.id, i + 1) })));
const SC_UNIT = Object.fromEntries(SC_UNITS.map((u, i) => [u.key, { ...u, idx: i }]));
const SC_L2 = " (좌암2숏)";
const SC_DEFAULT_SIZE = {
  "gdr-1": "3680 x 3130 + 230 (3숏)", "gdr-2": "3680 x 3130 + 230" + SC_L2, "gdr-3": "3180 x 3130 + 230" + SC_L2,
  ...Object.fromEntries([4, 5, 6, 7, 8, 9, 10, 11, 12, 13].map((n) => [`gdr-${n}`, "2680 x 3130 + 230" + SC_L2])),
  "gdr-14": "2480 x 3130 + 230" + SC_L2,
  "room-1": "5020 x 3000 + 320 (3숏)", "room-2": "4960 x 3000 + 320 (3숏)", "room-3": "4960 x 3000 + 320 (3숏)",
  "room-4": "4970 x 3000 + 320 (3숏)", "room-5": "4960 x 3000 + 320 (3숏)",
};
// 넘버원스크린천 문자 내역 (2025.10 ~ 2026.10) — 처음 열 때 한 번만 저장됨
const SC_SEED = [
  ["2025-10-10", "gdr", [6, 8, 10]], ["2025-10-14", "gdr", [7]], ["2025-10-22", "gdr", [12]], ["2025-11-18", "gdr", [5]],
  ["2025-11-27", "gdr", [14]], ["2025-12-06", "gdr", [13]], ["2026-02-11", "gdr", [2, 9, 10]], ["2026-02-19", "gdr", [11]],
  ["2026-02-21", "gdr", [6, 8]], ["2026-02-24", "gdr", [7]], ["2026-03-04", "gdr", [1, 3, 5]], ["2026-03-25", "gdr", [12]],
  ["2026-03-27", "gdr", [14]], ["2026-03-30", "gdr", [4]], ["2026-04-13", "gdr", [13]], ["2026-05-06", "gdr", [5, 9]],
  ["2026-05-13", "gdr", [6, 8, 10]], ["2026-05-16", "gdr", [2, 11]], ["2026-05-18", "gdr", [7]], ["2026-06-01", "room", [1, 3, 5]],
  ["2026-07-09", "gdr", [12]], ["2026-07-09", "room", [1]], ["2026-07-13", "gdr", [14]], ["2026-07-18", "gdr", [5]],
  ["2026-08-04", "gdr", [3]], ["2026-08-13", "room", [4]], ["2026-08-24", "gdr", [9, 11]], ["2026-08-26", "gdr", [8]],
  ["2026-08-31", "gdr", [6]], ["2026-09-03", "gdr", [4, 13]], ["2026-09-07", "gdr", [10]], ["2026-09-14", "gdr", [7]],
  ["2026-09-17", "gdr", [2]], ["2026-10-12", "gdr", [14]],
];
const scSeed = () => ({ orders: SC_SEED.flatMap(([date, g, nos]) => nos.map((no) => ({ id: `${date}_${g}-${no}`, date, unit: `${g}-${no}`, memo: "" }))) });
const scNorm = (v) => {
  let o = v;
  try { if (typeof v === "string") o = JSON.parse(v); } catch { o = null; }
  if (!o || !Array.isArray(o.orders)) return null;
  return {
    orders: o.orders.filter((r) => SC_UNIT[r.unit] && /^\d{4}-\d{2}-\d{2}$/.test(r.date)),
    units: Object.fromEntries(SC_UNITS.map((u) => { const x = o.units?.[u.key] || {}; return [u.key, { ...x, size: x.size || SC_DEFAULT_SIZE[u.key], cycleM: x.cycleM ?? "" }]; })),
  };
};

const SC_DAY = 86400e3;
const scDiff = (a, b) => Math.round((Date.parse(b) - Date.parse(a)) / SC_DAY);
const scAdd = (d, n) => new Date(Date.parse(d) + n * SC_DAY).toISOString().slice(0, 10);
const scShort = (d) => `${d.slice(2, 4)}.${d.slice(5, 7)}.${d.slice(8)}`;
const scWk = (d) => "일월화수목금토"[new Date(d).getUTCDay()];
const scMed = (a) => { if (!a.length) return 0; const s = [...a].sort((x, y) => x - y); const m = s.length >> 1; return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2; };
const scMonths = (days) => (days / 30.4).toFixed(1);
const scDday = (left) => (left === 0 ? "D-day" : left > 0 ? `D-${left}` : `${-left}일 지남`);
const SC_ST = {
  due: { label: "주문 필요", bar: "bg-rose-500", badge: "bg-rose-50 text-rose-600 border-rose-200", edge: "border-l-rose-500", txt: "text-rose-600" },
  soon: { label: "곧 주문", bar: "bg-amber-400", badge: "bg-amber-50 text-amber-700 border-amber-200", edge: "border-l-amber-400", txt: "text-amber-700" },
  ok: { label: "여유", bar: "bg-green-500", badge: "bg-green-50 text-green-700 border-green-200", edge: "border-l-green-500", txt: "text-green-700" },
  none: { label: "이력 없음", bar: "bg-slate-300", badge: "bg-slate-50 text-slate-500 border-slate-200", edge: "border-l-slate-300", txt: "text-slate-400" },
};

function scAnalyze(data, now) {
  const dates = {}, iv = {}, gIv = { gdr: [], room: [] };
  SC_UNITS.forEach((u) => { dates[u.key] = []; });
  data.orders.forEach((o) => dates[o.unit].push(o.date));
  const seqs = {};
  SC_UNITS.forEach((u) => {
    const d = [...new Set(dates[u.key])].sort(); dates[u.key] = d;
    const st = SC_START[u.g];
    const seq = st ? [st, ...d.filter((x) => x > st)] : d; seqs[u.key] = seq;
    iv[u.key] = seq.slice(1).map((x, i) => scDiff(seq[i], x)).filter((n) => n >= SC_MIN_IV);
    gIv[u.g].push(...iv[u.key]);
  });
  const allMed = scMed([...gIv.gdr, ...gIv.room]) || 105;
  const base = (g) => Math.round(gIv[g].length >= 3 ? scMed(gIv[g]) : allMed);
  return SC_UNITS.map((u) => {
    const d = dates[u.key], seq = seqs[u.key], v = iv[u.key], set = data.units[u.key];
    const manual = Number(set.cycleM) > 0;
    const cycle = manual ? Math.round(Number(set.cycleM) * 30.4) : Math.round((v.reduce((a, b) => a + b, 0) + base(u.g)) / (v.length + 1));
    const last = seq[seq.length - 1], fromStart = !d.length && !!last;
    const basis = manual ? "직접 지정" : v.length ? (SC_START[u.g] ? `운영시작 포함 교체 ${v.length}회 기준` : `주문 ${d.length}회 기준`)
      : fromStart ? `운영시작 후 미교체 · ${SC_GROUPS.find((g) => g.id === u.g).label} 평균` : "이력 부족 · 평균 적용";
    if (!last) return { ...u, size: set.size, dates: d, count: 0, cycle, status: "none", basis };
    const elapsed = scDiff(last, now), left = cycle - elapsed;
    return { ...u, size: set.size, dates: d, count: d.length, last, fromStart, elapsed, cycle, next: scAdd(last, cycle), left, basis,
      status: left <= 0 ? "due" : left <= 14 ? "soon" : "ok" };
  });
}

function scMessage(keys, units) {
  const list = SC_UNITS.filter((u) => keys.includes(u.key));
  const sz = (u) => units[u.key]?.size || "(사이즈 입력)";
  const body = list.length === 1
    ? `타석번호: ${list[0].g === "gdr" ? `${list[0].no}번` : list[0].name}\n사이즈: ${sz(list[0])}\n`
    : list.map((u) => `${u.name} : ${sz(u)}`).join("\n") + "\n\n";
  return `안녕하세요 올림픽파크포레온 2단지\n스크린천 ${list.length}장 구매요청드립니다\n\n${body}배송: 일반택배\n배송지: 양재대로 1300 올림픽파크포레온 스포츠컴플렉스 B1 안내데스크\n\n위 내용으로 부탁드립니다!`;
}
async function scCopy(text) {
  try { await navigator.clipboard.writeText(text); return true; } catch {
    const t = document.createElement("textarea"); t.value = text; document.body.appendChild(t); t.select();
    const ok = document.execCommand("copy"); t.remove(); return ok;
  }
}

function ScreenCloth() {
  const [data, setData] = useState(null);
  const [warn, setWarn] = useState("");
  const [view, setView] = useState("status");
  const [msg, setMsg] = useState("");
  const flash = (t) => { setMsg(t); setTimeout(() => setMsg(""), 2500); };

  React.useEffect(() => {
    let alive = true;
    (async () => {
      const seed = scNorm(scSeed());
      let row = null, error = null;
      try { ({ data: row, error } = await supabase.from("settings").select("value").eq("key", SC_KEY).maybeSingle()); }
      catch (e) { error = e; }
      if (!alive) return;
      if (error) { setWarn(`서버 저장소에 연결하지 못해 이 브라우저에만 저장됩니다 (${error.message})`); setData(scNorm(lsGet(SC_KEY)) || seed); return; }
      const cur = scNorm(row?.value);
      if (cur) { setData(cur); return; }
      setData(seed);
      if (row) { setWarn("저장된 주문 기록을 읽지 못해 기본 내역을 보여줍니다."); return; }
      let e2 = null;
      try { ({ error: e2 } = await supabase.from("settings").upsert({ key: SC_KEY, value: JSON.stringify(seed) })); } catch (e) { e2 = e; }
      if (e2) setWarn("기본 주문 내역 저장 실패: " + e2.message);
    })();
    return () => { alive = false; };
  }, []);

  const save = async (next) => {
    setData(next);
    try { localStorage.setItem(SC_KEY, JSON.stringify(next)); } catch { /* 무시 */ }
    let error = null;
    try { ({ error } = await supabase.from("settings").upsert({ key: SC_KEY, value: JSON.stringify(next) })); } catch (e) { error = e; }
    if (error) setWarn("서버 저장 실패 (이 브라우저에는 저장됨): " + error.message);
  };
  const addOrders = (date, keys, memo = "") => {
    const have = new Set(data.orders.map((o) => `${o.date}_${o.unit}`));
    const add = keys.filter((k) => !have.has(`${date}_${k}`)).map((k) => ({ id: `${date}_${k}`, date, unit: k, memo }));
    if (add.length) save({ ...data, orders: [...data.orders, ...add] });
    return add.length;
  };

  if (!data) return <p className="text-sm text-slate-400 py-10 flex items-center justify-center gap-2"><Loader2 size={14} className="animate-spin" /> 불러오는 중…</p>;
  const now = kstDay();
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex border border-byb-line bg-white">
          {[["status", "교체 현황"], ["input", "주문 입력"]].map(([id, l]) => (
            <button key={id} onClick={() => setView(id)}
              className={`px-5 h-9 text-[13px] ${view === id ? "bg-emerald-600 text-white font-medium" : "text-slate-600 hover:bg-slate-50"}`}>{l}</button>
          ))}
        </div>
        <p className="text-[11px] text-slate-400">거래처 넘버원스크린천 · 오늘 {scShort(now)} ({scWk(now)})</p>
      </div>
      {warn && <p className="text-xs text-slate-600 bg-amber-50 border border-amber-200 px-3 py-2">{warn}</p>}
      {msg && <p className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50 bg-slate-800 text-white text-xs px-4 py-2.5">{msg}</p>}
      {view === "status"
        ? <ScStatus data={data} now={now} addOrders={addOrders} flash={flash} />
        : <ScInput data={data} now={now} save={save} addOrders={addOrders} flash={flash} />}
    </div>
  );
}

function ScStatus({ data, now, addOrders, flash }) {
  const rows = useMemo(() => scAnalyze(data, now), [data, now]);
  const due = rows.filter((r) => r.status === "due").sort((a, b) => a.left - b.left);
  const soon = rows.filter((r) => r.status === "soon").sort((a, b) => a.left - b.left);
  const [pick, setPick] = useState(null);                 // null = '주문 필요' 전체 선택
  const picked = pick ?? due.map((r) => r.key);
  const toggle = (k) => setPick((p) => { const s = new Set(p ?? due.map((r) => r.key)); s.has(k) ? s.delete(k) : s.add(k); return [...s]; });
  const nextUp = rows.filter((r) => r.status === "ok").sort((a, b) => a.left - b.left)[0];
  const gdr = rows.filter((r) => r.g === "gdr" && r.count > 1);
  const gdrAvg = gdr.length ? gdr.reduce((a, r) => a + r.cycle, 0) / gdr.length : 0;
  const year = now.slice(0, 4);

  const months = [];
  {
    const first = (data.orders.map((o) => o.date).sort()[0] || now).slice(0, 7);
    let [y, m] = first.split("-").map(Number);
    for (let ym = first; ym <= now.slice(0, 7); ) { months.push(ym); m++; if (m > 12) { m = 1; y++; } ym = `${y}-${pad2(m)}`; }
  }
  const thisYm = now.slice(0, 7);

  const record = () => {
    if (!picked.length) return alert("주문할 타석을 선택하세요.");
    const names = SC_UNITS.filter((u) => picked.includes(u.key)).map((u) => u.name).join(", ");
    if (!askYesNo(`오늘(${scShort(now)}) 주문한 것으로 기록할까요?\n\n${names}`)) return;
    const n = addOrders(now, picked); setPick(null);
    flash(n ? `오늘 주문 ${n}장 기록됨` : "이미 오늘 기록된 타석입니다");
  };
  const copy = async () => {
    if (!picked.length) return alert("주문할 타석을 선택하세요.");
    await scCopy(scMessage(picked, data.units)); flash("주문 문자 복사됨 — 넘버원스크린천 문자창에 붙여넣기");
  };
  const Chip = ({ r }) => {
    const on = picked.includes(r.key);
    return (
      <button onClick={() => toggle(r.key)}
        className={`flex items-center gap-1.5 border px-2.5 py-1.5 text-[12px] ${on ? "bg-white border-rose-400 text-slate-900" : "bg-white/60 border-byb-line text-slate-500"}`}>
        <span className={`w-3.5 h-3.5 border flex items-center justify-center ${on ? "bg-rose-500 border-rose-500 text-white" : "border-slate-300"}`}>{on && <CheckCircle2 size={10} />}</span>
        <b>{r.name}</b><span className={SC_ST[r.status].txt}>{scDday(r.left)}</span>
      </button>
    );
  };

  return (
    <div className="space-y-4">
      {due.length ? (
        <div className="bg-rose-50 border border-rose-200 p-4">
          <p className="text-[15px] font-bold text-rose-700 flex items-center gap-1.5"><AlertTriangle size={17} /> 오늘 주문해야 할 스크린천 {due.length}장</p>
          <p className="text-[11px] text-rose-600/80 mt-0.5">교체주기가 지났거나 오늘인 타석입니다. 실제로 주문할 타석만 체크하세요.</p>
          <div className="flex flex-wrap gap-1.5 mt-3">{due.map((r) => <Chip key={r.key} r={r} />)}</div>
          {soon.length > 0 && <>
            <p className="text-[11px] text-slate-500 mt-3 mb-1.5">같이 주문해도 되는 2주 내 예정</p>
            <div className="flex flex-wrap gap-1.5">{soon.map((r) => <Chip key={r.key} r={r} />)}</div>
          </>}
          <div className="flex flex-wrap gap-2 mt-3">
            <button onClick={copy} className="flex items-center gap-1.5 bg-slate-800 text-white text-xs px-3 py-2"><Copy size={13} /> 주문 문자 복사 ({picked.length}장)</button>
            <button onClick={record} className="flex items-center gap-1.5 bg-rose-600 hover:bg-rose-700 text-white text-xs px-3 py-2"><CheckCircle2 size={13} /> 오늘 주문 완료로 기록</button>
          </div>
        </div>
      ) : (
        <div className="bg-green-50 border border-green-200 p-4">
          <p className="text-[15px] font-bold text-green-700 flex items-center gap-1.5"><CheckCircle2 size={17} /> 오늘은 주문할 스크린천이 없습니다</p>
          {soon.length > 0
            ? <p className="text-[12px] text-amber-700 mt-1">2주 안에 주문 예정: {soon.map((r) => `${r.name} ${scShort(r.next)}(${scDday(r.left)})`).join(" · ")}</p>
            : nextUp && <p className="text-[12px] text-slate-600 mt-1">다음 주문 예정: <b>{nextUp.name}</b> {scShort(nextUp.next)} ({scDday(nextUp.left)})</p>}
        </div>
      )}

      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        {[
          ["주문 필요", `${due.length}장`, "주기 지남·오늘", due.length ? "text-rose-600" : "text-slate-900"],
          ["2주 내 주문", `${soon.length}장`, "14일 안에 주기 도래", soon.length ? "text-amber-600" : "text-slate-900"],
          [`${year}년 주문`, `${data.orders.filter((o) => o.date.startsWith(year)).length}장`, `전체 ${data.orders.length}장 기록`, "text-slate-900"],
          ["GDR 평균 교체주기", gdrAvg ? `${scMonths(gdrAvg)}개월` : "-", gdrAvg ? `약 ${Math.round(gdrAvg)}일` : "이력 부족", "text-slate-900"],
        ].map(([l, v, f, c]) => (
          <div key={l} className="bg-white border border-byb-line px-3 py-2.5">
            <p className="text-[11px] text-slate-500">{l}</p><p className={`text-xl font-bold mt-0.5 ${c}`}>{v}</p><p className="text-[11px] text-slate-400 mt-0.5">{f}</p>
          </div>
        ))}
      </div>

      {SC_GROUPS.map((g) => (
        <div key={g.id} className="bg-white border border-byb-line">
          <div className="px-4 py-3 border-b border-byb-line flex flex-wrap items-center justify-between gap-2">
            <p className="text-[14px] font-bold text-slate-800">{g.label} <span className="font-normal text-xs text-slate-400">1~{g.n}번{g.start ? ` · ${g.start.slice(2, 4)}년 ${+g.start.slice(5, 7)}월 운영 시작` : ""}</span></p>
            <div className="flex gap-3 text-[11px] text-slate-500">
              {["due", "soon", "ok"].map((s) => <span key={s} className="flex items-center gap-1"><span className={`w-2 h-2 ${SC_ST[s].bar}`} />{SC_ST[s].label}</span>)}
            </div>
          </div>
          <div className="p-3 grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 xl:grid-cols-5 gap-2">
            {rows.filter((r) => r.g === g.id).map((r) => {
              const st = SC_ST[r.status];
              const pct = r.last ? Math.min(100, Math.round((r.elapsed / r.cycle) * 100)) : 0;
              return (
                <div key={r.key} className={`border border-byb-line border-l-4 ${st.edge} px-3 py-2.5`}>
                  <div className="flex items-center justify-between gap-1">
                    <b className="text-[16px] text-slate-900">{r.no}번</b>
                    <span className={`text-[10px] px-1.5 py-0.5 border ${st.badge}`}>{st.label}</span>
                  </div>
                  <p className="text-[10px] text-slate-400 truncate" title={r.size}>{r.size || "사이즈 미등록"}</p>
                  {r.last ? <>
                    <div className="h-1.5 bg-slate-100 mt-2"><div className={`h-full ${st.bar}`} style={{ width: `${pct}%` }} /></div>
                    <div className="flex justify-between text-[11px] mt-1.5"><span className="text-slate-500">{r.fromStart ? "운영시작" : "마지막"} {scShort(r.last)}</span><span className="text-slate-400">{r.elapsed}일 전</span></div>
                    <div className="flex justify-between text-[11px]"><span className="text-slate-500">다음 {scShort(r.next)}</span><b className={st.txt}>{scDday(r.left)}</b></div>
                    <p className="text-[10px] text-slate-400 mt-1">주기 약 {scMonths(r.cycle)}개월 · {r.basis}</p>
                  </> : <p className="text-[11px] text-slate-400 mt-2">주문 기록이 없습니다</p>}
                </div>
              );
            })}
          </div>
        </div>
      ))}

      <div className="bg-white border border-byb-line">
        <p className="px-4 py-3 border-b border-byb-line text-[14px] font-bold text-slate-800">구매 이력 <span className="font-normal text-xs text-slate-400">· 동그라미 숫자 = 주문한 날(일)</span></p>
        <div className="overflow-x-auto">
          <table className="text-[11px] border-collapse w-full">
            <thead>
              <tr className="bg-byb-bg text-slate-500">
                <th className="sticky left-0 z-10 bg-byb-bg px-3 py-2 text-left font-medium whitespace-nowrap">타석</th>
                {months.map((ym) => <th key={ym} className={`px-1 py-2 font-medium whitespace-nowrap min-w-[2.75rem] ${ym === thisYm ? "text-emerald-600" : ""}`}>{ym.slice(2, 4)}.{ym.slice(5)}</th>)}
                <th className="px-3 py-2 font-medium whitespace-nowrap">합계</th>
              </tr>
            </thead>
            <tbody>
              {SC_GROUPS.map((g) => (
                <React.Fragment key={g.id}>
                  <tr className="bg-slate-50 border-t border-byb-line"><td className="sticky left-0 bg-slate-50 px-3 py-1.5 font-bold text-slate-600 whitespace-nowrap">{g.label}</td><td colSpan={months.length + 1} /></tr>
                  {rows.filter((r) => r.g === g.id).map((r) => (
                    <tr key={r.key} className="border-t border-byb-line">
                      <td className="sticky left-0 z-10 bg-white px-3 py-1.5 whitespace-nowrap font-medium text-slate-700">{r.no}번</td>
                      {months.map((ym) => (
                        <td key={ym} className={`text-center py-1.5 whitespace-nowrap ${ym === thisYm ? "bg-emerald-50" : ""}`}>
                          {r.dates.filter((d) => d.startsWith(ym)).map((d) => (
                            <span key={d} title={`${scShort(d)} 주문`} className="inline-flex items-center justify-center w-[1.35rem] h-[1.35rem] rounded-full bg-emerald-600 text-white text-[10px] font-medium mx-px">{+d.slice(8)}</span>
                          ))}
                        </td>
                      ))}
                      <td className="text-center font-bold text-slate-700">{r.count}</td>
                    </tr>
                  ))}
                </React.Fragment>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
}

function ScInput({ data, now, save, addOrders, flash }) {
  const [date, setDate] = useState(now);
  const [sel, setSel] = useState([]);
  const [memo, setMemo] = useState("");
  const [fu, setFu] = useState("");
  const [draft, setDraft] = useState(null);
  const auto = useMemo(() => Object.fromEntries(scAnalyze({ ...data, units: Object.fromEntries(SC_UNITS.map((u) => [u.key, { ...data.units[u.key], cycleM: "" }])) }, now).map((r) => [r.key, r.cycle])), [data, now]);
  const toggle = (k) => setSel((s) => (s.includes(k) ? s.filter((x) => x !== k) : [...s, k]));

  const submit = () => {
    if (!date) return alert("주문 날짜를 입력하세요.");
    if (!sel.length) return alert("타석을 하나 이상 선택하세요.");
    const n = addOrders(date, sel, memo.trim());
    flash(n ? `${scShort(date)} 주문 ${n}장 기록됨` : "이미 같은 날짜에 기록된 타석입니다");
    if (n) { setSel([]); setMemo(""); }
  };
  const copy = async () => { if (!sel.length) return alert("타석을 선택하세요."); await scCopy(scMessage(sel, data.units)); flash("주문 문자 복사됨"); };
  const del = (o) => {
    if (askYesNo(`${scShort(o.date)} ${SC_UNIT[o.unit].name} 주문 기록을 삭제할까요?`)) save({ ...data, orders: data.orders.filter((x) => x.id !== o.id) });
  };
  const saveUnits = () => { save({ ...data, units: draft }); setDraft(null); flash("타석 설정 저장됨"); };

  const list = data.orders.filter((o) => !fu || o.unit === fu);
  const byDate = [...new Set(list.map((o) => o.date))].sort().reverse()
    .map((d) => [d, list.filter((o) => o.date === d).sort((a, b) => SC_UNIT[a.unit].idx - SC_UNIT[b.unit].idx)]);
  const inp = "border border-byb-line px-3 py-2 text-sm focus:outline-none focus:border-emerald-500";

  return (
    <div className="space-y-4">
      <Panel title="주문 입력">
        <div className="grid sm:grid-cols-[180px_1fr] gap-3">
          <L label="주문 날짜"><input type="date" value={date} onChange={(e) => setDate(e.target.value)} className={`${inp} w-full`} /></L>
          <L label="메모 (선택)"><input value={memo} onChange={(e) => setMemo(e.target.value)} placeholder="예: 찢어짐, 정정 등" className={`${inp} w-full`} /></L>
        </div>
        {SC_GROUPS.map((g) => (
          <div key={g.id} className="mt-4">
            <p className="text-[11px] font-medium text-slate-500 mb-1.5">{g.label}</p>
            <div className="flex flex-wrap gap-1.5">
              {SC_UNITS.filter((u) => u.g === g.id).map((u) => {
                const on = sel.includes(u.key);
                return (
                  <button key={u.key} onClick={() => toggle(u.key)} title={data.units[u.key].size}
                    className={`min-w-[3.25rem] h-10 px-2 border text-[13px] ${on ? "bg-emerald-600 border-emerald-600 text-white font-bold" : "bg-white border-byb-line text-slate-700 hover:border-emerald-400"}`}>{u.no}번</button>
                );
              })}
            </div>
          </div>
        ))}
        {sel.length > 0 && (
          <pre className="mt-4 bg-byb-bg border border-byb-line px-3 py-2.5 text-[12px] text-slate-700 whitespace-pre-wrap font-sans">{scMessage(sel, data.units)}</pre>
        )}
        <div className="flex flex-wrap gap-2 mt-4">
          <button onClick={submit} className="flex items-center gap-1.5 bg-emerald-600 hover:bg-emerald-700 text-white text-[13px] font-medium px-4 py-2"><Plus size={14} /> 주문 기록 저장{sel.length ? ` (${sel.length}장)` : ""}</button>
          <button onClick={copy} className="flex items-center gap-1.5 bg-slate-800 text-white text-[13px] px-4 py-2"><Copy size={14} /> 주문 문자 복사</button>
          {sel.length > 0 && <button onClick={() => setSel([])} className="text-[12px] text-slate-500 px-2">선택 해제</button>}
        </div>
      </Panel>

      <div className="bg-white border border-byb-line">
        <div className="px-4 py-3 border-b border-byb-line flex flex-wrap items-center justify-between gap-2">
          <p className="text-[14px] font-bold text-slate-800">주문 기록 <span className="font-normal text-xs text-slate-400">{list.length}장</span></p>
          <select value={fu} onChange={(e) => setFu(e.target.value)} className="border border-byb-line px-2.5 py-1.5 text-xs bg-white">
            <option value="">전체 타석</option>
            {SC_GROUPS.map((g) => <optgroup key={g.id} label={g.label}>{SC_UNITS.filter((u) => u.g === g.id).map((u) => <option key={u.key} value={u.key}>{u.name}</option>)}</optgroup>)}
          </select>
        </div>
        {!byDate.length ? <p className="text-sm text-slate-400 py-10 text-center">기록이 없습니다.</p> : (
          <div className="max-h-[32rem] overflow-y-auto divide-y divide-byb-line">
            {byDate.map(([d, os]) => (
              <div key={d} className="px-4 py-2.5 flex flex-wrap items-center gap-x-3 gap-y-1.5">
                <span className="w-[6.5rem] text-[12px] font-medium text-slate-700 whitespace-nowrap">{scShort(d)} ({scWk(d)})</span>
                <div className="flex flex-wrap gap-1.5 flex-1">
                  {os.map((o) => (
                    <span key={o.id} className={`inline-flex items-center gap-1 border pl-2 pr-1 py-0.5 text-[12px] ${o.unit.startsWith("room") ? "border-violet-200 bg-violet-50 text-violet-700" : "border-emerald-200 bg-emerald-50 text-emerald-700"}`}>
                      {SC_UNIT[o.unit].name}{o.memo && <span className="text-slate-400">· {o.memo}</span>}
                      <button onClick={() => del(o)} title="삭제" className="text-slate-400 hover:text-rose-600"><X size={12} /></button>
                    </span>
                  ))}
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      <div className="bg-white border border-byb-line">
        <button onClick={() => setDraft(draft ? null : JSON.parse(JSON.stringify(data.units)))}
          className="w-full px-4 py-3 text-left text-[14px] font-bold text-slate-800 flex items-center justify-between">
          타석 설정 <span className="font-normal text-xs text-slate-400">사이즈 · 교체주기 직접 지정 {draft ? "▲" : "▼"}</span>
        </button>
        {draft && (
          <div className="border-t border-byb-line p-4">
            <p className="text-[11px] text-slate-500 mb-3">교체주기를 비워두면 주문 이력으로 자동 계산합니다. 이력이 적은 타석(1번, 스크린룸)은 직접 넣어두면 알림이 정확해집니다.</p>
            <div className="overflow-x-auto">
              <table className="w-full text-[12px]">
                <thead><tr className="text-slate-500 text-left"><th className="font-medium py-1.5 pr-3">타석</th><th className="font-medium py-1.5 pr-3">사이즈 (문자에 그대로 들어감)</th><th className="font-medium py-1.5">교체주기(개월)</th></tr></thead>
                <tbody>
                  {SC_UNITS.map((u) => (
                    <tr key={u.key} className="border-t border-byb-line">
                      <td className="py-1.5 pr-3 whitespace-nowrap font-medium text-slate-700">{u.name}</td>
                      <td className="py-1.5 pr-3"><input value={draft[u.key].size} onChange={(e) => setDraft({ ...draft, [u.key]: { ...draft[u.key], size: e.target.value } })} className="w-full min-w-[15rem] border border-byb-line px-2 py-1" placeholder="예: 4960 x 3000 + 320 (3숏)" /></td>
                      <td className="py-1.5"><input value={draft[u.key].cycleM} onChange={(e) => setDraft({ ...draft, [u.key]: { ...draft[u.key], cycleM: e.target.value.replace(/[^0-9.]/g, "") } })} className="w-24 border border-byb-line px-2 py-1" placeholder={`자동 ${scMonths(auto[u.key])}`} /></td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
            <div className="flex gap-2 mt-3">
              <button onClick={saveUnits} className="bg-emerald-600 hover:bg-emerald-700 text-white text-[13px] font-medium px-4 py-2">설정 저장</button>
              <button onClick={() => setDraft(null)} className="text-[12px] text-slate-500 px-2">취소</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}

/* settings 표에 JSON 하나로 저장하는 공용 훅 (실패하면 이 브라우저 localStorage 에 임시 저장) */
function useSettingJson(key, seed, norm) {
  const [data, setData] = useState(null);
  const [warn, setWarn] = useState("");
  React.useEffect(() => {
    let alive = true;
    (async () => {
      let row = null, error = null;
      try { ({ data: row, error } = await supabase.from("settings").select("value").eq("key", key).maybeSingle()); } catch (e) { error = e; }
      if (!alive) return;
      if (error) { setWarn(`서버 저장소에 연결하지 못해 이 브라우저에만 저장됩니다 (${error.message})`); setData(norm(lsGet(key)) || seed()); return; }
      const cur = norm(row?.value);
      if (cur) { setData(cur); return; }
      setData(seed());
      if (row) setWarn("저장된 값을 읽지 못해 기본값을 보여줍니다.");
    })();
    return () => { alive = false; };
  }, [key]);
  const save = async (next) => {
    setData(next);
    try { localStorage.setItem(key, JSON.stringify(next)); } catch { /* 무시 */ }
    let error = null;
    try { ({ error } = await supabase.from("settings").upsert({ key, value: JSON.stringify(next) })); } catch (e) { error = e; }
    if (error) setWarn("서버 저장 실패 (이 브라우저에는 저장됨): " + error.message);
  };
  return [data, save, warn];
}

/* ═══════════ 시설 청소 관리 ═══════════
   구역마다 청소주기(일)를 두고 → 이번 주(월~일)에 해야 할 구역을 자동 배정 · 완료 체크 · 기록
   처음 하는 월간/분기 구역은 같은 시설 안에서 주차별로 나눠 배정(한 주에 몰리지 않게)
   저장: settings 표 key "cleaning" (JSON). 구역을 화면에서 직접 고치기 전까지는 코드의 기본 구역(CL_VER)을 따라감 */
const CL_KEY = "cleaning";
const CL_VER = 1;
const CL_FLOORS = ["2F", "1F", "B1", "B2"];
const CL_FACS = [
  { id: "tennis", f: "2F", name: "스크린테니스" }, { id: "pingpong", f: "2F", name: "탁구장" },
  { id: "fitness", f: "1F", name: "피트니스" }, { id: "futsal", f: "1F", name: "풋살장" },
  { id: "golf", f: "B1", name: "골프연습장" }, { id: "screengolf", f: "B1", name: "스크린골프룸" }, { id: "lounge", f: "B1", name: "골프라운지(카페)" },
  { id: "sauna_m1", f: "B2", name: "남자 사우나 1" }, { id: "sauna_m2", f: "B2", name: "남자 사우나 2" },
  { id: "sauna_w1", f: "B2", name: "여자 사우나 1" }, { id: "sauna_w2", f: "B2", name: "여자 사우나 2" },
  { id: "pool", f: "B2", name: "수영장" },
];
const CL_FAC = Object.fromEntries(CL_FACS.map((f) => [f.id, f]));
const CL_CYC = [[1, "매일"], [7, "주 1회"], [14, "2주 1회"], [30, "월 1회"], [90, "분기 1회"], [180, "반기 1회"]];
const clCyc = (c) => (CL_CYC.find((x) => x[0] === c) || [0, `${c}일마다`])[1];
// 기본 구역(대략) — 시설별 체크목록 받으면 여기를 고치고 CL_VER 올림
const CL_SAUNA = [["탕(온탕·냉탕) 물빼기·청소", 7], ["샤워부스·수전·거울", 7], ["바닥 물때·배수구", 7], ["사우나실 내부(의자·벽)", 7],
  ["파우더룸 세면대·드라이기·거울", 7], ["탈의실 락커 상부·내부", 30], ["벽·천장 타일 줄눈 곰팡이", 30], ["환기구·조명", 90]];
const CL_DEFAULT = {
  tennis: [["코트 바닥(매트)", 7], ["대기석·라켓 거치대", 7], ["출입문·유리", 7], ["스크린·센서 먼지", 30], ["볼머신 주변·공 정리", 30], ["에어컨 필터·조명", 90]],
  pingpong: [["바닥", 7], ["탁구대 표면", 7], ["대기 의자·출입문", 7], ["칸막이(펜스)·네트", 30], ["창문·유리", 30], ["에어컨 필터·조명", 90]],
  fitness: [["유산소 기구(러닝머신 벨트·하부)", 7], ["웨이트 기구 손잡이·패드 소독", 7], ["거울·유리", 7], ["덤벨·랙·소도구 정리", 7], ["스트레칭존 매트", 7],
    ["기구 하부 바닥 먼지", 30], ["에어컨 필터·환기구", 90]],
  futsal: [["대기 벤치·출입구", 7], ["인조잔디 이물질 제거", 30], ["골대·네트", 30], ["펜스·조명", 90]],
  golf: [["타석 매트·주변 바닥", 7], ["칸막이·대기석", 7], ["클럽 거치대·볼 박스", 30], ["네트·스크린 먼지", 30], ["천장·조명", 90]],
  screengolf: [["룸 바닥·스윙 매트", 7], ["소파·테이블", 7], ["출입문·유리", 7], ["스크린 먼지(털이개)", 30], ["센서·프로젝터 렌즈", 30], ["에어컨 필터", 90]],
  lounge: [["테이블·의자", 7], ["카운터·커피머신 주변", 7], ["바닥", 7], ["냉장고·제빙기 내부", 30], ["유리창·출입문", 30], ["에어컨 필터", 90]],
  sauna_m1: CL_SAUNA, sauna_m2: CL_SAUNA, sauna_w1: CL_SAUNA, sauna_w2: CL_SAUNA,
  pool: [["풀사이드 바닥·배수구", 7], ["수조 바닥·벽면(수중 청소기)", 7], ["샤워실 연결 통로·발판", 7], ["킥판·풀부이 등 강습 용품", 30], ["관람석·유리창", 30], ["천장·조명·환기구", 90]],
};
const clZones = () => CL_FACS.flatMap((f) => (CL_DEFAULT[f.id] || []).map(([name, cyc], i) => ({ id: `${f.id}-${i + 1}`, fac: f.id, name, cyc })));
const clMon = (d) => scAdd(d, -((new Date(d).getUTCDay() + 6) % 7));   // 그 주 월요일
const clMD = (d) => `${+d.slice(5, 7)}/${+d.slice(8)}`;

function clPlan(data, today) {
  const W = clMon(today), E = scAdd(W, 6), W0 = clMon(data.start || today);
  const last = {};
  data.logs.forEach((l) => { if (!last[l.zone] || l.date > last[l.zone]) last[l.zone] = l.date; });
  const rank = {}, cnt = {};
  data.zones.forEach((z) => { const k = `${z.fac}|${z.cyc}`; cnt[k] = (cnt[k] ?? -1) + 1; rank[z.id] = cnt[k]; });
  return data.zones.map((z) => {
    const L = last[z.id];
    if (z.cyc <= 1) return { ...z, last: L, due: today, thisWeek: true, state: L === today ? "done" : "todo", late: false };
    if (L && L >= W) return { ...z, last: L, thisWeek: true, state: "done", late: false };
    const due = L ? scAdd(L, z.cyc) : scAdd(W0, (rank[z.id] % Math.max(1, Math.round(z.cyc / 7))) * 7);
    const thisWeek = due <= E;
    return { ...z, last: L, due, thisWeek, state: thisWeek ? "todo" : "later", late: due < W, lateDays: due < W ? scDiff(due, today) : 0 };
  });
}

// 이번 주 배정: 시설마다 '해야 할 구역'을 청소 요일에 하루 한 구역씩 (밀린 것 먼저). 주가 바뀌면 자동으로 새로 배정
const CL_WD = [["월", 1], ["화", 2], ["수", 3], ["목", 4], ["금", 5], ["토", 6], ["일", 0]];
const CL_DAYS0 = [2, 3, 4, 5, 6];   // 기본 청소 요일: 화~토 (월 휴관, 일요일은 방문객이 많아 지시 안 함)
const CL_MSG = {
  head: "[클럽포레온 2단지] {날짜} 청소 구역 안내\n\n안녕하세요! 오늘 시설별로 아래 구역 청소 부탁드립니다.",
  foot: "평소 청소와 함께 위 구역 한 번 더 신경 써주세요.\n완료되면 이 방에 '완료' 남겨주시면 기록해두겠습니다. 감사합니다!",
};
function clMakePlan(data, today) {
  const W = clMon(today), days = data.days || CL_DAYS0;
  const dates = Array.from({ length: 7 }, (_, i) => scAdd(W, i)).filter((d) => d >= today && days.includes(new Date(d).getUTCDay()));
  const plan = clPlan(data, today), map = {};
  if (dates.length) CL_FACS.forEach((f) => {
    plan.filter((z) => z.fac === f.id && z.cyc > 1 && z.state === "todo")
      .sort((a, b) => (b.late ? 1 : 0) - (a.late ? 1 : 0) || a.due.localeCompare(b.due))
      .forEach((z, i) => { map[z.id] = dates[i % dates.length]; });
  });
  return { week: W, dates, map };
}
const clFacLines = (zones, pick) => CL_FACS.map((fac) => {
  const zs = zones.filter((z) => z.fac === fac.id && pick(z));
  return zs.length ? { f: fac.f, line: `- ${fac.name}: ${zs.map((z) => z.name).join(", ")}` } : null;
}).filter(Boolean);
function clDayMsg(data, date) {
  const map = data.plan?.map || {}, m = { ...CL_MSG, ...(data.msg || {}) };
  const rows = clFacLines(data.zones, (z) => map[z.id] === date);
  const body = CL_FLOORS.map((f) => { const r = rows.filter((x) => x.f === f); return r.length ? [`■ ${f}`, ...r.map((x) => x.line)].join("\n") : ""; }).filter(Boolean);
  const daily = clFacLines(data.zones, (z) => z.cyc <= 1);
  if (daily.length) body.push(["■ 매일 기본 청소", ...daily.map((x) => x.line)].join("\n"));
  return `${m.head.replaceAll("{날짜}", `${clMD(date)}(${scWk(date)})`)}\n\n${body.join("\n\n") || "오늘은 정해진 구역이 없습니다."}\n\n${m.foot}`;
}
function clWeekMsg(data) {
  const p = data.plan || { dates: [], map: {} };
  if (!p.dates?.length) return "";
  const body = CL_FLOORS.map((f) => {
    const r = CL_FACS.filter((x) => x.f === f).map((fac) => {
      const parts = p.dates.map((d) => { const zs = data.zones.filter((z) => z.fac === fac.id && p.map[z.id] === d); return zs.length ? `${scWk(d)} ${zs.map((z) => z.name).join(", ")}` : null; }).filter(Boolean);
      return parts.length ? `- ${fac.name}: ${parts.join(" / ")}` : null;
    }).filter(Boolean);
    return r.length ? [`■ ${f}`, ...r].join("\n") : "";
  }).filter(Boolean);
  return `[클럽포레온 2단지] 이번 주 청소 일정 (${clMD(p.dates[0])}~${clMD(p.dates[p.dates.length - 1])})\n\n안녕하세요! 이번 주 시설별 청소 구역 일정입니다.\n요일별 상세 안내는 당일 아침에 다시 올려드리겠습니다.\n\n${body.join("\n\n")}\n\n이번 주도 수고 많으십니다. 감사합니다!`;
}

function Cleaning({ me }) {
  const today = kstDay();
  const [data, save, warn] = useSettingJson(CL_KEY,
    () => ({ ver: CL_VER, custom: false, start: today, zones: clZones(), logs: [], days: CL_DAYS0, plan: null, msg: {} }),
    (v) => {
      let o = v;
      try { if (typeof v === "string") o = JSON.parse(v); } catch { o = null; }
      if (!o || !Array.isArray(o.zones)) return null;
      const zones = !o.custom && o.ver !== CL_VER ? clZones() : o.zones;
      return { ver: CL_VER, custom: !!o.custom, start: o.start || today, zones, logs: Array.isArray(o.logs) ? o.logs : [],
        days: o.daysSet && Array.isArray(o.days) ? o.days : CL_DAYS0, daysSet: !!o.daysSet, msg: o.msg || {},
        // 기본 구역이 바뀌었거나, 기본 요일이 바뀌었는데 직접 정한 적 없으면 이번 주 다시 배정
        plan: (!o.custom && o.ver !== CL_VER) || (!o.daysSet && JSON.stringify(o.days) !== JSON.stringify(CL_DAYS0)) ? null : o.plan || null };
    });
  const [view, setView] = useState("week");
  const [msg, setMsg] = useState("");
  const flash = (t) => { setMsg(t); setTimeout(() => setMsg(""), 2000); };
  React.useEffect(() => {
    if (data && data.plan?.week !== clMon(today)) save({ ...data, plan: clMakePlan(data, today) });
  }, [data]); // eslint-disable-line react-hooks/exhaustive-deps
  if (!data) return <p className="text-sm text-slate-400 py-10 flex items-center justify-center gap-2"><Loader2 size={14} className="animate-spin" /> 불러오는 중…</p>;
  const W = clMon(today);
  return (
    <div className="space-y-4">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div className="flex border border-byb-line bg-white">
          {[["week", "이번 주 청소"], ["notice", "단톡방 안내"], ["zones", "구역·주기 설정"], ["logs", "완료 기록"]].map(([id, l]) => (
            <button key={id} onClick={() => setView(id)}
              className={`px-4 h-9 text-[13px] ${view === id ? "bg-emerald-600 text-white font-medium" : "text-slate-600 hover:bg-slate-50"}`}>{l}</button>
          ))}
        </div>
        <p className="text-[11px] text-slate-400">이번 주 {clMD(W)}(월) ~ {clMD(scAdd(W, 6))}(일) · 오늘 {clMD(today)}({scWk(today)})</p>
      </div>
      {warn && <p className="text-xs text-slate-600 bg-amber-50 border border-amber-200 px-3 py-2">{warn}</p>}
      {msg && <p className="fixed bottom-6 left-1/2 -translate-x-1/2 z-50 bg-slate-800 text-white text-xs px-4 py-2.5">{msg}</p>}
      {view === "week" && <ClWeek data={data} save={save} me={me} today={today} flash={flash} />}
      {view === "notice" && <ClNotice data={data} save={save} today={today} flash={flash} />}
      {view === "zones" && <ClZones data={data} save={save} today={today} flash={flash} />}
      {view === "logs" && <ClLogs data={data} save={save} />}
    </div>
  );
}

function ClWeek({ data, save, me, today, flash }) {
  const [day, setDay] = useState(today);
  const [fl, setFl] = useState("");
  const plan = useMemo(() => clPlan(data, today), [data, today]);
  const W = clMon(today);
  const wk = plan.filter((z) => z.thisWeek);
  const done = wk.filter((z) => z.state === "done").length;
  const late = wk.filter((z) => z.state === "todo" && z.late).length;
  const pct = wk.length ? Math.round((done / wk.length) * 100) : 0;

  const check = (z) => {
    if (z.state === "done") {
      const from = z.cyc <= 1 ? today : W;
      if (!askYesNo(`${CL_FAC[z.fac].name} · ${z.name}\n완료 체크를 취소할까요?`)) return;
      save({ ...data, logs: data.logs.filter((l) => !(l.zone === z.id && l.date >= from)) });
      return;
    }
    save({ ...data, logs: [...data.logs, { id: `${Date.now().toString(36)}-${z.id}`, zone: z.id, date: day, by: me?.id || "" }] });
    flash(`${z.name} 완료 (${clMD(day)})`);
  };
  const order = { todo: 0, done: 1 };
  const pmap = data.plan?.week === W ? data.plan.map : {};
  return (
    <div className="space-y-4">
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        {[["이번 주 할 일", `${wk.length}건`, "주기상 이번 주 차례", "text-slate-900"],
          ["완료", `${done}건`, `진행률 ${pct}%`, "text-emerald-600"],
          ["남은 것", `${wk.length - done}건`, "이번 주 일요일까지", wk.length - done ? "text-amber-600" : "text-slate-900"],
          ["밀린 것", `${late}건`, "지난주 이전에 했어야 함", late ? "text-rose-600" : "text-slate-900"]].map(([l, v, f, c]) => (
          <div key={l} className="bg-white border border-byb-line px-3 py-2.5">
            <p className="text-[11px] text-slate-500">{l}</p><p className={`text-xl font-bold mt-0.5 ${c}`}>{v}</p><p className="text-[11px] text-slate-400 mt-0.5">{f}</p>
          </div>
        ))}
      </div>
      <div className="h-2 bg-white border border-byb-line"><div className="h-full bg-emerald-600" style={{ width: `${pct}%` }} /></div>

      <div className="flex flex-wrap items-center gap-2">
        <div className="flex flex-wrap gap-1">
          {["", ...CL_FLOORS].map((f) => (
            <button key={f || "all"} onClick={() => setFl(f)}
              className={`text-[12px] px-3 py-1.5 border ${fl === f ? "bg-emerald-600 border-emerald-600 text-white" : "bg-white border-byb-line text-slate-600"}`}>{f || "전체"}</button>
          ))}
        </div>
        <label className="ml-auto flex items-center gap-1.5 text-[11px] text-slate-500">체크 날짜
          <input type="date" value={day} max={today} onChange={(e) => setDay(e.target.value || today)} className="border border-byb-line px-2 py-1 text-xs" />
        </label>
      </div>

      {CL_FLOORS.filter((f) => !fl || fl === f).map((f) => (
        <div key={f}>
          <p className="text-[13px] font-bold text-slate-600 mb-2">{f}</p>
          <div className="grid md:grid-cols-2 gap-3">
            {CL_FACS.filter((x) => x.f === f).map((fac) => {
              const items = wk.filter((z) => z.fac === fac.id).sort((a, b) => order[a.state] - order[b.state] || (b.late ? 1 : 0) - (a.late ? 1 : 0) || (pmap[a.id] || "z").localeCompare(pmap[b.id] || "z"));
              const d = items.filter((z) => z.state === "done").length;
              const next = plan.filter((z) => z.fac === fac.id && z.state === "later").sort((a, b) => a.due.localeCompare(b.due))[0];
              return (
                <div key={fac.id} className="bg-white border border-byb-line">
                  <div className="px-4 py-2.5 border-b border-byb-line flex items-center justify-between gap-2">
                    <b className="text-[14px] text-slate-800">{fac.name}</b>
                    <span className={`text-[11px] ${items.length && d === items.length ? "text-emerald-600 font-bold" : "text-slate-500"}`}>
                      {items.length ? (d === items.length ? "이번 주 완료 ✓" : `${d}/${items.length} 완료`) : "이번 주 없음"}</span>
                  </div>
                  <div className="divide-y divide-byb-line">
                    {items.map((z) => {
                      const on = z.state === "done", pd = pmap[z.id], miss = !on && pd && pd < today;
                      return (
                        <button key={z.id} onClick={() => check(z)} className="w-full flex items-center gap-2.5 px-4 py-2 text-left hover:bg-slate-50">
                          <span className={`w-[18px] h-[18px] shrink-0 border flex items-center justify-center ${on ? "bg-emerald-600 border-emerald-600 text-white" : z.late ? "border-rose-400" : "border-slate-300"}`}>
                            {on && <CheckCircle2 size={12} />}</span>
                          <span className={`flex-1 text-[13px] ${on ? "text-slate-400 line-through" : "text-slate-800"}`}>{z.name}</span>
                          <span className="text-[10px] px-1.5 py-0.5 border border-byb-line text-slate-500 whitespace-nowrap">{clCyc(z.cyc)}</span>
                          <span className={`w-[4.5rem] text-right text-[11px] whitespace-nowrap ${on ? "text-emerald-600" : z.late || miss ? "text-rose-600 font-medium" : pd === today ? "text-emerald-600 font-bold" : "text-slate-400"}`}>
                            {on ? `${clMD(z.last)} 완료` : z.late ? `${z.lateDays}일 밀림` : z.cyc <= 1 || pd === today ? "오늘" : miss ? `${clMD(pd)} 못함` : pd ? `${clMD(pd)}(${scWk(pd)})` : "이번 주"}</span>
                        </button>
                      );
                    })}
                    {!items.length && <p className="px-4 py-3 text-[12px] text-slate-400">이번 주에 할 구역이 없습니다.</p>}
                  </div>
                  {next && <p className="px-4 py-2 border-t border-byb-line text-[11px] text-slate-400">다음: {next.name} · {clMD(clMon(next.due))} 주</p>}
                </div>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
}

function ClZones({ data, save, today, flash }) {
  const [draft, setDraft] = useState(() => data.zones.map((z) => ({ ...z })));
  const dirty = JSON.stringify(draft) !== JSON.stringify(data.zones);
  const set = (id, k, v) => setDraft((d) => d.map((z) => (z.id === id ? { ...z, [k]: v } : z)));
  const add = (fac) => setDraft((d) => [...d, { id: `${fac}-${Date.now().toString(36)}`, fac, name: "", cyc: 7 }]);
  const del = (z) => { if (!z.name || askYesNo(`'${z.name}' 구역을 삭제할까요? (완료 기록은 남습니다)`)) setDraft((d) => d.filter((x) => x.id !== z.id)); };
  const submit = () => {
    const zones = draft.filter((z) => z.name.trim()).map((z) => ({ ...z, name: z.name.trim() }));
    const next = { ...data, custom: true, zones }; save({ ...next, plan: clMakePlan(next, today) }); setDraft(zones); flash("구역·주기 저장됨 · 이번 주 재배정");
  };
  return (
    <div className="space-y-3 pb-14">
      <p className="text-[12px] text-slate-500">시설별 청소 구역과 주기를 정합니다. 월 1회·분기 1회 구역은 같은 시설 안에서 주차별로 나눠 배정됩니다.</p>
      {CL_FLOORS.map((f) => (
        <div key={f}>
          <p className="text-[13px] font-bold text-slate-600 mb-2">{f}</p>
          <div className="grid md:grid-cols-2 gap-3">
            {CL_FACS.filter((x) => x.f === f).map((fac) => (
              <div key={fac.id} className="bg-white border border-byb-line">
                <p className="px-4 py-2.5 border-b border-byb-line text-[14px] font-bold text-slate-800">{fac.name}</p>
                <div className="p-3 space-y-1.5">
                  {draft.filter((z) => z.fac === fac.id).map((z) => (
                    <div key={z.id} className="flex items-center gap-1.5">
                      <input value={z.name} onChange={(e) => set(z.id, "name", e.target.value)} placeholder="청소 구역"
                        className="flex-1 min-w-0 border border-byb-line px-2 py-1.5 text-[13px] focus:outline-none focus:border-emerald-500" />
                      <select value={z.cyc} onChange={(e) => set(z.id, "cyc", +e.target.value)} className="border border-byb-line px-1.5 py-1.5 text-[12px] bg-white">
                        {CL_CYC.map(([v, l]) => <option key={v} value={v}>{l}</option>)}
                      </select>
                      <button onClick={() => del(z)} className="w-7 h-7 flex items-center justify-center text-slate-400 hover:text-rose-600"><X size={14} /></button>
                    </div>
                  ))}
                  <button onClick={() => add(fac.id)} className="flex items-center gap-1 text-[12px] text-emerald-600 px-1 py-1"><Plus size={13} /> 구역 추가</button>
                </div>
              </div>
            ))}
          </div>
        </div>
      ))}
      {dirty && (
        <div className="fixed bottom-0 left-0 right-0 md:left-[200px] z-30 bg-white border-t border-byb-line px-4 py-3 flex items-center justify-end gap-2">
          <span className="text-[12px] text-slate-500 mr-auto">저장하지 않은 변경이 있습니다</span>
          <button onClick={() => setDraft(data.zones.map((z) => ({ ...z })))} className="text-[12px] text-slate-500 px-3">되돌리기</button>
          <button onClick={submit} className="bg-emerald-600 hover:bg-emerald-700 text-white text-[13px] font-medium px-5 py-2">변경 저장</button>
        </div>
      )}
    </div>
  );
}

function ClLogs({ data, save }) {
  const [ff, setFf] = useState("");
  const zone = Object.fromEntries(data.zones.map((z) => [z.id, z]));
  const facOf = (l) => zone[l.zone]?.fac || l.zone.replace(/-[^-]+$/, "");
  const list = data.logs.filter((l) => !ff || facOf(l) === ff).sort((a, b) => b.date.localeCompare(a.date) || b.id.localeCompare(a.id));
  const del = (l) => { if (askYesNo(`${clMD(l.date)} ${zone[l.zone]?.name || "삭제된 구역"} 완료 기록을 지울까요?`)) save({ ...data, logs: data.logs.filter((x) => x.id !== l.id) }); };
  return (
    <div className="bg-white border border-byb-line">
      <div className="px-4 py-3 border-b border-byb-line flex flex-wrap items-center justify-between gap-2">
        <p className="text-[14px] font-bold text-slate-800">완료 기록 <span className="font-normal text-xs text-slate-400">{list.length}건</span></p>
        <select value={ff} onChange={(e) => setFf(e.target.value)} className="border border-byb-line px-2.5 py-1.5 text-xs bg-white">
          <option value="">전체 시설</option>{CL_FACS.map((f) => <option key={f.id} value={f.id}>{f.f} {f.name}</option>)}
        </select>
      </div>
      {!list.length ? <p className="text-sm text-slate-400 py-10 text-center">아직 완료 기록이 없습니다. [이번 주 청소]에서 체크하면 쌓입니다.</p> : (
        <div className="max-h-[36rem] overflow-y-auto">
          <table className="w-full text-[12px]">
            <thead className="bg-byb-bg text-slate-500 sticky top-0"><tr>{["날짜", "시설", "구역", "체크", ""].map((h) => <th key={h} className="text-left font-medium px-4 py-2 whitespace-nowrap">{h}</th>)}</tr></thead>
            <tbody>
              {list.map((l) => (
                <tr key={l.id} className="border-t border-byb-line text-slate-700">
                  <td className="px-4 py-2 whitespace-nowrap">{scShort(l.date)} ({scWk(l.date)})</td>
                  <td className="px-4 py-2 whitespace-nowrap">{CL_FAC[facOf(l)]?.name || "-"}</td>
                  <td className="px-4 py-2">{zone[l.zone]?.name || <span className="text-slate-400">삭제된 구역</span>}</td>
                  <td className="px-4 py-2 whitespace-nowrap text-slate-400">{l.by}</td>
                  <td className="px-2 py-2 text-right"><button onClick={() => del(l)} className="text-slate-400 hover:text-rose-600"><Trash2 size={13} /></button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function ClNotice({ data, save, today, flash }) {
  const [edit, setEdit] = useState(null);
  const days = data.days || CL_DAYS0;
  const p = data.plan || { dates: [], map: {} };
  const regen = (next) => save({ ...next, plan: clMakePlan(next, today) });
  const toggleDay = (n) => {
    const nd = days.includes(n) ? days.filter((x) => x !== n) : [...days, n];
    if (!nd.length) return alert("청소 요일을 하루 이상 선택하세요.");
    regen({ ...data, days: nd, daysSet: true }); flash("청소 요일 변경 · 이번 주 재배정");
  };
  const copy = async (t, label) => { await scCopy(t); flash(`${label} 문구 복사됨`); };
  const weekMsg = clWeekMsg(data);
  const m = { ...CL_MSG, ...(data.msg || {}) };
  const ta = "w-full border border-byb-line px-3 py-2 text-[12px] focus:outline-none focus:border-emerald-500";
  return (
    <div className="space-y-4">
      <div className="bg-emerald-50 border border-emerald-200 px-4 py-3 text-[12px] text-slate-700 leading-relaxed">
        <b className="text-emerald-700">카톡 예약 메시지로 자동 발송하기</b><br />
        미화 단톡방 → 입력창 왼쪽 <b>+</b> → <b>예약 메시지</b> → 날짜·시간(예: 당일 오전 7시) 정하고 아래 요일별 문구를 붙여넣기.
        월요일에 한 주치를 한 번에 걸어두면 매일 자동으로 올라갑니다. (PC 카톡은 채팅방 메뉴에서 예약 메시지)
      </div>

      <div className="bg-white border border-byb-line px-4 py-3 flex flex-wrap items-center gap-2">
        <span className="text-[12px] text-slate-500 mr-1">청소 요일</span>
        {CL_WD.map(([l, n]) => (
          <button key={n} onClick={() => toggleDay(n)}
            className={`w-9 h-8 border text-[12px] ${days.includes(n) ? "bg-emerald-600 border-emerald-600 text-white font-medium" : "bg-white border-byb-line text-slate-500"}`}>{l}</button>
        ))}
        <span className="text-[11px] text-slate-400">시설마다 하루 한 구역씩 배정 (구역이 요일보다 많으면 하루 2개)</span>
        <button onClick={() => { if (askYesNo("오늘 이후 남은 구역을 다시 배정할까요? (완료한 구역은 빠집니다)")) { regen(data); flash("이번 주 재배정 완료"); } }}
          className="ml-auto flex items-center gap-1 border border-byb-line text-[12px] text-slate-600 px-3 py-1.5 hover:bg-slate-50"><RefreshCw size={12} /> 다시 배정</button>
      </div>

      {weekMsg && (
        <div className="bg-white border border-byb-line">
          <div className="px-4 py-2.5 border-b border-byb-line flex items-center justify-between gap-2">
            <b className="text-[14px] text-slate-800">이번 주 전체 일정 <span className="font-normal text-xs text-slate-400">· 월요일 공지용</span></b>
            <button onClick={() => copy(weekMsg, "주간 일정")} className="flex items-center gap-1 bg-slate-800 text-white text-[12px] px-3 py-1.5"><Copy size={12} /> 복사</button>
          </div>
          <pre className="px-4 py-3 text-[12px] text-slate-700 whitespace-pre-wrap font-sans max-h-64 overflow-y-auto">{weekMsg}</pre>
        </div>
      )}

      {!p.dates?.length ? <p className="text-sm text-slate-400 py-10 text-center bg-white border border-byb-line">이번 주에 남은 청소 요일이 없습니다.</p> : (
        <div className="grid md:grid-cols-2 gap-3">
          {p.dates.map((d) => {
            const t = clDayMsg(data, d), n = Object.values(p.map).filter((x) => x === d).length;
            return (
              <div key={d} className={`bg-white border ${d === today ? "border-emerald-400" : "border-byb-line"}`}>
                <div className="px-4 py-2.5 border-b border-byb-line flex items-center justify-between gap-2">
                  <b className="text-[14px] text-slate-800">{clMD(d)} ({scWk(d)}) {d === today && <span className="text-[11px] text-emerald-600">오늘</span>}
                    <span className="font-normal text-xs text-slate-400"> · {n}구역</span></b>
                  <button onClick={() => copy(t, `${clMD(d)}(${scWk(d)})`)} className="flex items-center gap-1 bg-emerald-600 hover:bg-emerald-700 text-white text-[12px] px-3 py-1.5"><Copy size={12} /> 복사</button>
                </div>
                <pre className="px-4 py-3 text-[12px] text-slate-700 whitespace-pre-wrap font-sans max-h-72 overflow-y-auto">{t}</pre>
              </div>
            );
          })}
        </div>
      )}

      <div className="bg-white border border-byb-line">
        <button onClick={() => setEdit(edit ? null : { head: m.head, foot: m.foot })}
          className="w-full px-4 py-3 text-left text-[14px] font-bold text-slate-800 flex items-center justify-between">
          안내 문구 수정 <span className="font-normal text-xs text-slate-400">머리말 · 맺음말 {edit ? "▲" : "▼"}</span>
        </button>
        {edit && (
          <div className="border-t border-byb-line p-4 space-y-3">
            <L label="머리말 ({날짜} 자리에 '11/4(수)'처럼 들어갑니다)"><textarea rows={3} value={edit.head} onChange={(e) => setEdit({ ...edit, head: e.target.value })} className={ta} /></L>
            <L label="맺음말"><textarea rows={3} value={edit.foot} onChange={(e) => setEdit({ ...edit, foot: e.target.value })} className={ta} /></L>
            <div className="flex gap-2">
              <button onClick={() => { save({ ...data, msg: edit }); setEdit(null); flash("안내 문구 저장됨"); }} className="bg-emerald-600 hover:bg-emerald-700 text-white text-[13px] font-medium px-4 py-2">저장</button>
              <button onClick={() => { setEdit({ ...CL_MSG }); }} className="text-[12px] text-slate-500 px-2">기본 문구로</button>
            </div>
          </div>
        )}
      </div>
    </div>
  );
}
