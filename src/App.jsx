import React, { useState, useMemo } from "react";
import {
  LogIn, LogOut, Search, PackageSearch, ClipboardList, CalendarDays, Sparkles, BarChart3,
  Plus, X, User, Loader2, MessageSquareWarning, ArrowLeftRight, Inbox, Send, Upload,
  CheckCircle2, Clock, FileText, GraduationCap, PlayCircle
} from "lucide-react";
import {
  LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, ResponsiveContainer, BarChart, Bar, Cell
} from "recharts";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════
   계정: foreon1 ~ foreon12 · 비밀번호 전부 "1234"
   부서 배정 (필요하면 여기 숫자만 바꾸면 됩니다)
   ═══════════════════════════════════════════════════════════ */
const DEPT_OF = {
  foreon1: "사무실", foreon2: "사무실", foreon3: "사무실", foreon4: "사무실", foreon5: "사무실",
  foreon6: "데스크", foreon7: "데스크", foreon8: "데스크", foreon9: "데스크",
  foreon10: "골프장", foreon11: "골프장",
  foreon12: "피트니스",
};
const DEFAULT_PW = "1234";
const AI_ADMIN = "foreon4"; // AI 정리 · 매출 분석 사용 가능한 계정

const 민원카테고리 = ["시설 고장", "이용 안내", "회원 접수", "결제·환불", "기타"];
const 담당부서 = ["데스크", "피트니스", "골프장"];

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
  방송: ["2단지 방송댄스(화목19시/방채빈)"],
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

const today = () => new Date().toISOString().slice(0, 10);
const nowLocal = () => {
  const d = new Date(); d.setMinutes(d.getMinutes() - d.getTimezoneOffset());
  return d.toISOString().slice(0, 16);
};
const won = (n) => n.toLocaleString("ko-KR") + "원";

export default function App() {
  const [me, setMe] = useState(null); // { id, dept }
  const [tab, setTab] = useState("worklog");

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

  const [complaints, setComplaints] = useState([
    { id: 1, time: "2026-08-13 14:20", cat: "시설 고장", dong: "208", ho: "1504", name: "임정윤", phone: "", content: "여자 탈의실 온수 미지근함", action: "시설팀 점검 요청, 밸브 조정 후 정상화 안내", owner: "데스크", status: "진행 중", by: "foreon6", dept: "데스크", date: "2026-08-13" },
    { id: 2, time: "2026-08-13 10:05", cat: "결제·환불", dong: "210", ho: "2506", name: "이건혁", phone: "", content: "스크린골프 당일 취소 환불 문의", action: "1시간 전까지 앱 취소 시 전액 환불 안내", owner: "데스크", status: "완료", by: "foreon7", dept: "데스크", date: "2026-08-13" },
  ]);
  const [handovers, setHandovers] = useState([
    { id: 1, type: "인계 메모", text: "내일 오전 스크린골프 3번기 A/S 방문 예정 (10시)", by: "foreon10", dept: "골프장", at: "2026-08-13 21:40" },
  ]);
  const [requests, setRequests] = useState([
    { id: 1, text: "데스크 프린터 토너 교체 요청합니다.", officeOnly: true, by: "foreon6", dept: "데스크", date: "2026-08-13", replies: [] },
  ]);
  const [events, setEvents] = useState([
    { id: 1, title: "8월 회원 정기 소독 안내", date: "2026-08-20", text: "8/20(수) 오전 6~8시 커뮤니티 전 시설 소독으로 이용이 제한됩니다.", file: null, by: "foreon1" },
  ]);
  const [lost, setLost] = useState([
    { id: 1, item: "검정 무선 이어폰", place: "헬스장", found: "2026-08-11", by: "foreon6", status: "보관중" },
  ]);

  if (!me) return <Login onLogin={setMe} />;
  const isAI = me.id === AI_ADMIN;

  const tabs = [
    { id: "worklog", label: "업무일지", icon: ClipboardList },
    { id: "search", label: "기록 검색", icon: Search },
    { id: "lost", label: "분실물", icon: PackageSearch },
    { id: "events", label: "주요 행사·일정", icon: CalendarDays },
    { id: "lessons", label: "강습 등록·취소", icon: GraduationCap },
    ...(isAI ? [{ id: "ai", label: "AI 카톡 정리", icon: Sparkles }, { id: "sales", label: "매출 분석", icon: BarChart3 }] : []),
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
            <button onClick={() => setMe(null)} className="text-slate-400 hover:text-slate-600 p-1.5"><LogOut size={16} /></button>
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

        {tab === "worklog" && <WorkLog me={me} complaints={complaints} setComplaints={setComplaints}
          handovers={handovers} setHandovers={setHandovers} requests={requests} setRequests={setRequests} />}
        {tab === "search" && <RecordSearch complaints={complaints} />}
        {tab === "lost" && <Lost me={me} rows={lost} setRows={setLost} />}
        {tab === "events" && <Events me={me} rows={events} setRows={setEvents} />}
        {tab === "lessons" && <Lessons me={me} />}
        {tab === "ai" && isAI && <TodayAI me={me} onSave={(items) => setComplaints((c) => [...items, ...c])} />}
        {tab === "sales" && isAI && <Sales />}
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
    if (DEPT_OF[key] && pw === DEFAULT_PW) { setErr(""); onLogin({ id: key, dept: DEPT_OF[key] }); }
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
    setRows([{ id: Date.now(), ...v, time: v.time.replace("T", " "), by: me.id, dept: me.dept, date: today() }, ...rows]);
    setV(blank); setOpen(false);
  };
  const toggle = (id) => setRows(rows.map((r) => r.id === id ? { ...r, status: r.status === "완료" ? "진행 중" : "완료" } : r));
  return (
    <div>
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-base font-bold">민원 대장 <span className="text-slate-400 font-normal text-sm">({rows.length})</span></h2>
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
      <div className="space-y-2">{rows.map((r) => <ComplaintCard key={r.id} r={r} onToggle={() => toggle(r.id)} />)}</div>
      <style>{`.in{width:100%;border:1px solid #e2e8f0;border-radius:8px;padding:8px 10px;font-size:13px;outline:none;background:#fff}
        .in:focus{border-color:#10b981;box-shadow:0 0 0 3px #d1fae5}`}</style>
    </div>
  );
}
function ComplaintCard({ r, onToggle }) {
  const done = r.status === "완료";
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
        </div>
        <button onClick={onToggle}
          className={`shrink-0 flex items-center gap-1 rounded-full px-2.5 py-1 text-xs font-medium ${done ? "bg-emerald-100 text-emerald-700" : "bg-amber-100 text-amber-700"}`}>
          {done ? <CheckCircle2 size={12} /> : <Clock size={12} />} {r.status}
        </button>
      </div>
    </div>
  );
}

/* 인수인계 — 같은 부서만 (사무실은 전체) */
function Handover({ me, rows, setRows, complaints }) {
  const [text, setText] = useState("");
  const canSee = (dept) => me.dept === "사무실" || me.dept === dept;
  const visible = rows.filter((r) => canSee(r.dept));
  const pending = complaints.filter((c) => c.status === "진행 중");
  const add = () => {
    if (!text.trim()) return;
    setRows([{ id: Date.now(), type: "인계 메모", text, by: me.id, dept: me.dept, at: nowLocal().replace("T", " ") }, ...rows]);
    setText("");
  };
  return (
    <div className="space-y-4">
      <div className="bg-amber-50 border border-amber-200 rounded-xl p-4">
        <p className="text-sm font-semibold text-amber-800 flex items-center gap-1.5 mb-2"><Clock size={15} /> 이전 근무 이관 건 (진행 중 민원 {pending.length})</p>
        {pending.length ? <div className="space-y-1.5">{pending.map((c) => (
          <div key={c.id} className="text-sm text-slate-700 bg-white rounded-lg px-3 py-2 border border-amber-100">
            <b>{c.cat}</b> · {c.content} <span className="text-xs text-slate-400 font-mono">({c.by})</span>
          </div>))}</div> : <p className="text-sm text-amber-700">인계할 진행 중 건이 없습니다.</p>}
      </div>
      <div className="bg-white rounded-xl border border-slate-200 p-4">
        <p className="text-xs font-semibold text-slate-500 mb-2">인계 메모 — 오늘 처리 못한 건, 공지, 익일 예약/행사 안내 등을 자유롭게 남기세요</p>
        <textarea rows={3} value={text} onChange={(e) => setText(e.target.value)} placeholder="다음 근무자에게 전달할 내용을 적으세요"
          className="w-full rounded-lg border border-slate-200 p-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 resize-none" />
        <button onClick={add} className="mt-2 flex items-center gap-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2"><Send size={14} /> 등록</button>
      </div>
      <div className="space-y-2">
        {visible.map((r) => (
          <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm">
            <div className="flex items-center gap-2 mb-1">
              <span className="text-xs px-2 py-0.5 rounded bg-sky-100 text-sky-700">{r.type}</span>
              <span className="text-xs text-slate-400 font-mono">{r.dept}</span>
            </div>
            <p className="text-sm">{r.text}</p>
            <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.at} · {r.by}</p>
          </div>
        ))}
        {!visible.length && <p className="text-center text-slate-400 text-sm py-6">표시할 인수인계가 없습니다.</p>}
      </div>
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
    setRows([{ id: Date.now(), text, officeOnly, by: me.id, dept: me.dept, date: today(), replies: [] }, ...rows]);
    setText(""); setOfficeOnly(true);
  };
  const reply = (id, msg) => setRows(rows.map((r) => r.id === id ? { ...r, replies: [...r.replies, { text: msg, by: me.id, date: today() }] } : r));
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
        {visible.map((r) => <RequestCard key={r.id} r={r} canReply={office} onReply={(m) => reply(r.id, m)} />)}
        {!visible.length && <p className="text-center text-slate-400 text-sm py-6">요청이 없습니다.</p>}
      </div>
    </div>
  );
}
function RequestCard({ r, canReply, onReply }) {
  const [msg, setMsg] = useState("");
  return (
    <div className="bg-white rounded-xl border border-slate-200 p-4 shadow-sm">
      <div className="flex items-center gap-2 mb-1">
        {r.officeOnly && <span className="text-[11px] px-1.5 py-0.5 rounded bg-slate-100 text-slate-500">🔒 사무실만</span>}
      </div>
      <p className="text-sm">{r.text}</p>
      <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.date} · {r.by} ({r.dept})</p>
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
  const add = () => { if (!(v.item || "").trim()) return; setRows([{ id: Date.now(), item: v.item, place: v.place || "", found: today(), by: me.id, status: "보관중" }, ...rows]); setV({}); setOpen(false); };
  const toggle = (id) => setRows(rows.map((r) => r.id === id ? { ...r, status: r.status === "보관중" ? "반환완료" : "보관중", by: me.id } : r));
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
          <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm flex items-start justify-between gap-3">
            <div className="min-w-0"><p className="font-medium truncate">{r.item}</p>
              <p className="text-xs text-slate-500 mt-0.5">{r.place} · 습득 <span className="font-mono">{r.found}</span></p>
              <p className="text-xs text-emerald-600 mt-1">👤 {r.by}</p></div>
            <button onClick={() => toggle(r.id)} className={`shrink-0 rounded-full px-2.5 py-1 text-xs font-medium ${r.status === "보관중" ? "bg-amber-100 text-amber-700" : "bg-emerald-100 text-emerald-700"}`}>{r.status}</button>
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
  const onFile = (e) => {
    const f = e.target.files?.[0]; if (!f) return;
    const reader = new FileReader();
    reader.onload = (ev) => setV((s) => ({ ...s, file: { name: f.name, type: f.type, url: ev.target.result } }));
    reader.readAsDataURL(f);
  };
  const add = () => { if (!v.title.trim()) return; setRows([{ id: Date.now(), ...v, by: me.id }, ...rows]); setV({ title: "", date: today(), text: "", file: null }); setOpen(false); };
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
            <div className="flex items-center gap-2 mb-1">
              <CalendarDays size={15} className="text-emerald-600" />
              <span className="font-semibold">{r.title}</span>
              <span className="text-xs text-slate-400 font-mono ml-auto">{r.date}</span>
            </div>
            {r.text && <p className="text-sm text-slate-600 mt-1 whitespace-pre-wrap">{r.text}</p>}
            {r.file && (r.file.type.includes("image")
              ? <img src={r.file.url} alt="" className="mt-2 rounded-lg border border-slate-200 max-h-56" />
              : <a href={r.file.url} download={r.file.name} className="mt-2 inline-flex items-center gap-1.5 text-sm text-sky-600 bg-sky-50 rounded-lg px-3 py-1.5"><FileText size={14} /> {r.file.name}</a>)}
            <p className="text-[11px] text-slate-400 mt-2 font-mono">등록 · {r.by}</p>
          </div>
        ))}
      </div>
    </div>
  );
}

/* ─────────────────────── 강습 등록·취소 */
function Lessons({ me }) {
  const [cat, setCat] = useState(LESSON_CATS[0]);
  const [course, setCourse] = useState("");
  const [action, setAction] = useState("등록"); // 등록 | 취소
  const [dong, setDong] = useState("");
  const [ho, setHo] = useState("");
  const [name, setName] = useState("");
  const [log, setLog] = useState([]);
  const [running, setRunning] = useState(false);

  const selectCat = (c) => { setCat(c); setCourse(""); };

  const run = () => {
    if (!course) { alert("강좌를 먼저 선택하세요."); return; }
    if (!dong.trim() || !ho.trim() || !name.trim()) { alert("동·호수·이름을 모두 입력하세요."); return; }
    setRunning(true);
    // TODO: 여기가 파이썬 프로그램 연동 지점입니다.
    // 지금은 백엔드가 없어서 실제로 실행되지는 않고, 요청 내용만 기록됩니다.
    // 파이썬 코드를 주시면 이 부분을 서버(API 라우트)로 연결해 진짜 실행되게 만들겠습니다.
    setTimeout(() => {
      setLog([{ id: Date.now(), cat, course, action, dong, ho, name, by: me.id, at: nowLocal().replace("T", " ") }, ...log]);
      setRunning(false);
      setDong(""); setHo(""); setName("");
    }, 500);
  };

  return (
    <div>
      <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 mb-4 text-xs text-amber-800">
        ⚠ 지금은 화면(폼)까지만 완성된 상태예요. "실행"을 누르면 요청 내용이 아래 목록에 기록만 되고,
        실제 BYB 등록·취소 처리는 파이썬 코드를 서버에 연결한 뒤부터 동작합니다.
      </div>

      {/* 분류 탭 */}
      <div className="flex gap-1.5 mb-4 overflow-x-auto pb-1">
        {LESSON_CATS.map((c) => (
          <button key={c} onClick={() => selectCat(c)}
            className={`shrink-0 text-sm font-medium px-3.5 py-1.5 rounded-full transition ${cat === c ? "bg-emerald-600 text-white" : "bg-white border border-slate-200 text-slate-500"}`}>
            {c}
          </button>
        ))}
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-4">
        {/* 등록/취소 토글 */}
        <div>
          <span className="text-xs font-medium text-slate-500 block mb-1.5">처리 구분</span>
          <div className="inline-flex p-1 bg-slate-100 rounded-lg">
            {["등록", "취소"].map((a) => (
              <button key={a} onClick={() => setAction(a)}
                className={`px-4 py-1.5 rounded-md text-sm font-medium transition ${action === a ? (a === "등록" ? "bg-emerald-600 text-white" : "bg-rose-500 text-white") : "text-slate-500"}`}>
                {a}
              </button>
            ))}
          </div>
        </div>

        {/* 강좌 선택 */}
        <label className="block">
          <span className="text-xs font-medium text-slate-500 block mb-1.5">강좌 선택 · {cat} ({COURSES[cat].length}개)</span>
          <select value={course} onChange={(e) => setCourse(e.target.value)}
            className="w-full rounded-lg border border-slate-200 px-3 py-2.5 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 bg-white">
            <option value="">— 강좌를 선택하세요 —</option>
            {COURSES[cat].map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </label>

        {/* 대상자 정보 */}
        <div className="grid grid-cols-3 gap-2">
          <label className="block"><span className="text-xs font-medium text-slate-500 block mb-1.5">동</span>
            <input value={dong} onChange={(e) => setDong(e.target.value)} placeholder="예: 208"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" /></label>
          <label className="block"><span className="text-xs font-medium text-slate-500 block mb-1.5">호수</span>
            <input value={ho} onChange={(e) => setHo(e.target.value)} placeholder="예: 1504"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" /></label>
          <label className="block"><span className="text-xs font-medium text-slate-500 block mb-1.5">이름</span>
            <input value={name} onChange={(e) => setName(e.target.value)} placeholder="회원 성함"
              className="w-full rounded-lg border border-slate-200 px-3 py-2 text-sm outline-none focus:border-emerald-500" /></label>
        </div>

        <button onClick={run} disabled={running}
          className={`w-full flex items-center justify-center gap-2 rounded-lg text-white text-sm font-medium py-2.5 transition disabled:opacity-50
            ${action === "등록" ? "bg-emerald-600 hover:bg-emerald-700" : "bg-rose-600 hover:bg-rose-700"}`}>
          {running ? <><Loader2 size={16} className="animate-spin" /> 처리 중…</> : <><PlayCircle size={16} /> 실행 ({action})</>}
        </button>
      </div>

      {log.length > 0 && (
        <div className="mt-4">
          <p className="text-sm font-semibold text-slate-700 mb-2">처리 요청 기록</p>
          <div className="space-y-2">
            {log.map((r) => (
              <div key={r.id} className="bg-white rounded-xl border border-slate-200 p-3.5 shadow-sm">
                <div className="flex items-center gap-2">
                  <span className={`text-xs font-medium px-2 py-0.5 rounded-full ${r.action === "등록" ? "bg-emerald-100 text-emerald-700" : "bg-rose-100 text-rose-700"}`}>{r.action}</span>
                  <span className="text-xs text-slate-400">{r.cat}</span>
                </div>
                <p className="text-sm font-medium mt-1">{r.course}</p>
                <p className="text-xs text-slate-500 mt-0.5">{r.dong}동 {r.ho}호 · {r.name}</p>
                <p className="text-[11px] text-slate-400 mt-1.5 font-mono">{r.at} · {r.by}</p>
              </div>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}


function TodayAI({ me, onSave }) {
  const [text, setText] = useState("");
  const [loading, setLoading] = useState(false);
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");
  const run = async () => {
    if (!text.trim()) return; setLoading(true); setErr(""); setResult(null);
    const prompt = `아래는 아파트 커뮤니티 데스크 직원들의 하루치 카카오톡 대화이다. 주요 민원/문의를 뽑아 정리하라.
반드시 아래 형식의 순수 JSON 배열로만 응답한다(설명·마크다운·코드블록 금지):
[ { "cat": "시설 고장"|"이용 안내"|"회원 접수"|"결제·환불"|"기타", "content": "핵심 내용 1~2문장", "action": "조치/대응 (없으면 빈 문자열)" } ]
카카오톡:
"""${text}"""`;
    try {
      const res = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST", headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "claude-sonnet-4-6", max_tokens: 1000, messages: [{ role: "user", content: prompt }] }),
      });
      const data = await res.json();
      const raw = data.content.map((c) => (c.type === "text" ? c.text : "")).join("");
      const items = JSON.parse(raw.replace(/```json|```/g, "").trim());
      setResult(items.map((it, i) => ({ id: Date.now() + i, time: nowLocal().replace("T", " "), cat: it.cat || "기타", dong: "", ho: "", name: "", phone: "", content: it.content, action: it.action || "", owner: "데스크", status: "진행 중", by: me.id, dept: me.dept, date: today() })));
    } catch (e) { setErr("정리에 실패했어요. 내용을 줄이거나 다시 시도해 주세요."); }
    finally { setLoading(false); }
  };
  const save = () => { onSave(result); setResult(null); setText(""); };
  return (
    <div>
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5">
        <div className="flex items-center gap-2 mb-1"><Sparkles size={18} className="text-emerald-600" /><h2 className="font-bold">AI 카톡 정리</h2></div>
        <p className="text-sm text-slate-500 mb-3">카카오톡 대화를 붙여넣으면 민원으로 자동 분류해 민원 대장·기록 검색에 저장합니다.</p>
        <textarea value={text} onChange={(e) => setText(e.target.value)} rows={7} placeholder="카카오톡 대화 내용을 붙여넣으세요…"
          className="w-full rounded-xl border border-slate-200 p-3 text-sm outline-none focus:border-emerald-500 focus:ring-2 focus:ring-emerald-100 resize-none" />
        <button onClick={run} disabled={loading || !text.trim()} className="mt-3 flex items-center justify-center gap-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 text-white text-sm font-medium px-4 py-2.5 w-full sm:w-auto">
          {loading ? <><Loader2 size={16} className="animate-spin" /> 정리하는 중…</> : <><Sparkles size={16} /> AI로 정리하기</>}
        </button>
        {err && <p className="text-xs text-red-500 mt-2">{err}</p>}
      </div>
      {result && (
        <div className="mt-4">
          <div className="flex items-center justify-between mb-2"><p className="text-sm font-semibold">정리 결과 {result.length}건</p>
            <button onClick={save} className="rounded-lg bg-slate-800 text-white text-xs font-medium px-3 py-2">민원 대장에 저장</button></div>
          <div className="space-y-2">{result.map((r) => <ComplaintCard key={r.id} r={r} onToggle={() => {}} />)}</div>
        </div>
      )}
    </div>
  );
}

/* ─────────────────────── 매출 분석 (foreon4 전용) */
const FAC_COLORS = ["#059669", "#0ea5e9", "#f59e0b", "#8b5cf6", "#ec4899", "#64748b"];
function facilityGroup(name = "") {
  if (name.includes("골프")) return "골프장";
  if (name.includes("독서실")) return "독서실";
  if (name.includes("체육관") || name.includes("농구")) return "체육관";
  if (name.includes("테니스")) return "스크린테니스";
  if (name.includes("탁구")) return "탁구장";
  if (name.includes("수영")) return "수영장";
  return "기타";
}
function Sales() {
  const [data, setData] = useState(null);
  const [err, setErr] = useState("");
  const onFile = (e) => {
    const file = e.target.files?.[0]; if (!file) return; setErr("");
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const wb = XLSX.read(new Uint8Array(ev.target.result), { type: "array", cellDates: true });
        const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { defval: "" });
        if (!rows.length) { setErr("데이터가 비어 있어요."); return; }
        const k = (kw, fb) => Object.keys(rows[0]).find((x) => x.includes(kw)) || fb;
        const dK = k("결제", "결제일시"), fK = k("요금", "요금"), gK = k("대상", "대상");
        const byDate = {}, byFac = {}; let total = 0, cancels = 0;
        rows.forEach((r) => {
          const raw = r[dK]; const d = raw instanceof Date ? raw.toISOString().slice(0, 10) : String(raw).slice(0, 10);
          const fee = Number(String(r[fK]).replace(/,/g, "")) || 0; if (!d) return;
          byDate[d] = (byDate[d] || 0) + fee;
          const g = facilityGroup(String(r[gK])); byFac[g] = (byFac[g] || 0) + fee;
          total += fee; if (fee < 0) cancels++;
        });
        setData({
          trend: Object.entries(byDate).sort().map(([date, revenue]) => ({ date: date.slice(5), revenue })),
          facs: Object.entries(byFac).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value),
          total, count: rows.length, cancels,
        });
      } catch { setErr("엑셀을 읽지 못했어요. BYB 내보내기 파일이 맞는지 확인해 주세요."); }
    };
    reader.readAsArrayBuffer(file);
  };
  return (
    <div>
      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 mb-4">
        <div className="flex items-center gap-2 mb-1"><BarChart3 size={18} className="text-emerald-600" /><h2 className="font-bold">매출 분석</h2></div>
        <p className="text-sm text-slate-500 mb-3">BYB 예약·정산 엑셀을 올리면 날짜별·시설별 매출을 보여줍니다.</p>
        <label className="inline-flex items-center gap-2 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-medium px-4 py-2.5 cursor-pointer"><Upload size={16} /> 엑셀 파일 선택
          <input type="file" accept=".xlsx,.xls" onChange={onFile} className="hidden" /></label>
        {err && <p className="text-xs text-red-500 mt-2">{err}</p>}
      </div>
      {data && (
        <>
          <div className="grid grid-cols-3 gap-2 mb-4">
            <Stat label="순매출" value={won(data.total)} accent />
            <Stat label="예약 건수" value={data.count.toLocaleString() + "건"} />
            <Stat label="취소 건수" value={data.cancels + "건"} />
          </div>
          <Panel title="날짜별 매출 추이">
            <ResponsiveContainer width="100%" height={220}>
              <LineChart data={data.trend} margin={{ top: 10, right: 10, left: -10, bottom: 0 }}>
                <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" />
                <XAxis dataKey="date" tick={{ fontSize: 11 }} stroke="#94a3b8" />
                <YAxis tick={{ fontSize: 11 }} stroke="#94a3b8" tickFormatter={(v) => v / 1000 + "k"} />
                <Tooltip formatter={(v) => won(v)} />
                <Line type="monotone" dataKey="revenue" stroke="#059669" strokeWidth={2.5} dot={{ r: 3 }} />
              </LineChart>
            </ResponsiveContainer>
          </Panel>
          <Panel title="시설별 매출">
            <ResponsiveContainer width="100%" height={Math.max(160, data.facs.length * 42)}>
              <BarChart data={data.facs} layout="vertical" margin={{ top: 0, right: 20, left: 10, bottom: 0 }}>
                <XAxis type="number" hide /><YAxis type="category" dataKey="name" width={70} tick={{ fontSize: 12 }} stroke="#64748b" />
                <Tooltip formatter={(v) => won(v)} />
                <Bar dataKey="value" radius={[0, 6, 6, 0]}>{data.facs.map((_, i) => <Cell key={i} fill={FAC_COLORS[i % FAC_COLORS.length]} />)}</Bar>
              </BarChart>
            </ResponsiveContainer>
          </Panel>
        </>
      )}
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
