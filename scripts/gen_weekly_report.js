const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── data.js 공유 데이터 연동 ────────────────────────────────
const { tasks, milestones, todayWeek: DATA_WEEK } = require("../docs/data.js");

// 진척률 자동 계산 (tasks 배열 기반)
const _all    = tasks.filter(t => !t.cat);
const _done   = _all.filter(t => t.st === "done");
const _prog   = _all.filter(t => t.st === "progress");
const _todo   = _all.filter(t => t.st === "todo");
const _rate   = _all.length > 0 ? _done.length / _all.length : 0;
const _planned= _all.filter(t => t.e <= DATA_WEEK).length;
const _planR  = _all.length > 0 ? _planned / _all.length : 0;
const _gap    = _rate - _planR;

// 트랙별 집계
const _td = {};
["기획","ACS","UI","PLC","통합"].forEach(tr => {
  const t = _all.filter(t => t.track === tr);
  _td[tr] = { done: t.filter(t => t.st === "done").length, total: t.length };
});

// 지연 항목: 종료주 지난 미완료
const _delayed = _all.filter(t => t.e < DATA_WEEK && t.st !== "done");
// 다음 주 착수 예정: 시작주가 이번~다음 주인 todo
const _upcoming = _all.filter(t => t.s >= DATA_WEEK && t.s <= DATA_WEEK + 1 && t.st === "todo");

// ── 커밋 → task ID 매핑 헬퍼 ─────────────────────────────────
/**
 * 커밋 메시지 배열을 task ID 기준으로 분류
 * 1순위: 커밋 메시지에 ID 패턴 (A-01, P-05 등) 직접 언급
 * 2순위: task 이름 키워드 포함 여부
 * 미매핑: "기타"
 */
function groupByTaskId(commits, taskList) {
  const grouped = {};
  commits.forEach(msg => {
    // 1) ID 직접 언급
    const idMatch = msg.match(/\b([APLUT]-\d+[a-z]?)\b/i);
    if (idMatch) {
      const id = idMatch[1].toUpperCase();
      const task = taskList.find(t => t.id && t.id.toUpperCase() === id);
      if (task) {
        if (!grouped[id]) grouped[id] = { name: task.name, commits: [] };
        grouped[id].commits.push(msg);
        return;
      }
    }
    // 2) task 이름 키워드 매칭
    let hit = null;
    for (const task of taskList) {
      if (!task.name) continue;
      const kws = task.name.split(/[\s\/\(\)\+·,\-]+/).filter(k => k.length >= 4);
      if (kws.some(kw => msg.toLowerCase().includes(kw.toLowerCase()))) { hit = task; break; }
    }
    if (hit) {
      const id = hit.id;
      if (!grouped[id]) grouped[id] = { name: hit.name, commits: [] };
      grouped[id].commits.push(msg);
    } else {
      if (!grouped["기타"]) grouped["기타"] = { name: "기타", commits: [] };
      grouped["기타"].commits.push(msg);
    }
  });
  return grouped;
}

/** groupByTaskId 결과를 PptxGenJS bullet 배열로 변환 */
function commitsToBullets(grouped, fs=9) {
  const out = [];
  Object.entries(grouped).forEach(([id, info]) => {
    if (id === "기타") return;
    out.push({ text:`[${id}] ${info.name}`, options:{bold:true,breakLine:true,fontSize:fs+0.5,color:NB} });
    info.commits.forEach(c => out.push({ text:`  ${c}`, options:{bullet:true,breakLine:true,fontSize:fs,color:NDARK} }));
  });
  if (grouped["기타"]?.commits?.length) {
    out.push({ text:"기타", options:{bold:true,breakLine:true,fontSize:fs,color:NDGR} });
    grouped["기타"].commits.forEach(c => out.push({ text:`  ${c}`, options:{bullet:true,breakLine:true,fontSize:fs-0.5,color:NDGR} }));
  }
  return out;
}

const OUTPUT_PM = "/home/sally/claude_cowork/HACS2.0_주간보고_11주차.pptx"; // v4 — 기획 행 공란
const TODAY = "2026.03.14";
const PERIOD = "2026년 3월 2주차 (3/9~3/13)";
const PERIOD_SHORT = "3/9~3/13";

// ── Navifra 브랜드 컬러 ──────────────────────────────
const NN   = "1A2744"; // Navifra Navy (primary dark)
const NB   = "2563EB"; // Navifra Blue (primary)
const NB2  = "1E3A6E"; // Navifra Blue Dark (sub-header)
const NBL  = "EBF3FE"; // Navifra Blue Light (fill)
const NW   = "FFFFFF";
const NGR  = "F5F7FA"; // Light gray bg
const NDGR = "64748B"; // Dark gray text
const NRED = "DC2626";
const NGN  = "059669";
const NAMB = "D97706";
const NTEAL= "0D9488";
const NPUR = "7C3AED";
const NDARK= "1E293B";

const shade = () => ({ type:"outer", blur:4, offset:2, angle:135, color:"000000", opacity:0.1 });

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author  = "나비프라 PM";
pres.title   = "H-ACS 2.0 주간 개발 현황 3월 2주차";

function hdr(slide, title, sub) {
  slide.background = { color: NGR };
  slide.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:0.8, fill:{color:NN} });
  slide.addText(title, { x:0.7,y:0,w:5.5,h:0.8, fontSize:20,fontFace:"Calibri",color:NW,bold:true,valign:"middle",margin:0 });
  if (sub) slide.addText(sub, { x:4.5,y:0,w:5.2,h:0.8, fontSize:11,fontFace:"Calibri",color:"93B5D8",align:"right",valign:"middle",margin:0 });
}

// ══════════════════════════════════════════════
// SLIDE 1 : 표지
// ══════════════════════════════════════════════
let s1 = pres.addSlide();
s1.background = { color: NN };
s1.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:0.05, fill:{color:NB} });
s1.addShape(pres.shapes.RECTANGLE, { x:0.7,y:0.6,w:2.0,h:0.42, fill:{color:NB} });
s1.addText("NAVIFRA", { x:0.7,y:0.6,w:2.0,h:0.42, fontSize:15,fontFace:"Calibri",color:NW,bold:true,align:"center",valign:"middle",margin:0 });
s1.addText("H-ACS 2.0", { x:0.7,y:1.65,w:8.6,h:0.85, fontSize:42,fontFace:"Calibri",color:NW,bold:true,margin:0 });
s1.addText("주간 개발 현황", { x:0.7,y:2.4,w:8.6,h:0.65, fontSize:28,fontFace:"Calibri",color:"93B5E8",margin:0 });
s1.addShape(pres.shapes.RECTANGLE, { x:0.7,y:3.2,w:2.2,h:0.03, fill:{color:NB} });
s1.addText(PERIOD, { x:0.7,y:3.42,w:7,h:0.38, fontSize:16,fontFace:"Calibri",color:"93B5E8",margin:0 });
s1.addText("현대자동차 H-ACS 2.0 개발 고도화 프로젝트", { x:0.7,y:4.58,w:7,h:0.28, fontSize:11,fontFace:"Calibri",color:"7A99B8",margin:0 });
s1.addText("나비프라  |  GitHub 자동 수집 기반", { x:0.7,y:4.86,w:5,h:0.28, fontSize:10,fontFace:"Calibri",color:"5A7998",margin:0 });
s1.addText(TODAY, { x:7.5,y:4.86,w:2,h:0.28, fontSize:10,fontFace:"Calibri",color:"5A7998",align:"right",margin:0 });

// ══════════════════════════════════════════════
// SLIDE 2 : 개발 진행 현황 표 (핵심 슬라이드)
// ══════════════════════════════════════════════
let s2 = pres.addSlide();
s2.background = { color: NW };

// 타이틀 영역
s2.addShape(pres.shapes.RECTANGLE, { x:0,y:0,w:10,h:0.55, fill:{color:NN} });
s2.addShape(pres.shapes.RECTANGLE, { x:0,y:0.55,w:10,h:0.06, fill:{color:NB} });
s2.addText("■ H-ACS 2.0 개발 진행 현황 (26년 3월 2주차)", {
  x:0.3,y:0,w:7.5,h:0.55, fontSize:17,fontFace:"맑은 고딕",color:NW,bold:true,valign:"middle",margin:0
});
// RESTRICTED 뱃지
s2.addShape(pres.shapes.RECTANGLE, { x:8.3,y:0.06,w:1.5,h:0.42, fill:{color:NW}, line:{color:NRED,pt:1.5} });
s2.addText("사내한", { x:8.3,y:0.06,w:1.5,h:0.22, fontSize:8,fontFace:"맑은 고딕",color:NRED,bold:true,align:"center",valign:"middle",margin:0 });
s2.addText("RESTRICTED", { x:8.3,y:0.26,w:1.5,h:0.22, fontSize:8,fontFace:"맑은 고딕",color:NRED,bold:true,align:"center",valign:"middle",margin:0 });

// 공통 셀 옵션
const cell = (extra={}) => ({
  fontFace:"맑은 고딕", fontSize:9, border:{pt:0.5,color:"C8D8EE"},
  valign:"top", ...extra
});
const hc = (extra={}) => cell({ bold:true, align:"center", valign:"middle", color:NW, fill:{color:NN}, fontSize:9.5, ...extra });
const sh = (extra={}) => cell({ bold:true, align:"center", valign:"middle", color:NW, fill:{color:NB2}, fontSize:9, ...extra });
const cat = (extra={}) => cell({ bold:true, align:"center", valign:"middle", color:NN, fill:{color:NBL}, fontSize:10, ...extra });
const sub = (extra={}) => cell({ bold:true, align:"center", valign:"middle", color:NN, fill:{color:"EEF4FD"}, fontSize:9, ...extra });
const con = (extra={}) => cell({ color:NDARK, fill:{color:NW}, fontSize:9, margin:[3,5,3,5], ...extra });
const altcon=(extra={})=> cell({ color:NDARK, fill:{color:"F8FAFC"}, fontSize:9, margin:[3,5,3,5], ...extra });

// 텍스트 데이터
const acs_실적 =
  "1. H-ACS Map Service 구현\n" +
  "  a. 지도/영역/포털/노드 메시지 처리 및 저장\n" +
  "  b. Config Adapter, MQTT Subscribe Callback 완성\n" +
  "  c. HTTP Token 인증 처리, Data Store Key 템플릿화\n" +
  "2. SDD v2 기반 클래스 구조 구현 (41개 클래스 수정)\n" +
  "  - MAP_ENDPOINT / PROTOCOL_TYPE / PROTOCOL Enum 신규\n" +
  "3. Linux 빌드 환경 개선 (configure 캐시 fresh 초기화)";
const acs_계획 =
  "1. VDA5050 MQTT 통신 기반 완료 (A-02)\n" +
  "2. Backend API 연동 완료 (A-03)\n" +
  "3. 로봇 시뮬레이터 1대 단순 경로 주행 (A-04)\n" +
  "4. 경로 생성 알고리즘(A*/Dijkstra) 완료 (A-05)\n" +
  "5. SDD 클래스 다이어그램 1차 완성 (P-05)";

const ui_실적 =
  "1. Audit Log 기능 구현 (Frontend+Backend) #707/#606\n" +
  "2. API Swagger 문서화 완료 (JSDoc 기반) #657\n" +
  "3. Brain Desktop 프로그램 통합 첫 빌드 #818\n" +
  "4. UX 개선 다수\n" +
  "  a. 다크 테마 배경 수정 #812\n" +
  "  b. 맵 리스트 Text Truncate #813\n" +
  "  c. AppBar Layout CSS 버그 수정 #809\n" +
  "5. 404 다국어, Favorite 아이콘, Package 업그레이드";
const ui_계획 =
  "1. 노드/맵 구성 요소 편집 완료 (U-01)\n" +
  "2. 로봇 객체 그룹 관리 UI 완료 (U-03)\n" +
  "3. PLC 객체 DB 관리 UI 착수 (U-04)\n" +
  "4. Swagger API 문서 정비 지속";

const gi_실적 = "";
const gi_계획 = "";

const rows = [
  // Row 0 — 헤더 (rowspan 포함, 진행현황 colspan=2)
  [
    { text:"과제명",   options: hc({ rowspan:2 }) },
    { text:"구분",     options: hc({ rowspan:2 }) },
    { text:"당당",     options: hc({ rowspan:2 }) },
    { text:"진행 현황",options: hc({ colspan:2  }) },
  ],
  // Row 1 — 서브헤더
  [
    { text:"실적 (3/9~3/13)",  options: sh() },
    { text:"계획 (3/16~3/20)", options: sh() },
  ],
  // Row 2 — ACS  (H-ACS 2.0 rowspan=3)
  [
    { text:"H-ACS 2.0", options: cat({ rowspan:3, fontSize:12 }) },
    { text:"ACS",       options: sub() },
    { text:"권익현\n최명서",    options: sub({ fontSize:8.5 }) },
    { text: acs_실적,   options: con() },
    { text: acs_계획,   options: con({ fill:{color:NBL} }) },
  ],
  // Row 3 — UX/UI
  [
    { text:"UX/UI",  options: sub() },
    { text:"종원선\n박새롬",    options: sub({ fontSize:8.5 }) },
    { text: ui_실적,  options: altcon() },
    { text: ui_계획,  options: altcon({ fill:{color:NBL} }) },
  ],
  // Row 4 — 기획
  [
    { text:"기획",   options: sub() },
    { text:"최준섭",           options: sub({ fontSize:8.5 }) },
    { text: gi_실적,  options: con() },
    { text: gi_계획,  options: con({ fill:{color:NBL} }) },
  ],
  // Row 5 — 주요 공지
  [
    { text:"주요 공지", options: sub({ colspan:2, valign:"middle" }) },
    { text:"당당",     options: sub({ fontSize:8.5, valign:"middle" }) },
    { text:"- P-02 시스템 아키텍처 설계 공식 완료 처리 필요 (종료주 W8 초과 진행 중)\n- Navifra-Garry 백엔드 기여분 H-ACS 이슈 라벨 적용 요청",
      options: con({ colspan:2, fontSize:8.5, valign:"middle" }) },
  ],
];

s2.addTable(rows, {
  x: 0.25, y: 0.68,
  w: 9.5,
  colW: [0.82, 0.78, 0.78, 3.56, 3.56],
  rowH: [0.3, 0.28, 1.1, 1.3, 1.05, 0.46],
  border: { pt:0.5, color:"C8D8EE" },
  autoPage: false,
});

// ══════════════════════════════════════════════
// SLIDE 3 : 주간 요약
// ══════════════════════════════════════════════
let s3 = pres.addSlide();
hdr(s3, "주간 요약", `3월 2주차 (${PERIOD_SHORT})`);

const kpis = [
  { label:"총 커밋",    value:"35건",  sub:"ACS 12 + UI/UX 23",          color:NB },
  { label:"신규 문서",  value:"2건",   sub:"SDD v2 + DB 설계서 v0.3",     color:NGN },
  { label:"활동 개발자",value:"4명",   sub:"익현·명서·숀·종원선",          color:NPUR },
  { label:"클래스 변경",value:"41개",  sub:"수정 41 + 신규 Enum 3",       color:NAMB },
];
kpis.forEach((k,i)=>{
  const x = 0.5 + i*2.35;
  s3.addShape(pres.shapes.RECTANGLE, {x,y:1.1,w:2.15,h:1.38,fill:{color:NW},shadow:shade()});
  s3.addShape(pres.shapes.RECTANGLE, {x,y:1.1,w:2.15,h:0.06,fill:{color:k.color}});
  s3.addText(k.label, {x,y:1.22,w:2.15,h:0.28,fontSize:9.5,fontFace:"Calibri",color:NDGR,align:"center",margin:0});
  s3.addText(k.value, {x,y:1.46,w:2.15,h:0.5,fontSize:26,fontFace:"Calibri",color:NN,bold:true,align:"center",margin:0});
  s3.addText(k.sub,   {x,y:1.96,w:2.15,h:0.38,fontSize:8,fontFace:"Calibri",color:NDGR,align:"center",margin:0});
});

s3.addShape(pres.shapes.RECTANGLE, {x:0.5,y:2.68,w:9,h:2.62,fill:{color:NW},shadow:shade()});
s3.addShape(pres.shapes.RECTANGLE, {x:0.5,y:2.68,w:0.06,h:2.62,fill:{color:NB}});
s3.addText("이번 주 핵심 성과", {x:0.8,y:2.78,w:8,h:0.35,fontSize:13,fontFace:"Calibri",color:NN,bold:true,margin:0});
s3.addText([
  {text:"DB 설계서 v0.3 완성 (2026-03-13) — 14개 테이블 전체 스키마 확정",options:{bullet:true,breakLine:true,fontSize:11,color:NDARK}},
  {text:"맵/스캔맵(11), 인증/사용자(3), PLC(1), 공통코드(2), ACS 그룹(1) 도메인 커버",options:{bullet:true,indentLevel:1,breakLine:true,fontSize:9.5,color:NDGR}},
  {text:"ACS Map Service 구현 완료 — SDD v2 6계층 아키텍처 기반, cMapService→cBaseService",options:{bullet:true,breakLine:true,fontSize:11,color:NDARK}},
  {text:"MQTT Subscribe Callback, HTTP Token, Config Adapter, Data Store Key 템플릿화",options:{bullet:true,indentLevel:1,breakLine:true,fontSize:9.5,color:NDGR}},
  {text:"UI/UX — Audit Log(audit_logs 연동), Swagger 문서화, Brain Desktop 통합 #818",options:{bullet:true,breakLine:true,fontSize:11,color:NDARK}},
  {text:"SDD v2 Ticket/Mission 상태 다이어그램·시퀀스 다이어그램 완성",options:{bullet:true,fontSize:11,color:NB}},
],{x:0.9,y:3.18,w:8.3,h:2.05,fontFace:"Calibri",paraSpaceAfter:3,margin:0});

// ══════════════════════════════════════════════
// SLIDE 4 : ACS 파트 + SDD 아키텍처
// ══════════════════════════════════════════════
let s4 = pres.addSlide();
hdr(s4, "ACS 파트 상세", "navifra-dev/HACS-2.0  |  SDD v2 기반 C++");

s4.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:4.0,h:2.45,fill:{color:NW},shadow:shade()});
s4.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:4.0,h:0.06,fill:{color:NGN}});
s4.addText("권익현 (ben-navifra)",{x:0.7,y:1.08,w:3.0,h:0.3,fontSize:12,fontFace:"Calibri",color:NN,bold:true,margin:0});
s4.addText("10커밋",{x:3.75,y:1.08,w:0.65,h:0.3,fontSize:10,fontFace:"Calibri",color:NGN,bold:true,align:"right",margin:0});
// ── 커밋 목록 정의 (매주 이 부분만 수정) ──────────────────
const ikHyunCommits = [
  "[feat] map service 추가 (A-01)",
  "[fix] device/area 업데이트 시 node device index 갱신",
  "[fix] data store key 템플릿 변경 / config adapter 추가 (A-02)",
  "[fix] MQTT subscribe callback, HTTP token 추가 (A-02)",
  "[fix] geometry math util, 역직렬화 주석, 테스트 코드",
  "[docs] map service README 추가",
];
s4.addText(commitsToBullets(groupByTaskId(ikHyunCommits, _all.filter(t=>t.track==="ACS"||t.track==="기획"))),
  {x:0.7,y:1.45,w:3.65,h:1.9,fontFace:"Calibri",paraSpaceAfter:2,margin:0});

s4.addShape(pres.shapes.RECTANGLE,{x:0.5,y:3.55,w:4.0,h:0.85,fill:{color:NW},shadow:shade()});
s4.addShape(pres.shapes.RECTANGLE,{x:0.5,y:3.55,w:4.0,h:0.06,fill:{color:NTEAL}});
s4.addText("최명서 (Navifra-Chris) — 2커밋",{x:0.7,y:3.68,w:3.7,h:0.28,fontSize:11,fontFace:"Calibri",color:NN,bold:true,margin:0});
const chrisCommits = [
  "[chore] output/tmp 무시 규칙",
  "[fix] Linux configure 캐시 fresh 초기화",
];
s4.addText(commitsToBullets(groupByTaskId(chrisCommits, _all.filter(t=>t.track==="ACS"))),
  {x:0.7,y:4.0,w:3.7,h:0.35,fontFace:"Calibri",paraSpaceAfter:2,margin:0});

// SDD 6계층
s4.addShape(pres.shapes.RECTANGLE,{x:4.7,y:0.95,w:4.8,h:3.85,fill:{color:NW},shadow:shade()});
s4.addShape(pres.shapes.RECTANGLE,{x:4.7,y:0.95,w:4.8,h:0.06,fill:{color:NN}});
s4.addText("SDD v2 — 6계층 시스템 아키텍처",{x:4.9,y:1.08,w:4.5,h:0.3,fontSize:11,fontFace:"Calibri",color:NN,bold:true,margin:0});

const lys = [
  {n:"1. State Ingress Adapter",d:"AGV/PLC state 수신·검증·정규화 → Service Layer 전달",c:NB},
  {n:"2. Service Layer",        d:"Robot/Mission/Traffic 상태 관리, 도메인 이벤트 발행",c:NGN},
  {n:"3. Agent Layer",          d:"Load Balancing / Charging / MAPF Agent 판단 로직",   c:NPUR},
  {n:"4. Agent Broker",         d:"Ticket 우선순위 선정·최종 실행 Ticket 결정",            c:NAMB},
  {n:"5. Main Controller",      d:"선택된 Ticket → Mission 형태로 실행기에 전달",           c:NTEAL},
  {n:"6. VDA5050 Robot Executor",d:"Mission → VDA5050 프로토콜로 AGV Fleet 전달",        c:NRED},
];
lys.forEach((l,i)=>{
  const y=1.52+i*0.42;
  s4.addShape(pres.shapes.RECTANGLE,{x:4.85,y:y+0.05,w:0.06,h:0.3,fill:{color:l.c}});
  s4.addText(l.n,{x:5.0,y,w:2.4,h:0.34,fontSize:8.5,fontFace:"Calibri",color:NDARK,bold:true,valign:"middle",margin:0});
  s4.addText(l.d,{x:7.4,y,w:2.0,h:0.34,fontSize:7.5,fontFace:"Calibri",color:NDGR,valign:"middle",margin:0});
});

// Ticket 상태 요약
s4.addShape(pres.shapes.RECTANGLE,{x:4.7,y:4.0,w:4.8,h:1.3,fill:{color:NBL},shadow:shade()});
s4.addText("Ticket / Mission 상태 흐름 (SDD v2)",{x:4.9,y:4.1,w:4.5,h:0.28,fontSize:10.5,fontFace:"Calibri",color:NN,bold:true,margin:0});
s4.addText([
  {text:"Ticket: Create → planned → Selected → Executing → Completed/Failed",options:{bullet:true,breakLine:true,fontSize:8.5,color:NDARK}},
  {text:"         ↘ Deferred → Canceled / Rejected (Agent Broker 보류/기각)",options:{breakLine:true,fontSize:8,color:NDGR}},
  {text:"Mission: Created → Accepted → InProgress → Completed/Failed/Canceled",options:{bullet:true,fontSize:8.5,color:NDARK}},
],{x:4.9,y:4.42,w:4.5,h:0.82,fontFace:"Calibri",paraSpaceAfter:4,margin:0});

// ══════════════════════════════════════════════
// SLIDE 5 : UI/UX 파트
// ══════════════════════════════════════════════
let s5 = pres.addSlide();
hdr(s5, "UI/UX 파트 상세", "navifra-ui/brain_frontend + brain_backend  |  Vue/Node");

s5.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:5.8,h:3.5,fill:{color:NW},shadow:shade()});
s5.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:5.8,h:0.06,fill:{color:NB}});
s5.addText("숀 (navifra-sean)",{x:0.7,y:1.08,w:4.5,h:0.3,fontSize:12,fontFace:"Calibri",color:NN,bold:true,margin:0});
s5.addText("21커밋 (FE+BE)",{x:5.2,y:1.08,w:1.0,h:0.3,fontSize:9,fontFace:"Calibri",color:NB,bold:true,align:"right",margin:0});
// ── 커밋 목록 정의 (매주 이 부분만 수정) ──────────────────
const seanCommits = [
  "[feat] Audit Log (Frontend+Backend) #707/#606 ← audit_logs 테이블 연동",
  "[feat] API Swagger Frontend/Backend 테스트 #657 / jsDoc 변환 #652",
  "[feat] Job.message 추가 #816/#655 / ACS Job 기록 최적화 #660",
  "[fix]  맵 리스트 Text Truncate #813 / AppBar CSS 버그 #809",
  "[feat] 다크 테마 배경 수정 #812 / Favorite star→heart #811",
  "[feat] Package Upgrade #810/#650 / 404 다국어 처리 #806",
  "[refac] Model allowNull: true 제거 #656",
];
s5.addText(commitsToBullets(groupByTaskId(seanCommits, _all.filter(t=>t.track==="UI"))),
  {x:0.7,y:1.45,w:5.4,h:2.9,fontFace:"Calibri",paraSpaceAfter:3,margin:0});

s5.addShape(pres.shapes.RECTANGLE,{x:6.7,y:0.95,w:2.8,h:1.0,fill:{color:NW},shadow:shade()});
s5.addShape(pres.shapes.RECTANGLE,{x:6.7,y:0.95,w:2.8,h:0.06,fill:{color:NAMB}});
s5.addText("종원선 (Navifra-Pedro)",{x:6.9,y:1.08,w:2.5,h:0.28,fontSize:11,fontFace:"Calibri",color:NN,bold:true,margin:0});
s5.addText("1커밋",{x:9.1,y:1.08,w:0.3,h:0.28,fontSize:9,fontFace:"Calibri",color:NAMB,bold:true,align:"right",margin:0});
const pedroCommits = [
  "[feat] brain desktop 프로그램 통합 #818",
];
s5.addText(commitsToBullets(groupByTaskId(pedroCommits, _all.filter(t=>t.track==="UI"))),
  {x:6.9,y:1.42,w:2.5,h:0.45,fontFace:"Calibri",paraSpaceAfter:2,margin:0});

// DB 설계서 박스
s5.addShape(pres.shapes.RECTANGLE,{x:6.7,y:2.1,w:2.8,h:2.35,fill:{color:NW},shadow:shade()});
s5.addShape(pres.shapes.RECTANGLE,{x:6.7,y:2.1,w:2.8,h:0.06,fill:{color:NGN}});
s5.addText("✅  DB 설계서 v0.3 완성",{x:6.9,y:2.22,w:2.5,h:0.3,fontSize:10.5,fontFace:"Calibri",color:NGN,bold:true,margin:0});
s5.addText("2026-03-12 작성 / 2026-03-13 최종",{x:6.9,y:2.56,w:2.5,h:0.3,fontSize:8.5,fontFace:"Calibri",color:NDGR,margin:0});
s5.addText([
  {text:"인증/사용자  (3 테이블)",options:{bullet:true,breakLine:true,fontSize:8.5,color:NDARK}},
  {text:"맵/스캔맵   (11 테이블)",options:{bullet:true,breakLine:true,fontSize:8.5,color:NDARK}},
  {text:"PLC          (1 테이블)",options:{bullet:true,breakLine:true,fontSize:8.5,color:NDARK}},
  {text:"공통코드/그룹(2 테이블)",options:{bullet:true,breakLine:true,fontSize:8.5,color:NDARK}},
  {text:"ACS 그룹     (1 테이블)",options:{bullet:true,fontSize:8.5,color:NDARK}},
],{x:6.9,y:2.92,w:2.5,h:1.45,fontFace:"Calibri",paraSpaceAfter:4,margin:0});

s5.addShape(pres.shapes.RECTANGLE,{x:0.5,y:4.6,w:9,h:0.7,fill:{color:NW},shadow:shade()});
s5.addText("이슈: #818 Brain Desktop · #707 Audit Log · #657 Swagger · #816 Job.message · #812 다크테마 · #810 패키지",{
  x:0.7,y:4.7,w:8.6,h:0.5,fontSize:10,fontFace:"Calibri",color:NN,align:"center",valign:"middle",margin:0
});

// ══════════════════════════════════════════════
// SLIDE 6 : 신규 산출물 (DB 테이블 + SDD 클래스)
// ══════════════════════════════════════════════
let s6 = pres.addSlide();
hdr(s6, "신규 산출물 — DB 설계서 v0.3 + SDD 클래스", "2026-03-12~13 완성");

// DB 도메인
s6.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:4.5,h:4.35,fill:{color:NW},shadow:shade()});
s6.addShape(pres.shapes.RECTANGLE,{x:0.5,y:0.95,w:4.5,h:0.06,fill:{color:NGN}});
s6.addText("H-ACS 2.0 DB 설계서 v0.3 — 도메인별 테이블",{x:0.7,y:1.08,w:4.2,h:0.3,fontSize:11,fontFace:"Calibri",color:NN,bold:true,margin:0});

const dbD=[
  {d:"인증/사용자",  ts:["acs_user","audit_logs","config"],c:NPUR},
  {d:"공통코드/그룹",ts:["comm_code_group","comm_code","acs_groups"],c:NTEAL},
  {d:"맵/스캔맵",   ts:["acs_map","acs_map_areas","acs_map_node","acs_map_link","acs_map_marker","acs_map_polygon","acs_map_portals","acs_map_portal_nodes","acs_map_devices","acs_map_plcs","acs_map_tags","acs_map_node_teching"],c:NB},
  {d:"PLC",         ts:["acs_plcs"],c:NAMB},
];
let yo=1.48;
dbD.forEach(d=>{
  s6.addShape(pres.shapes.RECTANGLE,{x:0.6,y:yo,w:0.1,h:0.24,fill:{color:d.c}});
  s6.addText(d.d,{x:0.78,y:yo,w:1.5,h:0.24,fontSize:9,fontFace:"Calibri",color:NDARK,bold:true,valign:"middle",margin:0});
  s6.addText(`(${d.ts.length}개)`,{x:2.22,y:yo,w:0.5,h:0.24,fontSize:8,fontFace:"Calibri",color:d.c,bold:true,valign:"middle",margin:0});
  yo+=0.26;
  const rows=Math.ceil(d.ts.length/2);
  s6.addText(d.ts.join("  ·  "),{x:0.78,y:yo,w:4.05,h:rows*0.28,fontSize:7.5,fontFace:"Calibri",color:NDGR,margin:0});
  yo+=rows*0.28+0.06;
});
s6.addText("공통 규칙: UUID PK · snake_case · soft delete · JSON 컬럼",{x:0.7,y:4.86,w:4.1,h:0.28,fontSize:7.5,fontFace:"Calibri",color:NDGR,margin:0});

// SDD 클래스
s6.addShape(pres.shapes.RECTANGLE,{x:5.2,y:0.95,w:4.3,h:4.35,fill:{color:NW},shadow:shade()});
s6.addShape(pres.shapes.RECTANGLE,{x:5.2,y:0.95,w:4.3,h:0.06,fill:{color:NN}});
s6.addText("SDD v2 — 주요 클래스 구조 (이번 주 구현)",{x:5.4,y:1.08,w:4.0,h:0.3,fontSize:11,fontFace:"Calibri",color:NN,bold:true,margin:0});

const cgs=[
  {l:"Service 계층",items:["cBaseService (abstract)","cMapService → cBaseService","SERVICE_STATE enum"],c:NGN},
  {l:"Agent 계층", items:["cBaseAgent (abstract)","AGENT_STATE enum","IControllerClient"],c:NPUR},
  {l:"Repository / Store",items:["IDataStore<T>","cConcurrentDataStore","cIndexedDataStore","cVersionedDataStore","cGraph/Device/RobotRepository"],c:NB},
  {l:"통신 모듈",items:["ICommunicationClient  (interface)","cCommunicator","cRequestClient / cStreamClient","cHttpsClient / cMqttClient","PROTOCOL_TYPE enum"],c:NTEAL},
];
let yc=1.48;
cgs.forEach(g=>{
  s6.addShape(pres.shapes.RECTANGLE,{x:5.3,y:yc,w:0.06,h:g.items.length*0.24,fill:{color:g.c}});
  s6.addText(g.l,{x:5.45,y:yc,w:3.9,h:0.26,fontSize:9.5,fontFace:"Calibri",color:NDARK,bold:true,margin:0});
  yc+=0.27;
  g.items.forEach(it=>{
    s6.addText("• "+it,{x:5.55,y:yc,w:3.8,h:0.24,fontSize:8.5,fontFace:"Calibri",color:NDGR,margin:0});
    yc+=0.24;
  });
  yc+=0.06;
});

// ══════════════════════════════════════════════
// SLIDE 7 : 리스크 & 이슈
// ══════════════════════════════════════════════
let s7 = pres.addSlide();
hdr(s7, "리스크 & 이슈", PERIOD_SHORT);

const risks=[
  {lv:"High",item:"P-02 시스템 아키텍처 설계 지연",detail:"종료주(W8) 초과 — SDD v2·실제 코드 정합성 확인됨, 공식 완료 처리 필요",c:NRED},
  {lv:"Mid", item:"ACS/UI PLC 연동 미착수",detail:"L-03/L-04 이번 주~다음 주 착수 예정 / DB 설계서 완성으로 구현 기반 마련됨",c:NAMB},
  {lv:"Mid", item:"전체 완료율 계획 대비 갭 -4.9%p",detail:"계획 기대치 7.4% vs 실제 2.5% — DB/SDD 문서 완성으로 다음 주 속도 개선 기반 확보",c:NAMB},
  {lv:"Low", item:"Navifra-Garry BE 기여 H-ACS 라벨 없음",detail:"JSDoc+Swagger 기여 확인됨 — H-ACS-V2_25065DNB 이슈 라벨 적용 요청 필요",c:NB},
];
risks.forEach((r,i)=>{
  const y=1.02+i*0.92;
  s7.addShape(pres.shapes.RECTANGLE,{x:0.5,y,w:9,h:0.8,fill:{color:NW},shadow:shade()});
  s7.addShape(pres.shapes.RECTANGLE,{x:0.5,y,w:0.06,h:0.8,fill:{color:r.c}});
  s7.addShape(pres.shapes.RECTANGLE,{x:0.65,y:y+0.12,w:0.65,h:0.26,fill:{color:r.c}});
  s7.addText(r.lv,{x:0.65,y:y+0.12,w:0.65,h:0.26,fontSize:8,fontFace:"Calibri",color:NW,bold:true,align:"center",valign:"middle",margin:0});
  s7.addText(r.item,{x:1.45,y:y+0.07,w:7.8,h:0.28,fontSize:11,fontFace:"Calibri",color:NDARK,bold:true,margin:0});
  s7.addText(r.detail,{x:1.45,y:y+0.38,w:7.8,h:0.28,fontSize:9.5,fontFace:"Calibri",color:NDGR,margin:0});
});

// ══════════════════════════════════════════════
// SLIDE 8 : 다음 주 계획
// ══════════════════════════════════════════════
let s8 = pres.addSlide();
hdr(s8, "다음 주 계획", "3월 3주차 (3/16~3/20)");

const plans=[
  {t:"ACS",c:NGN,items:["A-02 VDA5050 MQTT 통신 기반 완료 (W11 종료)","A-03 Backend API 연동 완료","A-04 로봇 시뮬레이터 단순 경로 주행 완료","A-05 경로 생성(A*/Dijkstra) 완료","P-05 SDD 클래스 다이어그램 1차 완성","P-02 아키텍처 설계 공식 완료 처리"]},
  {t:"UI/UX",c:NB,items:["U-01 노드/맵 편집 기능 완료 (W11 종료)","U-03 로봇 객체 그룹 관리 UI 완료","U-04 PLC 객체 DB 관리 UI 착수","Swagger API 문서 정비 지속"]},
  {t:"PLC",c:NAMB,items:["L-03 PLC 객체 등록/수정/삭제 API 착수","L-04 Polling 방식 TAG 수집 구현 시작","P-07 H-PLC 아키텍처 설계 완료 (W11)","acs_plcs / acs_plc_tags 테이블 기반 API"]},
];
plans.forEach((p,i)=>{
  const x=0.5+i*3.17;
  s8.addShape(pres.shapes.RECTANGLE,{x,y:0.95,w:2.95,h:4.3,fill:{color:NW},shadow:shade()});
  s8.addShape(pres.shapes.RECTANGLE,{x,y:0.95,w:2.95,h:0.06,fill:{color:p.c}});
  s8.addText(p.t,{x:x+0.2,y:1.08,w:2.5,h:0.3,fontSize:13,fontFace:"Calibri",color:NN,bold:true,margin:0});
  const items=p.items.map(it=>({text:it,options:{bullet:true,breakLine:true,fontSize:9.5,color:NDARK}}));
  s8.addText(items,{x:x+0.15,y:1.5,w:2.65,h:3.6,fontFace:"Calibri",paraSpaceAfter:6,margin:0});
});

// ══════════════════════════════════════════════
// SLIDE 9 : 마일스톤 진척률
// ══════════════════════════════════════════════
let s9 = pres.addSlide();
s9.background = { color: NGR };
s9.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.8,fill:{color:NN}});
s9.addText("마일스톤 진척률",{x:0.7,y:0,w:5.5,h:0.8,fontSize:20,fontFace:"Calibri",color:NW,bold:true,valign:"middle",margin:0});
s9.addText(`W${DATA_WEEK+1} 기준 (${TODAY})  |  전체 ${_all.length}개 Task`,{x:4.5,y:0,w:5.2,h:0.8,fontSize:11,fontFace:"Calibri",color:"93B5D8",align:"right",valign:"middle",margin:0});

const mk=[
  {l:"전체 완료율", v:`${(_rate*100).toFixed(1)}%`,   s:`${_done.length}/${_all.length} Task`, c:NGN},
  {l:"진행 중",    v:`${_prog.length}건`,              s:`${_prog.length}개 Task`,              c:NB},
  {l:"대기",       v:`${_todo.length}건`,              s:"미착수",                               c:NDGR},
  {l:"계획 대비 갭",v:`${_gap>=0?"+":""}${(_gap*100).toFixed(1)}%p`, s:`계획 ${(_planR*100).toFixed(1)}%`, c:_gap>=0?NGN:NRED},
];
mk.forEach((k,i)=>{
  const x=0.5+i*2.3;
  s9.addShape(pres.shapes.RECTANGLE,{x,y:0.95,w:2.1,h:1.2,fill:{color:NW},shadow:shade()});
  s9.addShape(pres.shapes.RECTANGLE,{x,y:0.95,w:2.1,h:0.05,fill:{color:k.c}});
  s9.addText(k.l,{x,y:1.06,w:2.1,h:0.25,fontSize:9,fontFace:"Calibri",color:NDGR,align:"center",margin:0});
  s9.addText(k.v,{x,y:1.3, w:2.1,h:0.45,fontSize:22,fontFace:"Calibri",color:NN,bold:true,align:"center",margin:0});
  s9.addText(k.s,{x,y:1.75,w:2.1,h:0.3, fontSize:8,fontFace:"Calibri",color:NDGR,align:"center",margin:0});
});

s9.addShape(pres.shapes.RECTANGLE,{x:0.5,y:2.3,w:5.5,h:3.0,fill:{color:NW},shadow:shade()});
s9.addText("트랙별 완료율",{x:0.7,y:2.42,w:5,h:0.3,fontSize:12,fontFace:"Calibri",color:NN,bold:true,margin:0});
const tc={"기획":NPUR,"ACS":NGN,"UI":NB,"PLC":NAMB,"통합":NTEAL};
const td=_td; // data.js에서 자동 계산
Object.entries(td).forEach(([tr,d],i)=>{
  const y=2.85+i*0.45; const r=d.total>0?d.done/d.total:0;
  s9.addText(tr,{x:0.7,y,w:0.8,h:0.32,fontSize:9,fontFace:"Calibri",color:NDARK,bold:true,valign:"middle",margin:0});
  s9.addText(`${d.done}/${d.total}`,{x:1.45,y,w:0.5,h:0.32,fontSize:8,fontFace:"Calibri",color:NDGR,valign:"middle",margin:0});
  s9.addShape(pres.shapes.RECTANGLE,{x:1.95,y:y+0.06,w:3.6,h:0.2,fill:{color:"E0E6EF"}});
  if(r>0) s9.addShape(pres.shapes.RECTANGLE,{x:1.95,y:y+0.06,w:Math.max(0.05,3.6*r),h:0.2,fill:{color:tc[tr]||NB}});
  s9.addText(`${Math.round(r*100)}%`,{x:5.6,y,w:0.35,h:0.32,fontSize:8,fontFace:"Calibri",color:NDARK,bold:true,valign:"middle",margin:0});
});

s9.addShape(pres.shapes.RECTANGLE,{x:6.2,y:2.3,w:3.3,h:3.0,fill:{color:NW},shadow:shade()});
s9.addText("지연 · 위험 항목",{x:6.4,y:2.42,w:3.0,h:0.3,fontSize:12,fontFace:"Calibri",color:NN,bold:true,margin:0});
// 지연 항목 자동 표시 (data.js 기반)
_delayed.slice(0,2).forEach((t,i)=>{
  const y=2.82+i*0.52;
  s9.addShape(pres.shapes.RECTANGLE,{x:6.3,y,w:0.55,h:0.22,fill:{color:NRED}});
  s9.addText("지연",{x:6.3,y,w:0.55,h:0.22,fontSize:7.5,fontFace:"Calibri",color:NW,bold:true,align:"center",valign:"middle",margin:0});
  s9.addText(`${t.id} ${t.name}`,{x:6.9,y,w:2.5,h:0.22,fontSize:8,fontFace:"Calibri",color:NDARK,valign:"middle",margin:0});
  s9.addText(`종료 W${t.e+1} 초과 → 진행 중`,{x:6.9,y:y+0.23,w:2.5,h:0.18,fontSize:7,fontFace:"Calibri",color:NDGR,margin:0});
});
s9.addText("다음 주 착수 예정",{x:6.4,y:3.4,w:3.0,h:0.25,fontSize:9,fontFace:"Calibri",color:NN,bold:true,margin:0});
// 착수 예정 항목 자동 표시 (data.js 기반)
_upcoming.slice(0,4).forEach((u,i)=>{
  const y=3.7+i*0.35;
  s9.addShape(pres.shapes.RECTANGLE,{x:6.3,y:y+0.02,w:0.45,h:0.2,fill:{color:NAMB}});
  s9.addText(`W${u.s+1}`,{x:6.3,y:y+0.02,w:0.45,h:0.2,fontSize:7,fontFace:"Calibri",color:NW,bold:true,align:"center",valign:"middle",margin:0});
  s9.addText(`${u.id} ${u.name}`,{x:6.8,y,w:2.6,h:0.3,fontSize:8,fontFace:"Calibri",color:NDARK,valign:"middle",margin:0});
});

// ══════════════════════════════════════════════
// SLIDE 10 : 프로젝트 로드맵 현황
// ══════════════════════════════════════════════
let s10 = pres.addSlide();
s10.background = { color: NGR };
s10.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.8,fill:{color:NN}});
s10.addText("프로젝트 로드맵 현황",{x:0.7,y:0,w:6,h:0.8,fontSize:20,fontFace:"Calibri",color:NW,bold:true,valign:"middle",margin:0});
s10.addText("W11 기준  |  2026.03.14",{x:5,y:0,w:4.7,h:0.8,fontSize:11,fontFace:"Calibri",color:"93B5D8",align:"right",valign:"middle",margin:0});

// 카드 정의: [x, y, w, h, color, badgeText, badgeColor, title, items[]]
const projCards = [
  {
    x:0.3, y:0.9, w:4.55, h:2.15,
    accent: "D97706",   // amber
    badge: "일정 변경",  badgeColor: "D97706",
    title: "2.1  Verification Toolkit  (구 시뮬레이션)",
    items: [
      "PLC 시뮬레이션: 4월 산출물 추가 필요",
      "명칭/방향: 시뮬레이션 → Verification Toolkit으로 정리",
      "10월 울산 OLT 일정 변경",
    ],
  },
  {
    x:5.15, y:0.9, w:4.55, h:2.15,
    accent: NB,
    badge: "논의중",     badgeColor: NB,
    title: "1.1  울산 C 프로젝트",
    items: [],
  },
  {
    x:0.3, y:3.15, w:4.55, h:2.15,
    accent: "0D9488",
    badge: "준비중",    badgeColor: "0D9488",
    title: "1.2  천안아산 프로젝트",
    items: [
      "로봇 2대 + 1대(?) 투입 예정 / 레이아웃 비교적 간단",
      "공사 기간: 9월 추석 휴무 기간 활용 계획",
      "9월 목표: 미들웨어 등 전체 적용 (NaviCore 포함 가능성)",
    ],
  },
  {
    x:5.15, y:3.15, w:4.55, h:2.15,
    accent: "7C3AED",
    badge: "8월 통합테스트",  badgeColor: "7C3AED",
    title: "1.3  EVO 화성 프로젝트",
    items: [
      "8월 통합 테스트 → 화성 공장에서 진행 예정",
      "9월 천안아산 적용 필수 (연계 일정)",
    ],
  },
];

projCards.forEach(p => {
  s10.addShape(pres.shapes.RECTANGLE,{x:p.x,y:p.y,w:p.w,h:p.h,fill:{color:NW},shadow:{type:"outer",blur:4,offset:2,angle:135,color:"000000",opacity:0.08}});
  s10.addShape(pres.shapes.RECTANGLE,{x:p.x,y:p.y,w:p.w,h:0.05,fill:{color:p.accent}});
  // 배지
  s10.addShape(pres.shapes.RECTANGLE,{x:p.x+p.w-1.18,y:p.y+0.12,w:1.1,h:0.24,fill:{color:p.badgeColor}});
  s10.addText(p.badge,{x:p.x+p.w-1.18,y:p.y+0.12,w:1.1,h:0.24,fontSize:7.5,fontFace:"Calibri",color:NW,bold:true,align:"center",valign:"middle",margin:0});
  // 제목
  s10.addText(p.title,{x:p.x+0.15,y:p.y+0.1,w:p.w-1.4,h:0.32,fontSize:10.5,fontFace:"Calibri",color:NN,bold:true,valign:"middle",margin:0});
  // 구분선
  s10.addShape(pres.shapes.RECTANGLE,{x:p.x+0.15,y:p.y+0.48,w:p.w-0.3,h:0.02,fill:{color:"E2E8F0"}});
  // 항목
  const txtItems = p.items.map(it=>({text:it,options:{bullet:true,breakLine:true,fontSize:9,color:NDARK}}));
  s10.addText(txtItems,{x:p.x+0.15,y:p.y+0.55,w:p.w-0.3,h:p.h-0.65,fontFace:"Calibri",paraSpaceAfter:4,margin:0});
});

pres.writeFile({fileName:OUTPUT_PM}).then(()=>console.log("✅ SAVED:", OUTPUT_PM))
.catch(e=>{console.error(e);process.exit(1);});
