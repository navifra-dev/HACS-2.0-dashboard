const pptxgen = require("pptxgenjs");
const fs = require("fs");

// ── data.js 공유 데이터 연동 ────────────────────────────────
const { tasks, milestones, todayWeek: DATA_WEEK } = require("../docs/data.js");
const _all  = tasks.filter(t => !t.cat);
const _done = _all.filter(t => t.st === "done");
const _prog = _all.filter(t => t.st === "progress");
const _rate = _all.length > 0 ? _done.length / _all.length : 0;

const PERIOD = "2026년 3월 2주차 (3/9~3/13)";
const OUTPUT = "/home/sally/claude_cowork/HACS2.0_주간보고_고객사_11주차.pptx";
const TODAY = "2026.03.14";

const NAVY = "1A2744";
const BLUE = "2563EB";
const WHITE = "FFFFFF";
const DARK = "1E293B";
const GRAY = "64748B";
const LIGHT_GRAY = "F8FAFC";
const GREEN = "059669";
const AMBER = "D97706";
const SLATE = "94A3B8";
const RED = "DC2626";
const PURPLE = "7C3AED";
const TEAL = "0D9488";

const makeShadow = () => ({ type: "outer", blur: 4, offset: 2, angle: 135, color: "000000", opacity: 0.08 });

let pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "나비프라";
pres.title = "H-ACS 2.0 주간 진척 보고 - 3월 2주차";

function addHeader(slide, title, subtitle) {
  slide.background = { color: LIGHT_GRAY };
  slide.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.85, fill: { color: NAVY } });
  slide.addText(title, { x: 0.7, y: 0, w: 6, h: 0.85, fontSize: 22, fontFace: "Calibri", color: WHITE, bold: true, valign: "middle", margin: 0 });
  if (subtitle) {
    slide.addText(subtitle, { x: 5.5, y: 0, w: 4.2, h: 0.85, fontSize: 10, fontFace: "Calibri", color: SLATE, align: "right", valign: "middle", margin: 0 });
  }
  slide.addText("H-ACS 2.0  |  나비프라", { x: 0.5, y: 5.3, w: 4, h: 0.25, fontSize: 7, fontFace: "Calibri", color: SLATE, margin: 0 });
  slide.addText("Confidential", { x: 7, y: 5.3, w: 2.5, h: 0.25, fontSize: 7, fontFace: "Calibri", color: SLATE, align: "right", margin: 0 });
}

// SLIDE 1: 표지
let s1 = pres.addSlide();
s1.background = { color: NAVY };
s1.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.05, fill: { color: BLUE } });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 0.55, w: 2.2, h: 0.45, fill: { color: BLUE } });
s1.addText("NAVIFRA", { x: 0.7, y: 0.55, w: 2.2, h: 0.45, fontSize: 16, fontFace: "Calibri", color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
s1.addText("H-ACS 2.0", { x: 0.7, y: 1.7, w: 8.6, h: 0.9, fontSize: 44, fontFace: "Calibri", color: WHITE, bold: true, margin: 0 });
s1.addText("주간 진척 보고", { x: 0.7, y: 2.5, w: 8.6, h: 0.65, fontSize: 30, fontFace: "Calibri", color: "93B5E8", margin: 0 });
s1.addShape(pres.shapes.RECTANGLE, { x: 0.7, y: 3.35, w: 2.5, h: 0.03, fill: { color: BLUE } });
s1.addText(PERIOD, { x: 0.7, y: 3.6, w: 6, h: 0.4, fontSize: 16, fontFace: "Calibri", color: "93B5E8", margin: 0 });
s1.addText("현대자동차그룹  |  H-ACS 2.0 개발 고도화", { x: 0.7, y: 4.6, w: 6, h: 0.3, fontSize: 11, fontFace: "Calibri", color: "7A99B8", margin: 0 });
s1.addText("나비프라  |  Confidential", { x: 0.7, y: 4.9, w: 5, h: 0.3, fontSize: 10, fontFace: "Calibri", color: "5A7998", margin: 0 });
s1.addText(TODAY, { x: 7.5, y: 4.9, w: 2, h: 0.3, fontSize: 10, fontFace: "Calibri", color: "5A7998", align: "right", margin: 0 });

// SLIDE 2: H-ACS 2.0 개발 진행 현황 표 (고객사 제출용)
let s2t = pres.addSlide();
s2t.background = { color: "F8FAFC" };

// 헤더 바
s2t.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 10, h: 0.6, fill: { color: NAVY } });
s2t.addText("H-ACS 2.0  개발 진행 현황", {
  x: 0.3, y: 0, w: 7, h: 0.6,
  fontSize: 18, fontFace: "Calibri", color: WHITE, bold: true, valign: "middle", margin: 0
});
// Confidential 배지
s2t.addShape(pres.shapes.RECTANGLE, { x: 8.1, y: 0.1, w: 1.6, h: 0.38, fill: { color: WHITE }, line: { color: "DC2626", width: 1.5 } });
s2t.addText("Confidential", {
  x: 8.1, y: 0.1, w: 1.6, h: 0.38,
  fontSize: 9, fontFace: "Calibri", color: "DC2626", bold: true, align: "center", valign: "middle", margin: 0
});

// 기간 표시
s2t.addText("보고 기간 : 2026.03.09 ~ 03.13", {
  x: 0.3, y: 0.65, w: 5, h: 0.28,
  fontSize: 9, fontFace: "Calibri", color: "64748B", margin: 0
});

// ACS 실적/계획 내용
const acs_실적_c =
  "• Map Service 완성 (지도/포털/노드/링크)\n" +
  "• MQTT Callback / HTTP Token 인증 연결\n" +
  "• Config Adapter / PROTOCOL_TYPE Enum 추가\n" +
  "• Ticket/Mission SDD v2 설계 완료";
const acs_계획_c =
  "• VDA5050 로봇 시뮬레이터 1대 주행 검증\n" +
  "• 경로 생성 알고리즘(A*/Dijkstra) 완료\n" +
  "• Agent Layer 구현 착수";

// UI 실적/계획 내용
const ui_실적_c =
  "• Audit Log 기능 완료 (FE+BE)\n" +
  "• Swagger API 문서화 완료\n" +
  "• Brain Desktop 통합 앱 첫 빌드 (#818)\n" +
  "• 노드/맵 편집 기능 (U-01) 완료";
const ui_계획_c =
  "• 로봇 객체 그룹 관리 UI (U-03) 완료\n" +
  "• PLC 객체 DB 관리 UI (U-04) 착수\n" +
  "• Swagger API 문서 정비 지속";

// 공지 내용
const notice_c =
  "• SDD 클래스 다이어그램 1차 완성 예정 (P-05, 3월 3주차)\n" +
  "• DB 설계서 v0.3 완성 (2026-03-13) — 14개 테이블 확정";

// 헬퍼 함수
const hc2 = (extra={}) => ({ fill:{color:NAVY}, color:WHITE, bold:true, fontSize:10, fontFace:"Calibri", align:"center", valign:"middle", margin:[4,4,4,4], border:[{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"}], ...extra });
const sh2 = (extra={}) => ({ fill:{color:"1E3A6E"}, color:"BDD7F7", bold:true, fontSize:9, fontFace:"Calibri", align:"center", valign:"middle", margin:[3,3,3,3], border:[{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"}], ...extra });
const cat2 = (extra={}) => ({ fill:{color:"1E3A6E"}, color:WHITE, bold:true, fontSize:11, fontFace:"Calibri", align:"center", valign:"middle", margin:[4,4,4,4], border:[{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"},{pt:0.5,color:"FFFFFF"}], ...extra });
const sub2 = (extra={}) => ({ fill:{color:"EBF3FE"}, color:NAVY, bold:true, fontSize:10, fontFace:"Calibri", align:"center", valign:"middle", margin:[4,4,4,4], border:[{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"}], ...extra });
const con2 = (extra={}) => ({ fill:{color:WHITE}, color:"1E293B", fontSize:9, fontFace:"Calibri", align:"left", valign:"top", margin:[5,5,5,5], border:[{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"}], ...extra });
const alt2 = (extra={}) => ({ fill:{color:"F1F5F9"}, color:"1E293B", fontSize:9, fontFace:"Calibri", align:"left", valign:"top", margin:[5,5,5,5], border:[{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"},{pt:0.3,color:"CBD5E1"}], ...extra });

const tableRows = [
  // Row 0 — 헤더
  [
    { text:"과제명",    options: hc2({ rowspan:2 }) },
    { text:"구분",      options: hc2({ rowspan:2 }) },
    { text:"담당",      options: hc2({ rowspan:2 }) },
    { text:"진행 현황", options: hc2({ colspan:2 }) },
  ],
  // Row 1 — 서브헤더
  [
    { text:"실적 (3/9~3/13)",  options: sh2() },
    { text:"계획 (3/16~3/20)", options: sh2() },
  ],
  // Row 2 — ACS
  [
    { text:"H-ACS 2.0", options: cat2({ rowspan:3, fontSize:11 }) },
    { text:"ACS",       options: sub2() },
    { text:"권익현\n최명서", options: sub2({ fontSize:8.5 }) },
    { text: acs_실적_c, options: con2() },
    { text: acs_계획_c, options: con2({ fill:{color:"EBF3FE"} }) },
  ],
  // Row 3 — UX/UI
  [
    { text:"UX/UI",     options: sub2() },
    { text:"종원선\n박새롬", options: sub2({ fontSize:8.5 }) },
    { text: ui_실적_c,  options: alt2() },
    { text: ui_계획_c,  options: alt2({ fill:{color:"EBF3FE"} }) },
  ],
  // Row 4 — 기획
  [
    { text:"기획",      options: sub2() },
    { text:"최준섭",    options: sub2({ fontSize:9 }) },
    { text: "",         options: con2() },
    { text: "",         options: con2({ fill:{color:"EBF3FE"} }) },
  ],
  // Row 5 — 주요 공지
  [
    { text:"주요 공지", options: sub2({ colspan:2, valign:"middle" }) },
    { text:"담당",      options: sub2({ fontSize:8.5, valign:"middle" }) },
    { text: notice_c,   options: con2({ colspan:2 }) },
  ],
];

s2t.addTable(tableRows, {
  x: 0.25, y: 0.98,
  w: 9.5,
  colW: [0.82, 0.78, 0.78, 3.56, 3.56],
  rowH: [0.3, 0.28, 1.1, 1.25, 0.75, 0.55],
});

// 하단 메모
s2t.addText("나비프라  |  H-ACS 2.0  |  " + TODAY, {
  x: 0.3, y: 5.3, w: 5, h: 0.25,
  fontSize: 7, fontFace: "Calibri", color: "94A3B8", margin: 0
});
s2t.addText("Confidential", {
  x: 7.5, y: 5.3, w: 2.2, h: 0.25,
  fontSize: 7, fontFace: "Calibri", color: "94A3B8", align: "right", margin: 0
});

// SLIDE 2 (→3): 금주 진척 요약
let s2 = pres.addSlide();
addHeader(s2, "금주 진척 요약", "2026.03.09 ~ 03.13");

const statusCards = [
  { label: "ACS Map Service", status: "개발 완료", color: GREEN, detail: "6계층 아키텍처 기반 핵심 서비스 완성" },
  { label: "DB 설계서 v0.3", status: "완성", color: TEAL, detail: "14개 테이블 확정 (2026-03-13 최종 수정)" },
  { label: "UI/UX", status: "고도화 진행", color: BLUE, detail: "Audit Log, Swagger, Desktop 통합 완료" },
  { label: "전체 진행률", status: `${(_rate*100).toFixed(1)}%`, color: PURPLE, detail: `${_all.length}개 Task 중 ${_done.length}완료, ${_prog.length}진행` }
];
statusCards.forEach((card, i) => {
  const x = 0.5 + i * 2.3;
  s2.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.1, h: 1.5, fill: { color: WHITE }, shadow: makeShadow() });
  s2.addShape(pres.shapes.RECTANGLE, { x, y: 0.95, w: 2.1, h: 0.05, fill: { color: card.color } });
  s2.addShape(pres.shapes.RECTANGLE, { x: x + 0.2, y: 1.12, w: 1.0, h: 0.22, fill: { color: card.color } });
  s2.addText(card.status, { x: x + 0.2, y: 1.12, w: 1.0, h: 0.22, fontSize: 7.5, fontFace: "Calibri", color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
  s2.addText(card.label, { x: x + 0.1, y: 1.42, w: 1.9, h: 0.38, fontSize: 11, fontFace: "Calibri", color: DARK, bold: true, align: "center", margin: 0 });
  s2.addText(card.detail, { x: x + 0.1, y: 1.82, w: 1.9, h: 0.5, fontSize: 8, fontFace: "Calibri", color: GRAY, align: "center", margin: 0 });
});

s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.65, w: 9, h: 2.65, fill: { color: WHITE }, shadow: makeShadow() });
s2.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 2.65, w: 0.05, h: 2.65, fill: { color: BLUE } });
s2.addText("금주 핵심 성과", { x: 0.75, y: 2.78, w: 8, h: 0.35, fontSize: 14, fontFace: "Calibri", color: NAVY, bold: true, margin: 0 });
s2.addText([
  { text: "H-ACS DB 설계서 v0.3 완성 — MariaDB 기반 전체 스키마 확정", options: { bullet: true, breakLine: true, fontSize: 11, color: DARK } },
  { text: "맵/노드/링크/포털/PLC 위치 정보 및 사용자·감사 로그 테이블 14개 설계 완료", options: { bullet: true, indentLevel: 1, breakLine: true, fontSize: 9.5, color: GRAY } },
  { text: "ACS 지도 서비스(Map Service) 6계층 아키텍처 기반 구현 완료", options: { bullet: true, breakLine: true, fontSize: 11, color: DARK } },
  { text: "MQTT·HTTP 통신 기반, Repository 패턴(그래프/장비/로봇), Config 추상화 완성", options: { bullet: true, indentLevel: 1, breakLine: true, fontSize: 9.5, color: GRAY } },
  { text: "UI 감사 로그(Audit Log) 구현, API 문서화(Swagger), Brain Desktop 통합(#818)", options: { bullet: true, breakLine: true, fontSize: 11, color: DARK } },
  { text: "Ticket/Mission 상태 관리 흐름 설계 완성 (SDD v2) — 다음 주 구현 착수", options: { bullet: true, fontSize: 11, color: DARK } }
], { x: 0.75, y: 3.18, w: 8.5, h: 2.05, fontFace: "Calibri", paraSpaceAfter: 4, margin: 0 });

// SLIDE 3: ACS 기능 개발 현황 (6계층 구조 반영)
let s3 = pres.addSlide();
addHeader(s3, "ACS 기능 개발 현황", "6계층 Clean Architecture 기반");

const acsFuncs = [
  { title: "Map Service", status: "완료", statusColor: GREEN, details: [
    "6계층 중 Service Layer 핵심 서비스 완성",
    "지도/영역/포털/노드/링크 메시지 처리",
    "Device & Node 실시간 인덱스 갱신",
    "cMapService → cBaseService 상속 구조 완성"
  ]},
  { title: "통신 기반 (Infrastructure)", status: "완료", statusColor: GREEN, details: [
    "MQTT Subscribe Callback 연결",
    "HTTP Token 인증 처리",
    "Config Adapter (환경설정 추상화)",
    "PROTOCOL_TYPE Enum 신규 추가"
  ]},
  { title: "Ticket/Mission 흐름 설계", status: "설계완료", statusColor: TEAL, details: [
    "Ticket 상태: planned → Selected → Executing → Complete/Fail",
    "Mission 상태: Created → Accepted → InProgress → 완료",
    "Agent Broker Ticket 우선순위 정책 설계 완료",
    "다음 주 Agent Layer 구현 착수"
  ]},
  { title: "경로 생성 알고리즘", status: "진행", statusColor: AMBER, details: [
    "A* / Dijkstra 알고리즘 구현 진행 중",
    "cGraphRepository 패턴 기반",
    "이번 주(W11) 완료 목표"
  ]},
];
acsFuncs.forEach((f, i) => {
  const col = i % 2; const row = Math.floor(i / 2);
  const x = 0.5 + col * 4.8; const y = 0.95 + row * 2.3;
  s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 2.1, fill: { color: WHITE }, shadow: makeShadow() });
  s3.addShape(pres.shapes.RECTANGLE, { x, y, w: 4.5, h: 0.05, fill: { color: f.statusColor } });
  s3.addShape(pres.shapes.RECTANGLE, { x: x + 3.2, y: y + 0.1, w: 1.1, h: 0.22, fill: { color: f.statusColor } });
  s3.addText(f.status, { x: x + 3.2, y: y + 0.1, w: 1.1, h: 0.22, fontSize: 8, fontFace: "Calibri", color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
  s3.addText(f.title, { x: x + 0.2, y: y + 0.1, w: 2.9, h: 0.32, fontSize: 12, fontFace: "Calibri", color: DARK, bold: true, margin: 0 });
  const items = f.details.map(d => ({ text: d, options: { bullet: true, breakLine: true, fontSize: 9, color: DARK } }));
  s3.addText(items, { x: x + 0.2, y: y + 0.52, w: 4.1, h: 1.5, fontFace: "Calibri", paraSpaceAfter: 3, margin: 0 });
});

// SLIDE 4: UI/UX 기능 개발 현황
let s4 = pres.addSlide();
addHeader(s4, "UI/UX 기능 개발 현황", "Vue.js + Node.js");

const uiFuncs = [
  { title: "감사 로그 (Audit Log)", status: "완료", statusColor: GREEN, details: ["시스템 사용자 행위 이력 저장/조회", "DB 설계서 audit_logs 테이블 기반 구현", "Frontend + Backend 동시 완료"] },
  { title: "API 문서화 (Swagger)", status: "완료", statusColor: GREEN, details: ["JSDoc 기반 Swagger 자동 생성", "Frontend API 테스트 인터페이스 제공", "개발·운영 팀 간 API 표준화"] },
  { title: "Brain Desktop 통합", status: "완료", statusColor: GREEN, details: ["PC 설치형 통합 클라이언트 첫 빌드", "웹 UI 기능 데스크탑 앱 통합 (#818)"] },
  { title: "UX 개선 다수", status: "완료", statusColor: BLUE, details: ["다크 테마, 맵 리스트, AppBar CSS 수정", "404 다국어 처리, Favorite 아이콘", "Package 최신화 (FE+BE)"] },
  { title: "노드/맵 편집 (U-01)", status: "진행", statusColor: AMBER, details: ["노드/링크/영역 편집 기능 구현 중", "이번 주(W11) 완료 목표"] },
  { title: "Job 관리", status: "완료", statusColor: GREEN, details: ["ACS Job 기록 최적화 (#660)", "Job.message 필드 추가 (#816/#655)"] },
];
uiFuncs.forEach((f, i) => {
  const col = i % 3; const row = Math.floor(i / 3);
  const x = 0.35 + col * 3.2; const y = 0.95 + row * 2.3;
  s4.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 2.1, fill: { color: WHITE }, shadow: makeShadow() });
  s4.addShape(pres.shapes.RECTANGLE, { x, y, w: 3.0, h: 0.05, fill: { color: f.statusColor } });
  s4.addShape(pres.shapes.RECTANGLE, { x: x + 2.05, y: y + 0.1, w: 0.8, h: 0.22, fill: { color: f.statusColor } });
  s4.addText(f.status, { x: x + 2.05, y: y + 0.1, w: 0.8, h: 0.22, fontSize: 7.5, fontFace: "Calibri", color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
  s4.addText(f.title, { x: x + 0.15, y: y + 0.1, w: 1.85, h: 0.32, fontSize: 10.5, fontFace: "Calibri", color: DARK, bold: true, margin: 0 });
  const items = f.details.map(d => ({ text: d, options: { bullet: true, breakLine: true, fontSize: 8.5, color: DARK } }));
  s4.addText(items, { x: x + 0.15, y: y + 0.52, w: 2.7, h: 1.5, fontFace: "Calibri", paraSpaceAfter: 3, margin: 0 });
});

// SLIDE 5: 시스템 아키텍처 (SDD 6계층 기반으로 업데이트)
let s5 = pres.addSlide();
addHeader(s5, "시스템 아키텍처 현황", "Ticket/Mission SDD v2 기반 6계층 구조");

s5.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 1.0, w: 9, h: 4.2, fill: { color: WHITE }, shadow: makeShadow() });
s5.addText("계층별 구현 현황 (3월 2주차 기준 — SDD v2 확정)", { x: 0.7, y: 1.1, w: 8.5, h: 0.35, fontSize: 13, fontFace: "Calibri", color: NAVY, bold: true, margin: 0 });

const sddLayers = [
  { name: "⑥ VDA5050 Robot Executor", desc: "VDA5050 프로토콜로 AGV Fleet에 명령/상태 송수신", progress: 15, color: RED, note: "구현 착수" },
  { name: "⑤ Main Controller", desc: "선택된 Ticket → Mission 형태 변환 후 Executor 전달", progress: 10, color: TEAL, note: "설계 완료" },
  { name: "④ Agent Broker", desc: "여러 Agent Ticket 우선순위 선정 / Broker Policy Rules 적용", progress: 10, color: AMBER, note: "설계 완료" },
  { name: "③ Agent Layer", desc: "Load Balancing / Charging / MAPF Agent 판단 로직", progress: 10, color: PURPLE, note: "설계 완료" },
  { name: "② Service Layer", desc: "Robot/Mission/Traffic 상태 관리, 도메인 이벤트 발행 (Map Service ✅)", progress: 25, color: GREEN, note: "Map Service 완료" },
  { name: "① State Ingress Adapter", desc: "AGV/PLC state 수신·검증·정규화 → Service Layer 전달", progress: 20, color: BLUE, note: "기반 구축" },
];

sddLayers.forEach((l, i) => {
  const y = 1.58 + i * 0.58;
  s5.addText(l.name, { x: 0.7, y, w: 2.9, h: 0.35, fontSize: 9.5, fontFace: "Calibri", color: DARK, bold: true, valign: "middle", margin: 0 });
  s5.addText(l.desc, { x: 3.6, y, w: 3.8, h: 0.35, fontSize: 8.5, fontFace: "Calibri", color: GRAY, valign: "middle", margin: 0 });
  s5.addShape(pres.shapes.RECTANGLE, { x: 7.45, y: y + 0.07, w: 1.3, h: 0.2, fill: { color: "E2E8F0" } });
  if (l.progress > 0) s5.addShape(pres.shapes.RECTANGLE, { x: 7.45, y: y + 0.07, w: 1.3 * l.progress / 100, h: 0.2, fill: { color: l.color } });
  s5.addText(`${l.progress}%`, { x: 8.8, y, w: 0.4, h: 0.35, fontSize: 8, fontFace: "Calibri", color: DARK, bold: true, valign: "middle", margin: 0 });
  s5.addShape(pres.shapes.RECTANGLE, { x: 9.22, y: y + 0.07, w: 0.6, h: 0.22, fill: { color: "E8F0FE" } });
  s5.addText(l.note, { x: 9.22, y: y + 0.07, w: 0.6, h: 0.22, fontSize: 6.5, fontFace: "Calibri", color: l.color, bold: true, align: "center", valign: "middle", margin: 0 });
});

s5.addText("* State Ingress가 최하단(①), VDA5050 Executor가 최상단(⑥) — 그림상 흐름은 ① → ⑥ 방향", {
  x: 0.7, y: 4.85, w: 9, h: 0.25, fontSize: 7.5, fontFace: "Calibri", color: SLATE, margin: 0
});

// SLIDE 6: 다음 주 계획
let s6 = pres.addSlide();
addHeader(s6, "다음 주 계획", "2026.03.16 ~ 03.20");

const nextPlans = [
  { area: "ACS", color: GREEN, items: [
    "VDA5050 로봇 시뮬레이터 1대 주행 완전 검증",
    "경로 생성(A*/Dijkstra) 완료 및 테스트",
    "Agent Layer 구현 착수 (SDD v2 기반)",
    "Ticket 생성 → Agent Broker 연동 시작"
  ]},
  { area: "UI/UX", color: BLUE, items: [
    "노드/맵 편집 기능(U-01) 완료",
    "로봇 객체 그룹 관리 UI(U-03) 완료",
    "PLC 객체 DB 관리 UI(U-04) 착수",
    "DB 설계서 기반 acs_plcs 연동 개발"
  ]},
  { area: "PLC", color: AMBER, items: [
    "PLC 등록/수정/삭제 API 착수 (acs_plcs 테이블 기반)",
    "Polling TAG 수집 구현 (acs_plc_tags 기반)",
    "H-PLC 아키텍처 설계 완료 (W11)"
  ]},
];
nextPlans.forEach((p, i) => {
  const x = 0.5 + i * 3.17;
  s6.addShape(pres.shapes.RECTANGLE, { x, y: 1.0, w: 2.95, h: 4.2, fill: { color: WHITE }, shadow: makeShadow() });
  s6.addShape(pres.shapes.RECTANGLE, { x, y: 1.0, w: 2.95, h: 0.05, fill: { color: p.color } });
  s6.addShape(pres.shapes.RECTANGLE, { x: x + 0.15, y: 1.15, w: 0.65, h: 0.28, fill: { color: p.color } });
  s6.addText(p.area, { x: x + 0.15, y: 1.15, w: 0.65, h: 0.28, fontSize: 9, fontFace: "Calibri", color: WHITE, bold: true, align: "center", valign: "middle", margin: 0 });
  s6.addText("계획", { x: x + 0.9, y: 1.18, w: 1.8, h: 0.22, fontSize: 11, fontFace: "Calibri", color: DARK, bold: true, margin: 0 });
  const items = p.items.map(it => ({ text: it, options: { bullet: true, breakLine: true, fontSize: 9.5, color: DARK } }));
  s6.addText(items, { x: x + 0.15, y: 1.55, w: 2.65, h: 3.5, fontFace: "Calibri", paraSpaceAfter: 6, margin: 0 });
});

// SLIDE 7: 프로젝트 로드맵 현황
let s7 = pres.addSlide();
addHeader(s7, "프로젝트 로드맵 현황", "W11 기준  |  2026.03.14");

// 2x2 프로젝트 카드
const custProjCards = [
  {
    x:0.3, y:0.92, w:4.6, h:2.1,
    accent: "D97706",
    badge: "일정 변경",  badgeColor: "D97706",
    title: "2.1  Verification Toolkit",
    items: [
      "PLC 시뮬레이션: 4월 산출물 추가 예정",
      "명칭/방향: 시뮬레이션 → Verification Toolkit으로 정리",
      "10월 울산 OLT 일정 변경",
    ],
  },
  {
    x:5.1, y:0.92, w:4.6, h:2.1,
    accent: BLUE,
    badge: "논의중",  badgeColor: BLUE,
    title: "1.1  울산 C 프로젝트",
    items: [],
  },
  {
    x:0.3, y:3.1, w:4.6, h:2.15,
    accent: TEAL,
    badge: "준비중",  badgeColor: TEAL,
    title: "1.2  천안아산 프로젝트",
    items: [
      "로봇 2대 + 1대(?) 투입 예정 / 레이아웃 비교적 간단",
      "공사: 9월 추석 휴무 기간 활용 계획",
      "9월 목표: 미들웨어 전체 적용 (NaviCore 포함 검토)",
    ],
  },
  {
    x:5.1, y:3.1, w:4.6, h:2.15,
    accent: PURPLE,
    badge: "8월 통합테스트",  badgeColor: PURPLE,
    title: "1.3  EVO 화성 프로젝트",
    items: [
      "8월 통합 테스트 → 화성 공장에서 진행 예정",
      "9월 천안아산 적용 필수 (연계 일정)",
    ],
  },
];

custProjCards.forEach(p => {
  s7.addShape(pres.shapes.RECTANGLE, { x:p.x, y:p.y, w:p.w, h:p.h, fill:{color:WHITE}, shadow:makeShadow() });
  s7.addShape(pres.shapes.RECTANGLE, { x:p.x, y:p.y, w:p.w, h:0.05, fill:{color:p.accent} });
  // 배지
  s7.addShape(pres.shapes.RECTANGLE, { x:p.x+p.w-1.55, y:p.y+0.1, w:1.45, h:0.24, fill:{color:p.badgeColor} });
  s7.addText(p.badge, { x:p.x+p.w-1.55, y:p.y+0.1, w:1.45, h:0.24, fontSize:7, fontFace:"Calibri", color:WHITE, bold:true, align:"center", valign:"middle", margin:0 });
  // 제목
  s7.addText(p.title, { x:p.x+0.15, y:p.y+0.1, w:p.w-1.75, h:0.32, fontSize:11, fontFace:"Calibri", color:NAVY, bold:true, valign:"middle", margin:0 });
  // 구분선
  s7.addShape(pres.shapes.RECTANGLE, { x:p.x+0.15, y:p.y+0.47, w:p.w-0.3, h:0.02, fill:{color:"E2E8F0"} });
  // 항목
  const txtItems = p.items.map(it => ({ text:it, options:{ bullet:true, breakLine:true, fontSize:9.5, color:DARK } }));
  s7.addText(txtItems, { x:p.x+0.15, y:p.y+0.55, w:p.w-0.3, h:p.h-0.65, fontFace:"Calibri", paraSpaceAfter:5, margin:0 });
});

pres.writeFile({ fileName: OUTPUT }).then(() => {
  console.log("Customer PPTX saved:", OUTPUT);
}).catch(err => { console.error(err); process.exit(1); });
