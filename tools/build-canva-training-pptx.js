const pptxgen = require("pptxgenjs");
const React = require("react");
const RDS = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const { applyTheme } = require(process.env.SKILL + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "canva-teacher-training.pptx";

const THEME = {
  name: "Canva Classroom",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: "1F1640", lt1: "FFFFFF", dk2: "2B1055", lt2: "F3EEFF",
    accent1: "7D2AE8", accent2: "00A3AA", accent3: "FFB938", accent4: "F2557A",
    accent5: "2FA86B", accent6: "6B6680", hlink: "7D2AE8", folHlink: "5A1FB0",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9"; // 10 x 5.625
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "캔바(Canva) 활용 교사 연수";
pres.author = "교사 연수 담당";
const C = pres.SchemeColor;

// ---------- icons ----------
const iconCache = {};
async function icon(name, hex = "FFFFFF") {
  const key = name + hex;
  if (iconCache[key]) return iconCache[key];
  const Comp = fa[name];
  if (!Comp) throw new Error("missing icon " + name);
  const svg = RDS.renderToStaticMarkup(React.createElement(Comp, { color: "#" + hex, size: 256 }));
  const buf = await sharp(Buffer.from(svg)).resize(256, 256, { fit: "contain", background: { r: 0, g: 0, b: 0, alpha: 0 } }).png().toBuffer();
  return (iconCache[key] = "image/png;base64," + buf.toString("base64"));
}

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text2 },
  objects: [
    { ellipse: { x: 6.6, y: -1.2, w: 4.6, h: 4.6, fill: { color: C.accent1, transparency: 30 }, line: { type: "none" } } },
    { ellipse: { x: 7.9, y: 2.9, w: 3.2, h: 3.2, fill: { color: C.accent2, transparency: 35 }, line: { type: "none" } } },
    { ellipse: { x: 6.2, y: 3.7, w: 1.1, h: 1.1, fill: { color: C.accent3, transparency: 10 }, line: { type: "none" } } },
    { placeholder: { options: { name: "kicker", type: "body", x: 0.6, y: 1.05, w: 6.0, h: 0.4, fontSize: 14, bold: true, color: C.accent3, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.5, w: 6.4, h: 1.7, fontSize: 38, bold: true, color: C.background1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "subtitle", type: "body", x: 0.6, y: 3.3, w: 6.0, h: 0.6, fontSize: 16, color: C.background2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "info", type: "body", x: 0.6, y: 4.5, w: 6.0, h: 0.4, fontSize: 12, color: C.background2, margin: 0 }, text: "" } },
  ],
});

pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text2 },
  objects: [
    { ellipse: { x: 7.2, y: 1.0, w: 3.6, h: 3.6, fill: { color: C.accent1, transparency: 40 }, line: { type: "none" } } },
    { placeholder: { options: { name: "num", type: "body", x: 0.8, y: 1.15, w: 4, h: 0.9, fontSize: 48, bold: true, color: C.accent2, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 2.1, w: 6.6, h: 0.9, fontSize: 36, bold: true, color: C.background1, align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "desc", type: "body", x: 0.8, y: 3.1, w: 6.2, h: 1.0, fontSize: 16, color: C.background2, valign: "top", margin: 0 }, text: "" } },
  ],
  slideNumber: { x: 9.0, y: 5.2, w: 0.6, h: 0.3, fontSize: 10, color: C.background2, align: "right" },
});

pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: 0.5, y: 0.3, w: 9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, margin: 0 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: 0.5, y: 0.6, w: 9, h: 0.65, fontSize: 28, bold: true, color: C.text1, valign: "middle", align: "left", margin: 0 }, text: "" } },
    { text: { text: "캔바(Canva) 활용 교사 연수", options: { x: 0.5, y: 5.2, w: 5, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 8.9, y: 5.2, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
});

// ---------- helpers ----------
let curSection = "";
function section(title) { curSection = title; pres.addSection({ title }); }
function content(kicker, title, notes) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: curSection });
  s.addText(kicker, { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  if (notes) s.addNotes(notes);
  return s;
}
function divider(num, title, desc, notes) {
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: curSection });
  s.addText(num, { placeholder: "num" });
  s.addText(title, { placeholder: "title" });
  s.addText(desc, { placeholder: "desc" });
  if (notes) s.addNotes(notes);
  return s;
}
async function iconCircle(s, x, y, d, fill, name, iconHex = "FFFFFF") {
  s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: fill }, line: { type: "none" }, objectName: "icon-bg" });
  const p = d * 0.25;
  s.addImage({ data: await icon(name, iconHex), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p, objectName: "icon" });
}
function card(s, x, y, w, h, fill = C.background2) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { type: "none" }, rectRadius: 0.12, objectName: "card" });
}
function txt(s, text, o) { s.addText(text, { isTextBox: true, margin: 0, valign: "top", color: C.text1, ...o }); }
const ACC = [C.accent1, C.accent2, C.accent3, C.accent4, C.accent5, C.accent1];

// cards with icon at top: items [{icon, head, desc}]
async function cardGrid(s, items, cols, { x = 0.5, y = 1.55, w = 9, h = 3.35, gap = 0.3, d = 0.6, headSize = 16, descSize = 14 } = {}) {
  const rows = Math.ceil(items.length / cols);
  const cw = (w - gap * (cols - 1)) / cols;
  const ch = (h - gap * (rows - 1)) / rows;
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const cx = x + (i % cols) * (cw + gap);
    const cy = y + Math.floor(i / cols) * (ch + gap);
    card(s, cx, cy, cw, ch);
    await iconCircle(s, cx + 0.2, cy + 0.2, d, it.color || ACC[i % ACC.length], it.icon);
    if (rows > 1) {
      // compact: header beside icon
      txt(s, it.head, { x: cx + 0.3 + d, y: cy + 0.2, w: cw - d - 0.45, h: d, fontSize: headSize, bold: true, valign: "middle" });
      txt(s, it.desc, { x: cx + 0.2, y: cy + 0.3 + d, w: cw - 0.4, h: ch - d - 0.45, fontSize: descSize });
    } else {
      txt(s, it.head, { x: cx + 0.2, y: cy + 0.35 + d, w: cw - 0.4, h: 0.4, fontSize: headSize, bold: true });
      txt(s, it.desc, { x: cx + 0.2, y: cy + 0.8 + d, w: cw - 0.4, h: ch - d - 0.95, fontSize: descSize });
    }
  }
}

// icon rows: icon circle + head + desc
async function iconRows(s, items, { x, y, w, rowH = 0.85, d = 0.55, headSize = 16, descSize = 14 }) {
  for (let i = 0; i < items.length; i++) {
    const it = items[i];
    const ry = y + i * rowH;
    await iconCircle(s, x, ry, d, it.color || ACC[i % ACC.length], it.icon);
    txt(s, [
      { text: it.head, options: { bold: true, fontSize: headSize, breakLine: true } },
      { text: it.desc, options: { fontSize: descSize, color: C.accent6 } },
    ], { x: x + d + 0.2, y: ry - 0.03, w: w - d - 0.2, h: rowH - 0.05 });
  }
}

// numbered horizontal process
async function steps(s, items, { x = 0.5, y = 1.7, w = 9, h = 2.4, gap = 0.35 } = {}) {
  const n = items.length;
  const cw = (w - gap * (n - 1)) / n;
  for (let i = 0; i < n; i++) {
    const it = items[i];
    const cx = x + i * (cw + gap);
    card(s, cx, y, cw, h);
    s.addShape(pres.shapes.OVAL, { x: cx + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fill: { color: ACC[i % ACC.length] }, line: { type: "none" }, objectName: "step-num-bg" });
    s.addText(String(i + 1), { x: cx + 0.2, y: y + 0.2, w: 0.55, h: 0.55, fontSize: 18, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    txt(s, it.head, { x: cx + 0.2, y: y + 0.9, w: cw - 0.4, h: 0.45, fontSize: 16, bold: true });
    txt(s, it.desc, { x: cx + 0.2, y: y + 1.35, w: cw - 0.4, h: h - 1.45, fontSize: 13, color: C.accent6 });
    if (i < n - 1) {
      s.addImage({ data: await icon("FaArrowRight", HEX.accent6), x: cx + cw + gap / 2 - 0.1, y: y + h / 2 - 0.1, w: 0.2, h: 0.2, objectName: "arrow" });
    }
  }
}

function tipBox(s, text, { x = 0.5, y = 4.35, w = 9, h = 0.6, label = "TIP" } = {}) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: C.accent3, transparency: 75 }, line: { type: "none" }, rectRadius: 0.1, objectName: "tip" });
  txt(s, [
    { text: label + "  ", options: { bold: true, color: C.text2 } },
    { text, options: { color: C.text1 } },
  ], { x: x + 0.25, y, w: w - 0.5, h, fontSize: 14, valign: "middle" });
}

function chips(s, labels, { x, y, w, h = 0.4, gap = 0.15, fill = C.background2, color = C.accent1, fontSize = 13 }) {
  let cx = x, cy = y;
  labels.forEach((l) => {
    const latin = (l.match(/[\x00-\x7F]/g) || []).length;
    const lw = Math.max(0.7, 0.35 + (l.length - latin) * fontSize * 0.0145 + latin * fontSize * 0.0075);
    if (cx + lw > x + w) { cx = x; cy += h + gap; }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: cy, w: lw, h, fill: { color: fill }, line: { type: "none" }, rectRadius: h / 2, objectName: "chip" });
    s.addText(l, { x: cx, y: cy, w: lw, h, fontSize, bold: true, color, align: "center", valign: "middle", margin: 0, isTextBox: true });
    cx += lw + gap;
  });
}

(async () => {
  // ================= 1. Title =================
  section("시작");
  let s = pres.addSlide({ masterName: "TITLE", sectionTitle: curSection });
  s.addText("2026 교원 디지털 역량 강화 연수", { placeholder: "kicker" });
  s.addText("수업이 달라지는\n캔바(Canva) 활용법", { placeholder: "title" });
  s.addText("디자인 초보 선생님도 바로 쓰는 수업 자료 제작과 학생 참여 수업", { placeholder: "subtitle" });
  s.addText("연수 시간 약 120분  ·  실습 포함", { placeholder: "info" });
  await iconCircle(s, 8.1, 0.9, 1.6, C.background1, "FaPalette", HEX.accent1);
  s.addNotes("인사 및 연수 소개. 오늘은 캔바 계정 만들기부터 수업 자료 제작, 학생과 함께하는 협업 수업, AI 기능과 유의사항까지 다룹니다. 노트북과 계정을 미리 준비해 주세요.");

  // ================= 2. Goals =================
  s = content("INTRO", "오늘 연수가 끝나면 이렇게 됩니다",
    "연수 목표 3가지를 안내합니다. 각 목표는 뒤쪽 실습 미션과 연결됩니다.");
  await cardGrid(s, [
    { icon: "FaPaintBrush", head: "직접 만든다", desc: "템플릿을 활용해 수업용 슬라이드·학습지·포스터를 10분 안에 제작합니다." },
    { icon: "FaUsers", head: "함께 만든다", desc: "클래스를 개설하고 과제를 배포해 학생과 실시간으로 협업합니다." },
    { icon: "FaShieldAlt", head: "바르게 쓴다", desc: "AI 기능, 저작권, 개인정보 보호 원칙을 알고 안전하게 활용합니다." },
  ], 3, { h: 2.9 });

  // ================= 3. Agenda =================
  s = content("AGENDA", "연수 순서",
    "전체 흐름을 안내합니다. 이론 40분, 실습 60분, 질의응답 20분 정도로 운영합니다.");
  const agenda = [
    ["01", "캔바 이해하기", "교육용 무료 혜택과 가입·인증"],
    ["02", "기본 기능 익히기", "화면 구성, 편집, 단축키, 공유"],
    ["03", "수업 자료 만들기", "PPT·학습지·포스터·영상·화이트보드"],
    ["04", "학생 참여 수업", "클래스, 과제 배포, 교과별 아이디어"],
    ["05", "AI 기능과 유의사항", "매직 기능, 저작권, 개인정보"],
    ["06", "실습과 마무리", "실습 미션, 자주 묻는 질문"],
  ];
  agenda.forEach(([n, t, d], i) => {
    const cx = 0.5 + (i % 2) * 4.65, cy = 1.55 + Math.floor(i / 2) * 1.2;
    card(s, cx, cy, 4.35, 0.95);
    txt(s, n, { x: cx + 0.25, y: cy, w: 0.8, h: 0.95, fontSize: 28, bold: true, color: ACC[i], valign: "middle" });
    txt(s, [
      { text: t, options: { fontSize: 17, bold: true, breakLine: true } },
      { text: d, options: { fontSize: 13, color: C.accent6 } },
    ], { x: cx + 1.1, y: cy + 0.12, w: 3.1, h: 0.75, valign: "middle" });
  });

  // ================= PART 1 =================
  section("PART 1 캔바 이해하기");
  divider("PART 01", "캔바 이해하기", "캔바가 무엇인지, 왜 교실에서 쓰는지,\n교사 무료 계정을 어떻게 만드는지 알아봅니다.");

  // 5. What is Canva
  s = content("PART 1 · 캔바 이해하기", "캔바(Canva)란?",
    "캔바는 웹 브라우저와 앱에서 쓰는 온라인 디자인 도구입니다. 설치 없이 템플릿을 고르고 드래그만으로 디자인할 수 있다는 점이 핵심입니다.");
  txt(s, "누구나 쉽게 쓰는 온라인 디자인 플랫폼", { x: 0.5, y: 1.55, w: 4.6, h: 0.45, fontSize: 18, bold: true, color: C.accent1 });
  txt(s, [
    { text: "설치 없이 웹 브라우저·모바일 앱에서 바로 사용", options: { bullet: true, breakLine: true } },
    { text: "수많은 템플릿·사진·아이콘·폰트를 드래그로 편집", options: { bullet: true, breakLine: true } },
    { text: "프레젠테이션, 문서, 영상, 화이트보드까지 한 곳에서", options: { bullet: true, breakLine: true } },
    { text: "링크 하나로 공유하고 여러 명이 동시에 편집", options: { bullet: true } },
  ], { x: 0.5, y: 2.1, w: 4.6, h: 2.6, fontSize: 15, paraSpaceAfter: 8 });
  const stats = [["100+", "지원 언어 (한국어 포함)", C.accent1], ["0원", "K-12 교사·학생 교육용 요금", C.accent2], ["1개", "계정으로 PC·태블릿·스마트폰", C.accent4]];
  stats.forEach(([n, l, col], i) => {
    const cy = 1.55 + i * 1.12;
    card(s, 5.6, cy, 3.9, 0.95);
    txt(s, n, { x: 5.85, y: cy, w: 1.4, h: 0.95, fontSize: 30, bold: true, color: col, valign: "middle" });
    txt(s, l, { x: 7.25, y: cy, w: 2.1, h: 0.95, fontSize: 14, valign: "middle" });
  });

  // 6. Why in classroom
  s = content("PART 1 · 캔바 이해하기", "왜 교실에서 캔바인가?",
    "교사 입장에서는 자료 제작 시간 단축, 학생 입장에서는 표현 도구와 협업 경험이 가장 큰 장점입니다.");
  await cardGrid(s, [
    { icon: "FaClock", head: "제작 시간 단축", desc: "검증된 템플릿으로 시작해 자료 준비 시간을 크게 줄입니다." },
    { icon: "FaComments", head: "협업과 소통", desc: "모둠 전원이 하나의 디자인을 동시에 편집하고 댓글로 피드백합니다." },
    { icon: "FaLightbulb", head: "학생 표현력 향상", desc: "글쓰기만 어려운 학생도 이미지·영상으로 생각을 표현합니다." },
    { icon: "FaMobileAlt", head: "어디서나 접근", desc: "크롬북, 태블릿, 스마트폰 등 학교 기기 환경에 상관없이 사용합니다." },
  ], 2, { d: 0.55 });

  // 7. Canva for Education
  s = content("PART 1 · 캔바 이해하기", "Canva for Education: 교사·학생 무료",
    "초중고 교사 인증을 받으면 Canva for Education을 무료로 쓸 수 있습니다. 학생은 교사가 만든 클래스로 초대받아 함께 사용합니다. 세부 기능은 정책에 따라 바뀔 수 있으니 공식 페이지를 확인하세요.");
  const rows = [
    ["기능", "일반 무료", "교육용"],
    ["기본 템플릿·편집 기능", "●", "●"],
    ["프리미엄 템플릿·사진", "—", "●"],
    ["배경 제거 등 프로 기능", "—", "●"],
    ["클래스·과제·피드백", "—", "●"],
    ["구글 클래스룸 등 연동", "—", "●"],
  ];
  s.addTable(rows.map((r, ri) => r.map((c, ci) => ({
    text: c,
    options: {
      bold: ri === 0 || (ci === 0),
      color: ri === 0 ? HEX.lt1 : (ci === 2 && c === "●" ? HEX.accent1 : (ci === 0 ? HEX.dk1 : HEX.accent6)),
      fill: { color: ri === 0 ? HEX.dk2 : (ri % 2 ? HEX.lt1 : HEX.lt2) },
      align: ci === 0 ? "left" : "center",
      fontSize: 14, valign: "middle",
    },
  }))), { x: 0.5, y: 1.55, w: 6.0, colW: [3.2, 1.4, 1.4], rowH: 0.5, border: { type: "none" }, margin: [0, 0.12, 0, 0.12] });
  card(s, 6.85, 1.55, 2.65, 3.0, C.text2);
  await iconCircle(s, 7.05, 1.8, 0.6, C.accent3, "FaGift", HEX.dk2);
  txt(s, "핵심 포인트", { x: 7.05, y: 2.55, w: 2.3, h: 0.4, fontSize: 16, bold: true, color: C.background1 });
  txt(s, "교사 인증 한 번으로 우리 반 학생 모두 프리미엄 기능을 무료로 사용", { x: 7.05, y: 3.0, w: 2.3, h: 1.4, fontSize: 14, color: C.background2 });

  // 8. Signup
  s = content("PART 1 · 캔바 이해하기", "가입과 교사 인증 4단계",
    "인증 심사에 시간이 걸릴 수 있으니 연수 전에 미리 신청하도록 안내합니다. 승인되지 않은 경우 일반 무료 계정으로 실습해도 대부분의 기능은 체험할 수 있습니다.");
  await steps(s, [
    { head: "접속", desc: "canva.com/education 접속 후 ‘교사’로 시작하기" },
    { head: "계정 만들기", desc: "학교 이메일 또는 구글 계정으로 가입" },
    { head: "재직 증빙", desc: "재직증명서 등 교사 확인 서류 업로드" },
    { head: "승인·사용", desc: "승인 메일 확인 후 교육용 기능 활성화" },
  ], { h: 2.45 });
  tipBox(s, "승인까지 시간이 걸릴 수 있어요. 연수 며칠 전에 미리 신청하세요.", { y: 4.4, h: 0.55 });

  // ================= PART 2 =================
  section("PART 2 기본 기능 익히기");
  divider("PART 02", "기본 기능 익히기", "화면 구성, 템플릿 검색, 편집 기초, 단축키,\n다운로드와 공유 방법을 익힙니다.");

  // 10. Editor layout mock
  s = content("PART 2 · 기본 기능 익히기", "편집 화면 한눈에 보기",
    "실제 캔바 편집 화면을 띄우고 번호 순서대로 위치를 짚어 줍니다. 왼쪽 사이드바가 거의 모든 기능의 출발점입니다.");
  const mx = 0.5, my = 1.55, mw = 5.0, mh = 3.4;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: mx, y: my, w: mw, h: mh, fill: { color: C.background2 }, line: { color: C.accent6, width: 0.75 }, rectRadius: 0.08, objectName: "mock-frame" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.05, y: my + 0.05, w: mw - 0.1, h: 0.4, fill: { color: C.text2 }, line: { type: "none" }, objectName: "mock-top" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.05, y: my + 0.5, w: 0.45, h: mh - 0.55, fill: { color: C.background1 }, line: { type: "none" }, objectName: "mock-rail" });
  for (let i = 0; i < 6; i++) s.addShape(pres.shapes.OVAL, { x: mx + 0.15, y: my + 0.65 + i * 0.42, w: 0.25, h: 0.25, fill: { color: C.accent6, transparency: 50 }, line: { type: "none" }, objectName: "mock-rail-icon" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 0.55, y: my + 0.5, w: 1.15, h: mh - 0.55, fill: { color: C.background1, transparency: 30 }, line: { type: "none" }, objectName: "mock-panel" });
  for (let i = 0; i < 4; i++) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: mx + 0.65, y: my + 0.65 + i * 0.6, w: 0.95, h: 0.48, fill: { color: [C.accent1, C.accent2, C.accent3, C.accent4][i], transparency: 55 }, line: { type: "none" }, rectRadius: 0.05, objectName: "mock-thumb" });
  s.addShape(pres.shapes.RECTANGLE, { x: mx + 1.95, y: my + 0.75, w: 2.8, h: 1.6, fill: { color: C.background1 }, line: { type: "none" }, shadow: { type: "outer", color: "000000", opacity: 0.18, blur: 4, offset: 1, angle: 90 }, objectName: "mock-canvas" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: mx + 2.15, y: my + 0.95, w: 1.6, h: 0.22, fill: { color: C.accent1, transparency: 30 }, line: { type: "none" }, rectRadius: 0.05, objectName: "mock-heading" });
  s.addShape(pres.shapes.OVAL, { x: mx + 3.85, y: my + 1.3, w: 0.7, h: 0.7, fill: { color: C.accent3, transparency: 20 }, line: { type: "none" }, objectName: "mock-shape" });
  for (let i = 0; i < 3; i++) s.addShape(pres.shapes.RECTANGLE, { x: mx + 2.15 + i * 0.95, y: my + 2.65, w: 0.8, h: 0.45, fill: { color: C.background1 }, line: { color: i === 0 ? C.accent1 : C.background2, width: 1.5 }, objectName: "mock-page" });
  const badge = (n, x, y) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: 0.36, h: 0.36, fill: { color: C.accent4 }, line: { color: C.background1, width: 1.5 }, objectName: "badge" });
    s.addText(String(n), { x, y, w: 0.36, h: 0.36, fontSize: 12, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  };
  badge(1, mx + 4.45, my + 0.07); badge(2, mx + 0.3, my + 0.55); badge(3, mx + 4.55, my + 0.6); badge(4, mx + 4.55, my + 2.7);
  const legend = [
    ["상단 메뉴", "파일 · 크기 조정 · 공유 · 다운로드"],
    ["사이드바", "디자인 · 요소 · 텍스트 · 업로드"],
    ["편집 캔버스", "클릭·드래그로 바로 편집"],
    ["페이지 영역", "페이지 추가 · 순서 변경 · 메모"],
  ];
  legend.forEach(([h, d], i) => {
    const ly = 1.6 + i * 0.85;
    s.addShape(pres.shapes.OVAL, { x: 5.85, y: ly, w: 0.4, h: 0.4, fill: { color: C.accent4 }, line: { type: "none" }, objectName: "legend-num" });
    s.addText(String(i + 1), { x: 5.85, y: ly, w: 0.4, h: 0.4, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    txt(s, [{ text: h, options: { bold: true, fontSize: 16, breakLine: true } }, { text: d, options: { fontSize: 13, color: C.accent6 } }], { x: 6.4, y: ly - 0.04, w: 3.1, h: 0.75 });
  });

  // 11. Templates
  s = content("PART 2 · 기본 기능 익히기", "템플릿, 검색어가 반이다",
    "검색창에 학년·교과·용도를 함께 입력하면 원하는 템플릿을 빠르게 찾을 수 있습니다. 영어 키워드로도 검색해 보세요.");
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.55, w: 9, h: 0.65, fill: { color: C.background1 }, line: { color: C.accent1, width: 2 }, rectRadius: 0.32, objectName: "search-bar" });
  s.addImage({ data: await icon("FaSearch", HEX.accent1), x: 0.8, y: 1.7, w: 0.35, h: 0.35, objectName: "search-icon" });
  txt(s, "중학교 과학 학습지", { x: 1.35, y: 1.55, w: 6, h: 0.65, fontSize: 18, valign: "middle" });
  txt(s, "추천 검색어", { x: 0.5, y: 2.4, w: 1.4, h: 0.4, fontSize: 13, bold: true, color: C.accent6, valign: "middle" });
  chips(s, ["학습지", "수업 PPT", "학급 규칙", "독서 기록장", "worksheet", "infographic"], { x: 1.9, y: 2.4, w: 7.6 });
  await cardGrid(s, [
    { icon: "FaFilter", head: "필터 활용", desc: "스타일·색상·주제 필터로 범위를 좁힙니다." },
    { icon: "FaExchangeAlt", head: "그대로 말고 바꿔서", desc: "글꼴·색만 바꿔도 우리 반 자료가 됩니다." },
    { icon: "FaHeart", head: "즐겨찾기", desc: "좋은 템플릿은 별표로 저장해 두세요." },
  ], 3, { y: 3.05, h: 1.85, d: 0.5, descSize: 13 });

  // 12. Edit basics
  s = content("PART 2 · 기본 기능 익히기", "편집 기초 6가지",
    "각 기능을 직접 시연합니다. 특히 ‘위치’ 메뉴의 정렬·레이어 기능을 알면 디자인이 훨씬 깔끔해집니다.");
  await cardGrid(s, [
    { icon: "FaFont", head: "텍스트", desc: "제목·본문 추가, 글꼴과 효과 적용" },
    { icon: "FaShapes", head: "요소", desc: "도형·아이콘·스티커·표·차트 삽입" },
    { icon: "FaImage", head: "사진", desc: "무료 사진 검색, 프레임에 끌어 넣기" },
    { icon: "FaCloudUploadAlt", head: "업로드", desc: "내 사진·영상·PDF 불러와 편집" },
    { icon: "FaLayerGroup", head: "위치·정렬", desc: "앞뒤 순서, 가운데 정렬, 간격 맞춤" },
    { icon: "FaPlay", head: "애니메이션", desc: "등장 효과와 페이지 전환 설정" },
  ], 3, { d: 0.5 });

  // 13. Shortcuts
  s = content("PART 2 · 기본 기능 익히기", "알아두면 편한 단축키",
    "Windows 기준이며 Mac은 Ctrl 대신 Cmd를 사용합니다. 슬래시(/) 키는 빠른 작업 메뉴를 엽니다.");
  const keys = [
    ["T", "텍스트 상자 추가"], ["R / C / L", "사각형 / 원 / 선 추가"],
    ["Ctrl + D", "선택한 요소 복제"], ["Ctrl + G", "그룹으로 묶기"],
    ["Shift + 클릭", "여러 요소 동시 선택"], ["Ctrl + Z", "실행 취소"],
  ];
  keys.forEach(([k, d], i) => {
    const cx = 0.5 + (i % 2) * 3.0, cy = 1.55 + Math.floor(i / 2) * 1.15;
    card(s, cx, cy, 2.75, 0.95);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.2, y: cy + 0.15, w: 2.35, h: 0.38, fill: { color: C.background1 }, line: { color: C.accent1, width: 1 }, rectRadius: 0.06, objectName: "key" });
    s.addText(k, { x: cx + 0.2, y: cy + 0.15, w: 2.35, h: 0.38, fontSize: 14, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: "Arial" });
    txt(s, d, { x: cx + 0.2, y: cy + 0.57, w: 2.35, h: 0.32, fontSize: 13, align: "center" });
  });
  card(s, 6.7, 1.55, 2.8, 3.25, C.text2);
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.55, y: 1.85, w: 1.1, h: 1.1, fill: { color: C.background1 }, line: { type: "none" }, rectRadius: 0.15, objectName: "slash-key" });
  s.addText("/", { x: 7.55, y: 1.85, w: 1.1, h: 1.1, fontSize: 44, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: "Arial" });
  txt(s, "빠른 작업 메뉴", { x: 6.9, y: 3.15, w: 2.4, h: 0.4, fontSize: 16, bold: true, color: C.background1, align: "center" });
  txt(s, "편집 화면에서 / 를 누르면 요소 추가·AI 기능을 바로 검색", { x: 6.95, y: 3.6, w: 2.3, h: 1.0, fontSize: 13, color: C.background2, align: "center" });

  // 14. Download & share
  s = content("PART 2 · 기본 기능 익히기", "다운로드와 공유",
    "인쇄용 학습지는 ‘PDF 인쇄’, 영상은 MP4를 선택합니다. 학생에게 나눠 줄 때는 원본이 훼손되지 않도록 ‘템플릿 링크’를 꼭 활용하세요.");
  card(s, 0.5, 1.55, 4.35, 3.35);
  await iconCircle(s, 0.75, 1.8, 0.55, C.accent2, "FaDownload");
  txt(s, "다운로드 형식", { x: 1.45, y: 1.8, w: 3.2, h: 0.55, fontSize: 18, bold: true, valign: "middle" });
  chips(s, ["PNG", "JPG", "PDF 표준", "PDF 인쇄", "MP4", "GIF", "PPTX"], { x: 0.75, y: 2.6, w: 3.9, fill: C.background1, color: C.accent2 });
  txt(s, "인쇄물은 PDF 인쇄, 영상은 MP4, 한글·파워포인트 편집이 필요하면 PPTX", { x: 0.75, y: 3.75, w: 3.9, h: 0.95, fontSize: 13, color: C.accent6 });
  card(s, 5.15, 1.55, 4.35, 3.35);
  await iconCircle(s, 5.4, 1.8, 0.55, C.accent1, "FaShareAlt");
  txt(s, "공유 방법", { x: 6.1, y: 1.8, w: 3.2, h: 0.55, fontSize: 18, bold: true, valign: "middle" });
  const shareItems = [["보기 전용 링크", "학부모·학생 안내용"], ["편집 권한 링크", "동료 교사와 공동 제작"], ["템플릿 링크 ★", "학생마다 자기 사본 생성"]];
  shareItems.forEach(([h, d], i) => {
    const ry = 2.6 + i * 0.72;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.4, y: ry, w: 3.85, h: 0.6, fill: { color: i === 2 ? C.accent1 : C.background1 }, line: { type: "none" }, rectRadius: 0.08, objectName: "share-row" });
    txt(s, [{ text: h + "   ", options: { bold: true } }, { text: d, options: { color: i === 2 ? C.background2 : C.accent6 } }], { x: 5.6, y: ry, w: 3.6, h: 0.6, fontSize: 13, valign: "middle", color: i === 2 ? C.background1 : C.text1 });
  });

  // ================= PART 3 =================
  section("PART 3 수업 자료 만들기");
  divider("PART 03", "수업 자료 만들기", "프레젠테이션, 학습지, 포스터, 영상, 화이트보드까지\n수업에 바로 쓰는 다섯 가지 자료를 만들어 봅니다.");

  // 16. Presentation
  s = content("PART 3 · 수업 자료 만들기", "수업용 프레젠테이션",
    "캔바 프레젠테이션은 발표자 보기, 발표 녹화 기능이 있어 원격 수업이나 거꾸로 수업 영상 제작에도 활용됩니다.");
  card(s, 0.5, 1.55, 4.0, 3.35, C.text2);
  txt(s, "제작 순서", { x: 0.8, y: 1.8, w: 3.4, h: 0.4, fontSize: 16, bold: true, color: C.accent3 });
  txt(s, [
    { text: "‘프레젠테이션’ 템플릿 선택", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "학습 목표 → 활동 → 정리 순으로 페이지 구성", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "글은 짧게, 이미지는 크게", options: { bullet: { type: "number" }, breakLine: true } },
    { text: "애니메이션은 핵심 장면에만", options: { bullet: { type: "number" } } },
  ], { x: 0.8, y: 2.3, w: 3.5, h: 2.4, fontSize: 15, color: C.background1, paraSpaceAfter: 8 });
  await iconRows(s, [
    { icon: "FaDesktop", head: "발표자 보기", desc: "다음 슬라이드와 메모를 보면서 수업 진행" },
    { icon: "FaMicrophone", head: "발표 녹화", desc: "화면과 목소리를 녹화해 복습 영상으로 공유" },
    { icon: "FaFilePowerpoint" in fa ? "FaFilePowerpoint" : "FaFileAlt", head: "PPTX 내보내기", desc: "기존 파워포인트 환경에서도 그대로 사용" },
  ], { x: 4.9, y: 1.65, w: 4.6, rowH: 1.1 });

  // 17. Worksheet
  s = content("PART 3 · 수업 자료 만들기", "학습지·워크시트",
    "A4 문서 크기로 시작해 인쇄용 PDF로 내보냅니다. 같은 학습지를 디지털로 배포하면 학생이 캔바에서 바로 작성할 수도 있습니다.");
  await iconRows(s, [
    { icon: "FaFileAlt", head: "A4 문서로 시작", desc: "‘A4 문서’ 또는 ‘학습지’ 템플릿 선택" },
    { icon: "FaTasks", head: "표·체크리스트 요소", desc: "요소 메뉴에서 표, 체크박스, 빈칸 상자 추가" },
    { icon: "FaPrint", head: "PDF 인쇄로 저장", desc: "선명한 인쇄를 위해 ‘PDF 인쇄’ 형식 선택" },
    { icon: "FaRegCopy", head: "디지털 학습지로 배포", desc: "템플릿 링크로 공유하면 학생이 직접 작성" },
  ], { x: 0.5, y: 1.6, w: 5.2, rowH: 0.85 });
  // worksheet mock
  s.addShape(pres.shapes.RECTANGLE, { x: 6.5, y: 1.5, w: 2.6, h: 3.45, fill: { color: C.background1 }, line: { color: C.background2, width: 1 }, shadow: { type: "outer", color: "000000", opacity: 0.2, blur: 6, offset: 2, angle: 90 }, objectName: "sheet" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.7, y: 1.7, w: 2.2, h: 0.45, fill: { color: C.accent2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "sheet-head" });
  s.addText("관찰 기록지", { x: 6.7, y: 1.7, w: 2.2, h: 0.45, fontSize: 13, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
  for (let i = 0; i < 4; i++) {
    s.addShape(pres.shapes.RECTANGLE, { x: 6.75, y: 2.4 + i * 0.45, w: 0.2, h: 0.2, fill: { color: C.background1 }, line: { color: C.accent6, width: 1 }, objectName: "sheet-box" });
    s.addShape(pres.shapes.LINE, { x: 7.1, y: 2.5 + i * 0.45, w: 1.75, h: 0, line: { color: C.background2, width: 1.5 }, objectName: "sheet-line" });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.75, y: 4.2, w: 2.1, h: 0.55, fill: { color: C.background2 }, line: { type: "none" }, rectRadius: 0.06, objectName: "sheet-note" });

  // 18. Posters
  s = content("PART 3 · 수업 자료 만들기", "포스터·인포그래픽·카드뉴스",
    "교사용 게시물뿐 아니라 학생 결과물로도 자주 쓰입니다. 정보 정리 능력과 시각적 표현을 동시에 기를 수 있습니다.");
  await cardGrid(s, [
    { icon: "FaBullhorn", head: "학급 게시물", desc: "학급 규칙, 행사 안내, 가정통신 이미지" },
    { icon: "FaChartBar", head: "인포그래픽", desc: "조사 결과를 차트·아이콘으로 한 장에 정리" },
    { icon: "FaBook", head: "카드뉴스", desc: "독서 감상, 진로 탐색을 5~6장 카드로 구성" },
  ], 3, { h: 2.55 });
  tipBox(s, "‘매직 스위치(크기 조정)’로 포스터를 카드뉴스·SNS 크기로 한 번에 변환할 수 있어요.", { y: 4.35, h: 0.6 });

  // 19. Video
  s = content("PART 3 · 수업 자료 만들기", "동영상 만들기",
    "캔바 동영상 편집기는 타임라인 기반이라 초보자도 쉽게 다룹니다. 학생 프로젝트 발표 영상, 수업 안내 영상 등에 활용합니다.");
  await steps(s, [
    { head: "소스 준비", desc: "촬영 영상 업로드 또는 무료 영상·사진 검색" },
    { head: "컷 편집", desc: "타임라인에서 자르고 순서 바꾸기" },
    { head: "자막·음악", desc: "텍스트 자막, 배경음악, 전환 효과 추가" },
    { head: "MP4 저장", desc: "다운로드 후 클래스룸·패들렛에 공유" },
  ], { h: 2.45 });
  tipBox(s, "학생 얼굴이 나오는 영상은 공개 범위를 반드시 ‘링크가 있는 사람’ 이하로 제한하세요.", { y: 4.4, h: 0.55, label: "주의" });

  // 20. Whiteboard
  s = content("PART 3 · 수업 자료 만들기", "화이트보드로 함께 생각하기",
    "화이트보드는 끝없이 넓어지는 캔버스입니다. 모둠별 영역을 미리 나누어 두면 활동이 혼란스럽지 않습니다.");
  card(s, 0.5, 1.55, 3.6, 3.35, C.accent1);
  await iconCircle(s, 0.8, 1.85, 0.7, C.background1, "FaChalkboard", HEX.accent1);
  txt(s, "무한 캔버스에서\n반 전체가 동시에", { x: 0.8, y: 2.75, w: 3.1, h: 1.0, fontSize: 22, bold: true, color: C.background1 });
  txt(s, "브레인스토밍 · 모둠 토의 · 생각 정리", { x: 0.8, y: 3.95, w: 3.1, h: 0.6, fontSize: 14, color: C.background2 });
  await iconRows(s, [
    { icon: "FaStickyNote", head: "포스트잇", desc: "학생마다 색을 정해 의견을 붙이기" },
    { icon: "FaThLarge" in fa ? "FaThLarge" : "FaLayerGroup", head: "모둠별 영역", desc: "도형으로 구역을 나누고 모둠 이름 표시" },
    { icon: "FaClock", head: "타이머", desc: "활동 시간을 화면에 띄워 집중 유도" },
    { icon: "FaVoteYea", head: "스티커 투표", desc: "좋은 아이디어에 스티커로 투표" },
  ], { x: 4.5, y: 1.6, w: 5.0, rowH: 0.83 });

  // ================= PART 4 =================
  section("PART 4 학생 참여 수업");
  divider("PART 04", "학생 참여 수업", "클래스를 만들고 과제를 배포해\n학생과 함께 만드는 수업을 설계합니다.");

  // 22. Class
  s = content("PART 4 · 학생 참여 수업", "클래스 만들고 학생 초대하기",
    "학생은 교사의 초대를 통해 교육용 계정을 받습니다. 구글 클래스룸 등 이미 쓰는 LMS가 있다면 연동하는 것이 가장 편합니다.");
  await steps(s, [
    { head: "클래스 생성", desc: "홈 화면 ‘사람’ 메뉴에서 새 클래스 만들기" },
    { head: "학생 초대", desc: "초대 링크·이메일 또는 LMS로 학생 추가" },
    { head: "모둠 구성", desc: "필요 시 그룹을 만들어 폴더·디자인 공유" },
  ], { h: 2.0 });
  txt(s, "연동 가능한 LMS", { x: 0.5, y: 3.95, w: 2.0, h: 0.4, fontSize: 14, bold: true, color: C.accent6, valign: "middle" });
  chips(s, ["Google Classroom", "Microsoft Teams", "Canvas", "Schoology", "Moodle"], { x: 2.4, y: 3.95, w: 7.1, fontSize: 12 });
  txt(s, "※ 메뉴 이름과 위치는 캔바 업데이트에 따라 달라질 수 있습니다.", { x: 0.5, y: 4.7, w: 9, h: 0.3, fontSize: 11, color: C.accent6 });

  // 23. Assign & feedback cycle
  s = content("PART 4 · 학생 참여 수업", "과제 배포부터 피드백까지",
    "교사가 만든 템플릿을 과제로 배포하면 학생마다 사본이 생성됩니다. 제출된 결과물에 댓글로 피드백하고, 학생이 수정해 다시 제출하는 순환 구조를 만듭니다.");
  const cyc = [
    ["FaPaintBrush", "템플릿 제작", "교사"],
    ["FaShareAlt", "과제로 배포", "교사"],
    ["FaUsers", "사본에서 작업", "학생"],
    ["FaClipboardCheck", "제출", "학생"],
    ["FaComments", "댓글 피드백", "교사"],
  ];
  for (let i = 0; i < cyc.length; i++) {
    const [ic, h, who] = cyc[i];
    const cx = 0.5 + i * 1.86;
    const teacher = who === "교사";
    card(s, cx, 1.65, 1.56, 2.2, teacher ? C.background2 : C.background1);
    if (!teacher) s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx, y: 1.65, w: 1.56, h: 2.2, fill: { type: "none" }, line: { color: C.accent2, width: 1.5 }, rectRadius: 0.12, objectName: "student-outline" });
    await iconCircle(s, cx + 0.45, 1.85, 0.66, teacher ? C.accent1 : C.accent2, ic);
    txt(s, h, { x: cx + 0.1, y: 2.65, w: 1.36, h: 0.45, fontSize: 15, bold: true, align: "center" });
    txt(s, who, { x: cx + 0.1, y: 3.15, w: 1.36, h: 0.35, fontSize: 12, bold: true, color: teacher ? C.accent1 : C.accent2, align: "center" });
    if (i < cyc.length - 1) s.addImage({ data: await icon("FaArrowRight", HEX.accent6), x: cx + 1.6, y: 2.62, w: 0.22, h: 0.22, objectName: "arrow" });
  }
  tipBox(s, "피드백 후 학생이 수정해 다시 제출하는 ‘수정 → 재제출’ 순환을 1회 이상 넣어 보세요.", { y: 4.25, h: 0.65 });

  // 24. Subject ideas
  s = content("PART 4 · 학생 참여 수업", "교과별 활용 아이디어",
    "교과별 예시입니다. 참가 선생님들께 자기 교과에서 쓸 아이디어를 하나씩 말해 보게 하면 좋습니다.");
  await cardGrid(s, [
    { icon: "FaBook", head: "국어", desc: "책 표지 다시 만들기, 인물 SNS 프로필" },
    { icon: "FaCalculator", head: "수학", desc: "통계 인포그래픽, 도형 개념 카드" },
    { icon: "FaGlobeAsia", head: "사회", desc: "지역 홍보 포스터, 역사 타임라인" },
    { icon: "FaFlask", head: "과학", desc: "실험 보고서, 생태계 관계도" },
    { icon: "FaLanguage", head: "영어", desc: "자기소개 영상, 여행 브로슈어" },
    { icon: "FaPalette", head: "미술·창체", desc: "캠페인 포스터, 진로 명함 디자인" },
  ], 3, { d: 0.5 });

  // 25. Lesson timeline
  s = content("PART 4 · 학생 참여 수업", "40분 수업 흐름 예시",
    "예시는 중학교 사회 ‘우리 지역 홍보 포스터’ 수업입니다. 막대 길이는 시간 비율을 나타냅니다.");
  const tl = [["도입", 5, "목표 확인, 우수 예시 감상"], ["전개 ①", 10, "모둠별 자료 조사, 역할 분담"], ["전개 ②", 17, "캔바 템플릿으로 포스터 공동 제작"], ["정리", 8, "갤러리 워크, 댓글 피드백"]];
  let tx = 0.5;
  const unit = 9 / 40;
  tl.forEach(([h, m, d], i) => {
    const w = m * unit;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: tx + 0.04, y: 1.7, w: w - 0.08, h: 0.75, fill: { color: ACC[i] }, line: { type: "none" }, rectRadius: 0.1, objectName: "timeline-bar" });
    s.addText(`${m}분`, { x: tx, y: 1.7, w, h: 0.75, fontSize: 15, bold: true, color: i === 2 ? C.text1 : C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
    tx += w;
  });
  tl.forEach(([h, m, d], i) => {
    const cx = 0.5 + i * 2.325;
    card(s, cx, 2.8, 2.1, 1.7);
    txt(s, h, { x: cx + 0.2, y: 2.95, w: 1.7, h: 0.45, fontSize: 17, bold: true, color: ACC[i] === C.accent3 ? C.text2 : ACC[i] });
    txt(s, d, { x: cx + 0.2, y: 3.45, w: 1.7, h: 1.3, fontSize: 14 });
  });

  // ================= PART 5 =================
  section("PART 5 AI 기능과 유의사항");
  divider("PART 05", "AI 기능과 유의사항", "캔바의 AI(매직) 기능을 살펴보고\n저작권·개인정보 보호 원칙을 확인합니다.");

  // 27. AI tools
  s = content("PART 5 · AI 기능과 유의사항", "캔바 AI(매직) 도구 6가지",
    "기능 이름과 제공 범위는 자주 바뀝니다. 학생 계정의 AI 기능은 관리자 또는 교사 설정에 따라 제한될 수 있으니 수업 전 확인하세요.");
  await cardGrid(s, [
    { icon: "FaMagic", head: "매직 디자인", desc: "주제를 입력하면 디자인 초안 자동 생성" },
    { icon: "FaPenFancy" in fa ? "FaPenFancy" : "FaFont", head: "매직 라이팅", desc: "글 초안 작성, 요약, 문체 바꾸기" },
    { icon: "FaCut", head: "배경 제거", desc: "클릭 한 번으로 사진 배경 지우기" },
    { icon: "FaEraser", head: "매직 지우개", desc: "사진 속 불필요한 물체 지우기" },
    { icon: "FaExpand", head: "매직 확장", desc: "사진 바깥 영역을 자연스럽게 확장" },
    { icon: "FaLanguage", head: "번역", desc: "디자인 전체를 다른 언어로 번역" },
  ], 3, { d: 0.5 });

  // 28. Copyright & privacy
  s = content("PART 5 · AI 기능과 유의사항", "저작권과 개인정보, 이것만은 꼭",
    "캔바 소스는 캔바 이용 약관 범위 안에서 사용할 수 있습니다. 상업적 이용, 로고·상표 사용은 별도 확인이 필요합니다.");
  const colItems = [
    ["FaCopyright", "저작권", C.accent1, ["캔바 제공 소스는 약관 범위 안에서 사용", "외부 이미지·음원은 출처와 이용 허락 확인", "AI 생성물임을 결과물에 표시하도록 지도", "학생 작품의 저작권은 학생에게 있음을 안내"]],
    ["FaUserShield", "개인정보", C.accent4, ["학생 실명·얼굴 사진 공개 링크 공유 금지", "학교 계정으로 클래스 운영, 개인 계정 지양", "14세 미만은 학교 방침과 보호자 동의 확인", "학기 종료 후 불필요한 클래스·자료 정리"]],
  ];
  for (let i = 0; i < 2; i++) {
    const [ic, h, col, list] = colItems[i];
    const cx = 0.5 + i * 4.65;
    card(s, cx, 1.55, 4.35, 3.35);
    await iconCircle(s, cx + 0.25, 1.8, 0.55, col, ic);
    txt(s, h, { x: cx + 0.95, y: 1.8, w: 3.2, h: 0.55, fontSize: 18, bold: true, valign: "middle" });
    txt(s, list.map((t, j) => ({ text: t, options: { bullet: true, breakLine: j < list.length - 1 } })), { x: cx + 0.25, y: 2.55, w: 3.9, h: 2.2, fontSize: 14, paraSpaceAfter: 6 });
  }

  // ================= PART 6 =================
  section("PART 6 실습과 마무리");
  divider("PART 06", "실습과 마무리", "직접 만들어 보고,\n자주 묻는 질문으로 정리합니다.");

  // 30. Missions
  s = content("PART 6 · 실습과 마무리", "실습 미션",
    "미션 1은 필수, 2·3은 선택입니다. 완성작은 연수 클래스에 제출해 서로 감상합니다.");
  const missions = [
    ["MISSION 1", "15분", "수업 PPT 5장 만들기", "템플릿 검색 → 글꼴·색 변경 → 사진·아이콘 추가", C.accent1],
    ["MISSION 2", "15분", "디지털 학습지 배포", "A4 학습지 제작 → 템플릿 링크 만들기 → 동료와 교환", C.accent2],
    ["MISSION 3", "20분", "화이트보드 협업 체험", "모둠 영역 나누기 → 포스트잇 의견 → 스티커 투표", C.accent4],
  ];
  missions.forEach(([m, t, h, d, col], i) => {
    const cx = 0.5 + i * 3.1;
    card(s, cx, 1.55, 2.8, 3.35);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: cx + 0.2, y: 1.75, w: 1.35, h: 0.38, fill: { color: col }, line: { type: "none" }, rectRadius: 0.19, objectName: "mission-tag" });
    s.addText(m, { x: cx + 0.2, y: 1.75, w: 1.35, h: 0.38, fontSize: 11, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: "Arial" });
    txt(s, t, { x: cx + 1.7, y: 1.75, w: 0.9, h: 0.38, fontSize: 13, bold: true, color: C.accent6, align: "right", valign: "middle" });
    txt(s, h, { x: cx + 0.2, y: 2.35, w: 2.4, h: 0.85, fontSize: 18, bold: true });
    txt(s, d, { x: cx + 0.2, y: 3.3, w: 2.4, h: 1.45, fontSize: 14, color: C.accent6 });
  });

  // 31. FAQ
  s = content("PART 6 · 실습과 마무리", "자주 묻는 질문",
    "연수에서 자주 나오는 질문입니다. 정책 관련 답변은 캔바 공식 도움말 센터에서 최신 내용을 확인하세요.");
  const faq = [
    ["교사 인증이 거절됐어요", "재직을 확인할 수 있는 서류를 다시 제출하거나 도움말 센터에 문의하세요."],
    ["학생도 가입해야 하나요?", "교사가 보낸 클래스 초대로 참여하면 교육용 기능을 함께 씁니다."],
    ["한글 글꼴이 깨져요", "캔바 기본 한글 글꼴을 쓰고, 공유는 PDF로 하면 안전합니다."],
    ["인터넷이 느린 교실이라면?", "자료는 미리 PDF·MP4로 내려받아 두고, 협업 활동은 모둠당 1대로 운영하세요."],
  ];
  for (let i = 0; i < faq.length; i++) {
    const [q, a] = faq[i];
    const cx = 0.5 + (i % 2) * 4.65, cy = 1.55 + Math.floor(i / 2) * 1.75;
    card(s, cx, cy, 4.35, 1.5);
    s.addShape(pres.shapes.OVAL, { x: cx + 0.2, y: cy + 0.2, w: 0.45, h: 0.45, fill: { color: ACC[i] }, line: { type: "none" }, objectName: "q-badge" });
    s.addText("Q", { x: cx + 0.2, y: cy + 0.2, w: 0.45, h: 0.45, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true, fontFace: "Arial" });
    txt(s, q, { x: cx + 0.8, y: cy + 0.2, w: 3.4, h: 0.45, fontSize: 15, bold: true, valign: "middle" });
    txt(s, a, { x: cx + 0.2, y: cy + 0.75, w: 3.95, h: 0.7, fontSize: 13, color: C.accent6 });
  }

  // 32. Closing
  s = pres.addSlide({ masterName: "TITLE", sectionTitle: curSection });
  s.addText("THANK YOU", { placeholder: "kicker" });
  s.addText("오늘 만든 자료로\n내일 수업을 시작하세요", { placeholder: "title" });
  s.addText("직접 만든다 · 함께 만든다 · 바르게 쓴다", { placeholder: "subtitle" });
  s.addText("질의응답  ·  canva.com/education", { placeholder: "info" });
  await iconCircle(s, 8.1, 0.9, 1.6, C.background1, "FaGraduationCap", HEX.accent1);
  s.addNotes("핵심 세 가지를 다시 짚고 질의응답을 진행합니다. 연수 클래스는 연수 후에도 유지되니 자료를 계속 공유해 주세요.");

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
