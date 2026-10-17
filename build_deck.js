const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/root/.claude/skills/synced/2ede6a27-38b2-4f07-b097-5ba2eeb0cc76_c40bee45-0730-4a80-b5a1-41038ef817ca/pptx/scripts/apply_theme.js");

const fs = require("fs");
const QRPATH = "/home/user/arctic_presentation_2026_10_20/qr_contacts.png";
const HAS_QR = fs.existsSync(QRPATH);
const OUT = "/home/user/arctic_presentation_2026_10_20/Arktika_dengi_i_bezopasnost.pptx";
const LOGO = "/home/user/arctic_presentation_2026_10_20/logo.png";

const THEME = {
  name: "Severny Proekt",
  headFontFace: "Cambria",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1E1640", lt1: "FFFFFF", dk2: "3D2F73", lt2: "F0EDF7",
    accent1: "3E8E34", accent2: "E8771E", accent3: "6B5CA5",
    accent4: "B9AEE0", accent5: "C0392B", accent6: "2E8B5A",
    hlink: "3E8E34", folHlink: "6B5CA5",
  },
};
const HEX = { navy: "1E1640", blue: "3D2F73", ice: "F0EDF7", teal: "3E8E34", orange: "E8771E", steel: "6B5CA5" };

const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "Арктика: деньги и безопасность";
pres.author = "Береговой офис";
const C = pres.SchemeColor;

// ---------- layouts ----------
const titlePh = (color) => ({
  placeholder: {
    options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 7.8, h: 0.8, fontSize: 28, bold: true, color, align: "left", valign: "middle", margin: 0 },
    text: "",
  },
});
pres.defineSlideMaster({
  title: "LIGHT", background: { color: "FFFFFF" },
  objects: [titlePh(C.text1), { image: { x: 8.5, y: 0.22, w: 1.0, h: 0.667, path: LOGO } }],
  slideNumber: { x: 9.1, y: 5.22, w: 0.5, h: 0.3, fontSize: 10, color: C.accent3 },
});
pres.defineSlideMaster({
  title: "DARK", background: { color: HEX.navy },
  objects: [titlePh(C.background1), { rect: { x: 8.42, y: 0.17, w: 1.16, h: 0.78, fill: { color: "FFFFFF" } } }, { image: { x: 8.5, y: 0.22, w: 1.0, h: 0.667, path: LOGO } }],
  slideNumber: { x: 9.1, y: 5.22, w: 0.5, h: 0.3, fontSize: 10, color: C.accent4 },
});

pres.defineSlideMaster({
  title: "TITLE", background: { color: HEX.navy },
  objects: [{ rect: { x: 0.6, y: 0.4, w: 1.6, h: 1.07, fill: { color: "FFFFFF" } } }, { image: { x: 0.65, y: 0.44, w: 1.5, h: 1.0, path: LOGO } }, { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 1.75, w: 6.2, h: 1.4, fontSize: 40, bold: true, color: C.background1, align: "left", valign: "top", margin: 0 }, text: "" } }],
});

// ---------- helpers ----------
const src = (s, text, dark) =>
  s.addText(text, { x: 0.5, y: 5.2, w: 8.5, h: 0.3, fontSize: 10, color: dark ? C.accent4 : C.accent3, margin: 0, valign: "middle", isTextBox: true, objectName: "Источник" });

const card = (s, x, y, w, h, opts = {}) =>
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: opts.fill || C.background2 }, line: { type: "none" }, objectName: opts.name || "Карточка" });

const stat = (s, x, y, w, h, big, label, opts = {}) => {
  card(s, x, y, w, h, { fill: opts.fill, name: "Показатель" });
  s.addText(big, { x: x + 0.15, y: y + 0.08, w: w - 0.3, h: 0.55, fontSize: opts.bigSize || 34, bold: true, fontFace: "Cambria", color: opts.bigColor || C.accent1, valign: "middle", margin: 0, isTextBox: true });
  s.addText(label, { x: x + 0.15, y: y + 0.68, w: w - 0.3, h: h - 0.76, fontSize: 14, color: opts.labelColor || C.text1, valign: "top", margin: 0, isTextBox: true });
};

const circleNum = (s, x, y, n, fill) => {
  s.addShape(pres.shapes.OVAL, { x, y, w: 0.42, h: 0.42, fill: { color: fill || C.accent1 }, line: { type: "none" }, objectName: "Номер" });
  s.addText(String(n), { x, y, w: 0.42, h: 0.42, fontSize: 16, bold: true, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
};

const question = (s, y, text, dark) => {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 9, h: 0.5, rectRadius: 0.08, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Вопрос залу" });
  s.addText([{ text: "Вопрос залу: ", options: { bold: true } }, { text }], { x: 0.7, y, w: 8.6, h: 0.5, fontSize: 14, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
};

const chartBase = (title) => ({
  showTitle: true, title, titleFontSize: 12, titleColor: HEX.navy, titleFontFace: "+mn-lt",
  showLegend: false, showValue: true, dataLabelPosition: "outEnd", dataLabelFontSize: 12, dataLabelFontFace: "+mn-lt", dataLabelColor: HEX.navy, dataLabelFontBold: true, dataLabelFormatCode: "0.0",
  catAxisLabelFontSize: 11, catAxisLabelFontFace: "+mn-lt", catAxisLabelColor: HEX.navy,
  valAxisLabelFontSize: 10, valAxisLabelFontFace: "+mn-lt", valAxisLabelColor: HEX.steel,
  valGridLine: { color: "D5E3EA", size: 0.5 }, catGridLine: { style: "none" },
  valAxisMinVal: 0,
});

// =====================================================================
// 1. Титул
// =====================================================================
pres.addSection({ title: "Вступление" });
let s = pres.addSlide({ masterName: "TITLE", sectionTitle: "Вступление" });
// «Трасса»: дуги и точки
s.addShape(pres.shapes.OVAL, { x: 5.6, y: -1.2, w: 6.8, h: 6.8, fill: { type: "none" }, line: { color: C.accent3, width: 1.5 }, objectName: "Дуга 1" });
s.addShape(pres.shapes.OVAL, { x: 6.5, y: -0.3, w: 5.0, h: 5.0, fill: { type: "none" }, line: { color: C.accent1, width: 1.5 }, objectName: "Дуга 2" });
s.addShape(pres.shapes.OVAL, { x: 7.5, y: 0.7, w: 3.0, h: 3.0, fill: { type: "none" }, line: { color: C.accent4, width: 1.5 }, objectName: "Дуга 3" });
s.addShape(pres.shapes.OVAL, { x: 8.85, y: 2.05, w: 0.3, h: 0.3, fill: { color: C.accent2 }, line: { type: "none" }, objectName: "Точка" });
s.addText("Арктика: деньги и безопасность", { placeholder: "title" });
s.addText("Государство, ледоколы, экономика рейса и safety first: где проходит баланс для нашего флота на трассе Севморпути", { x: 0.6, y: 3.2, w: 5.6, h: 0.8, fontSize: 16, color: C.accent4, margin: 0, valign: "top", isTextBox: true });
s.addText("Пицуков Михаил Георгиевич", { x: 0.6, y: 4.05, w: 6.5, h: 0.32, fontSize: 18, bold: true, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
s.addText("менеджер отдела танкерных перевозок", { x: 0.6, y: 4.37, w: 6.5, h: 0.3, fontSize: 14, color: C.accent4, margin: 0, valign: "middle", isTextBox: true });
if (HAS_QR) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.0, y: 3.55, w: 1.5, h: 1.65, rectRadius: 0.08, fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: "QR фон" });
  s.addImage({ path: QRPATH, x: 8.1, y: 3.62, w: 1.3, h: 1.3, altText: "QR-код с контактами докладчика" });
  s.addText("Контакты", { x: 8.0, y: 4.94, w: 1.5, h: 0.22, fontSize: 11, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
s.addText("ООО «ТК Северный Проект»  ·  Доклад для капитанов и берегового офиса  ·  октябрь 2026", { x: 0.6, y: 4.95, w: 6.5, h: 0.3, fontSize: 12, color: C.background2, margin: 0, isTextBox: true });
s.addNotes("[0:30] Добрый день, коллеги. Меня зовут Михаил Георгиевич Пицуков, я менеджер отдела танкерных перевозок." + (HAS_QR ? " Мои контакты — по QR-коду на этом и на последнем слайде." : "") + " Тема — Арктика через призму двух вещей, которые в нашей профессии вечно спорят: деньги и безопасность. Формат: 15 минут рассказа, по ходу несколько вопросов к залу — отвечайте с места, мне важнее ваш опыт, чем мой слайд. Потом 15 минут обсуждения. Данные — на октябрь 2026, по официальным источникам: Росатом, Минтранс, Минвостокразвития, Правительство, плюс международные для сверки. Где источники расходятся — скажу прямо.");

// =====================================================================
// 2. Интерактив
// =====================================================================
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Вступление" });
s.addText("Начнём с вопроса залу", { placeholder: "title" });
s.addText("Какая доля грузов на Севморпути — транзит (сквозной проход трассы)?", { x: 0.5, y: 1.3, w: 9, h: 0.9, fontSize: 24, bold: true, fontFace: "Cambria", color: C.background1, margin: 0, valign: "middle", isTextBox: true });
const opts2 = [["А", "около 3%"], ["Б", "около 9%"], ["В", "около 25%"], ["Г", "около 50%"]];
opts2.forEach(([l, t], i) => {
  const x = 0.5 + i * 2.28;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 2.6, w: 2.1, h: 1.6, rectRadius: 0.1, fill: { color: C.text2 }, line: { color: C.accent3, width: 1 }, objectName: "Вариант " + l });
  s.addText(l, { x, y: 2.7, w: 2.1, h: 0.7, fontSize: 36, bold: true, fontFace: "Cambria", color: C.accent4, align: "center", valign: "middle", margin: 0, isTextBox: true });
  s.addText(t, { x, y: 3.45, w: 2.1, h: 0.6, fontSize: 18, color: C.background1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
s.addText("Голосуем поднятием руки. Ответ — через пару слайдов.", { x: 0.5, y: 4.55, w: 9, h: 0.4, fontSize: 14, color: C.accent4, margin: 0, isTextBox: true });
s.addNotes("[1:00] Интерактив №1. Дайте залу 20-30 секунд, посчитайте руки. Не раскрывайте ответ — вернёмся на слайде «Что на самом деле везут». Подсказка на случай тишины: в публичных разговорах про Севморпуть чаще всего слышен именно транзит.");

// =====================================================================
// 3. Объёмы
// =====================================================================
pres.addSection({ title: "Картина" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Картина" });
s.addText("Объёмы: плато и ставка на рост", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "млн т", labels: ["2024 факт", "2025 факт", "2026 Минтранс", "2026 Росатом"], values: [37.9, 37.0, 38.2, 40] }],
  Object.assign(chartBase("Грузоперевозки по СМП, млн т"), { x: 0.5, y: 1.3, w: 5.2, h: 3.7, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.teal], valAxisMaxVal: 50, valAxisMajorUnit: 10 }));
stat(s, 6.0, 1.3, 3.5, 1.15, "+16–17%", "к 2025 г.: около 25 млн т на конец августа");
stat(s, 6.0, 2.6, 3.5, 1.15, "80 млн т", "цель на 2024 г. (2018). Факт: 37,9", { bigColor: C.accent2 });
stat(s, 6.0, 3.9, 3.5, 1.15, "110–150", "млн т к 2035 г., официальные сценарии");
src(s, "Источники: Росатом, Минтранс (прогнозы, интервью А. Лихачева, 2025–26), Carnegie для сверки факта 2025");
s.addNotes("[1:30] 2024 — рекорд, 37,9 млн т. 2025 — чуть ниже, около 37 (Росатом называет 37,02; в отдельных обзорах встречается 36,2 — разная методика: грузооборот против объёма перевозок). 2026: Росатом обещает больше 40 млн т, Минтранс осторожнее — 38,2. С начала года темп действительно выше прошлогоднего: на конец августа около 25 млн т, +16–17%. Но напомню: в 2018 году ставили цель 80 млн т к 2024-му. До сих пор меньше половины. Поэтому ключевой вопрос не «растёт ли», а «что именно растёт и кто за это платит».");

// =====================================================================
// 4. Структура
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Картина" });
s.addText("Что на самом деле везут", { placeholder: "title" });
stat(s, 0.5, 1.2, 2.9, 1.45, "3,2 млн т", "сквозной транзит, 103 прохода в обе стороны: менее 9%", { bigColor: C.accent2 });
stat(s, 3.55, 1.2, 2.9, 1.45, "5,4 млн т", "порты России и Китая: около 15% потока");
stat(s, 6.6, 1.2, 2.9, 1.45, "−18,3%", "сухие грузы в портах Арктики: 21,3 млн т, 2025 к 2024", { bigColor: C.accent5 });
s.addChart(pres.charts.BAR, [{ name: "рейсов", labels: ["2023", "2024", "2025", "2026 план"], values: [7, 14, 24, 30] }],
  Object.assign(chartBase("Контейнерные рейсы китайских перевозчиков (Росатом)"), { x: 0.5, y: 2.8, w: 4.9, h: 2.35, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.steel], dataLabelFormatCode: "0", valAxisMaxVal: 40, valAxisMajorUnit: 10 }));
card(s, 5.6, 2.8, 3.9, 2.3, { name: "Вывод" });
s.addText([
  { text: "Что это значит для нас", options: { bold: true, breakLine: true } },
  { text: "Основа потока — внутренние отгрузки (СПГ, конденсат, уголь), а не транзит.", options: { bullet: true, breakLine: true } },
  { text: "Наши танкеры работают в основном сегменте: зимой конденсат Ямала, летом северный завоз.", options: { bullet: true, breakLine: true } },
  { text: "Сухие грузы пока слабое звено.", options: { bullet: true } },
], { x: 5.75, y: 2.9, w: 3.6, h: 2.1, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 4, isTextBox: true });
src(s, "Источники: Росатом; АСОП (порты); CHNL (состав 103 проходов, не официальные данные); Carnegie");
s.addNotes("[1:30] Возвращаемся к вопросу. Правильный ответ — Б: сквозной транзит по Севморпути — 3,2 млн т из 37, то есть меньше 9%. Основа потока — вывоз углеводородов (по данным Востокгосплана, 86% потока — СПГ, нефть, нефтепродукты) и северный завоз. Россия—Китай — ещё 5,4 млн т, около 15%. Откуда 103 прохода: это не контейнерные рейсы. По оценке CHNL (не официальные данные) в 2025 г. это 34 танкера, 23 навалочных судна, 15 контейнеровозов, 12 судов генгруза, 7 рыболовных, 5 газовозов, 3 рефрижераторных и 4 прочих. Считаются проходы в обе стороны, многие — в балласте. Отдельно, контейнерные рейсы китайских перевозчиков (7, 14, 24 и план 30+) — другая статистика Росатома: туда входят и рейсы между портами России и Китая, не только транзит Азия–Европа. Методики разные, поэтому цифры не складываются. Про −18,3%: это абсолютный объём перевалки сухих грузов в портах Арктического бассейна (АСОП): 21,3 млн т в 2025 г. против 26 млн т годом ранее; за 10 месяцев было −16,4%. Это не доля, а именно падение тонн; и это порты Арктики (Мурманск, Архангельск, Сабетта и др.), а не только трасса СМП. Вопрос залу: кто видит этот спад в своих рейсах?");

// =====================================================================
// 5. Государство
// =====================================================================
pres.addSection({ title: "Деньги" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Деньги" });
s.addText("Государство: кто за что отвечает", { placeholder: "title" });
const roles = [
  ["Правительство и Минтранс", "Правила плавания (ПП № 1487 от 18.09.2020, действуют до 01.03.2028). Безопасность, спасательный флот."],
  ["Росатом", "Администрация СМП выдаёт разрешения; «Росатомфлот» водит ледоколы; «Росатом Арктика» единый оператор северного завоза с 2026 г."],
  ["Минвостокразвития", "Развитие Арктики и грузопоток; позиция по тарифам на ледокольную проводку."],
  ["МЧС и Северный флот", "Поиск и спасение, учения. Последнее: 14.10.2026 в Карском море."],
];
roles.forEach(([h, t], i) => {
  const x = 0.5 + i * 2.28;
  card(s, x, 1.3, 2.1, 2.9);
  circleNum(s, x + 0.15, 1.45, i + 1);
  s.addText(h, { x: x + 0.15, y: 1.95, w: 1.8, h: 0.6, fontSize: 14, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addText(t, { x: x + 0.15, y: 2.6, w: 1.8, h: 1.55, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true });
});
card(s, 0.5, 4.35, 9, 0.7, { fill: C.text2 });
s.addText("Для нас: правила пишет Минтранс, допуск и проводку даёт Росатом, а цену проводки ещё решают", { x: 0.7, y: 4.35, w: 8.6, h: 0.7, fontSize: 14, bold: true, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
src(s, "Источники: Правительство РФ, Минтранс, Росатом, Минвостокразвития; Interfax, Коммерсантъ");
s.addNotes("[1:30] Государство здесь не фон, а участник каждого рейса. Правила плавания — постановление Правительства № 1487 от 18 сентября 2020 года, действует до 1 марта 2028-го; значит в ближайшие полтора года ждите новой редакции. Разрешение на плавание выдаёт Администрация СМП, то есть структура Росатома. Росатомфлот водит ледоколы. С 1 января 2026 года «Росатом Арктика» — единый морской оператор северного завоза. Минтранс и МЧС отвечают за аварийно-спасательную готовность. Мысль для зала: чем больше функций у одного игрока, тем важнее для нас предсказуемость тарифов и сроков. Вопрос: кто в вашей практике чаще всего «узкое горлышко» — разрешение, проводка или погода?");

// =====================================================================
// 6. Ледоколы
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Деньги" });
s.addText("Ледоколы: кто платит", { placeholder: "title" });
s.addChart(pres.charts.BAR, [{ name: "млн ₽/сут", labels: ["Готовы платить грузоотправители", "Расчёт со сбором", "Расчёт без сбора"], values: [15.7, 16, 52] }],
  Object.assign(chartBase("Суточный тариф проводки новыми ледоколами, млн ₽ в сутки"), { x: 0.5, y: 1.25, w: 4.7, h: 2.3, barDir: "col", barGapWidthPct: 60, chartColors: [HEX.teal], valAxisMaxVal: 60, valAxisMajorUnit: 20 }));
s.addText("Слева: максимум, который готовы платить грузоотправители. В центре: расчёт тарифа при введении сбора. Справа: расчёт без внешнего финансирования.", { x: 0.5, y: 3.6, w: 4.7, h: 0.7, fontSize: 11, color: C.accent3, margin: 0, valign: "top", isTextBox: true });
card(s, 5.4, 1.3, 4.1, 2.9);
s.addText([
  { text: "17 ледоколов к 2030 г., 18 к 2035 г.", options: { bullet: true, breakLine: true } },
  { text: "«Ленинград» (сдача 12.2028) и «Сталинград» (12.2030), проект 22220: стоимость растёт", options: { bullet: true, breakLine: true } },
  { text: "Проект сбора на их достройку: 20,07 ₽ за тонну экспортного груза до 2036 г., 27 компаний", options: { bullet: true, breakLine: true } },
  { text: "«Лидер»: по словам В. Путина, к 2030 г.", options: { bullet: true } },
], { x: 5.55, y: 1.4, w: 3.8, h: 2.7, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 5, isTextBox: true });
question(s, 4.4, "кто должен платить за ледоколы: бюджет, грузоотправитель или судовладелец?");
src(s, "Источники: Коммерсантъ, The Insider (проект Минтранса, не документ в силе), dp.ru, Балтийский завод");
s.addNotes("[1:45] Это самый денежный слайд. Группировка растёт: зимой 2024 года работало восемь атомных ледоколов — исторический максимум; цель — 17 к 2030-му, 18 к 2035-му. Два новых ледокола проекта 22220 строит Балтийский завод: «Ленинград» (сдача декабрь 2028) и «Сталинград» (декабрь 2030). Они дорожают: по данным dp.ru «Ленинград» оценивался в 56,6 млрд, сейчас около 61,3; Балтзавод допускал рост цены до 30%. Как читать график. Все три столбца — суточный тариф проводки новыми ледоколами в млн рублей. Левый, 15,7, — максимум, который готовы платить грузоотправители. Средний, 16, — расчёт тарифа, если ввести сбор. Правый, 52, — расчёт, если ледоколы окупать только через тариф, без внешнего финансирования. Это цифры из публикаций Коммерсанта и The Insider по документам Минтранса; официально не принято, перед докладом перепроверьте. Сбор — 20,07 рубля за тонну экспортного груза до конца 2036 года, платить будут 27 компаний; предлагают также фиксированный тариф с 2028 года. Интерактив №2: кто должен платить? Капитаны и береговой офис часто отвечают по-разному.");

// =====================================================================
// 7. Экономика рейса
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Деньги" });
s.addText("Наши танкеры: лето и зима", { placeholder: "title" });
const seasons = [
  ["Лето: северный завоз", C.accent1, [
    "«Росатом Арктика» — единый оператор северного завоза с 2026 г.",
    "Подтверждено 264 тыс. т на 2026 г., +65% к 2025 г.",
    "Окно короткое: график зависит от ледовой обстановки на трассе",
  ]],
  ["Зима: вывоз конденсата", C.accent3, [
    "Ямал СПГ и Арктик СПГ-2; перевозка конденсата только под российским флагом",
    "Ледокольная проводка и ледовый класс — главные статьи затрат и риска",
    "Арктик СПГ-2 под санкциями; отгрузки растут: СПГ 3,6 млн т к 12.10.2026",
  ]],
];
seasons.forEach(([h, col, items], i) => {
  const x = 0.5 + i * 4.6;
  card(s, x, 1.3, 4.4, 3.1);
  s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.45, w: 0.42, h: 0.42, fill: { color: col }, line: { type: "none" }, objectName: "Сезон " + (i + 1) });
  s.addText(h, { x: x + 0.7, y: 1.45, w: 3.6, h: 0.42, fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(items.map((tx, k) => ({ text: tx, options: { bullet: true, breakLine: k < items.length - 1 } })), { x: x + 0.15, y: 2.05, w: 4.1, h: 2.25, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 6, isTextBox: true });
});
card(s, 0.5, 4.55, 9, 0.5, { fill: C.text2 });
s.addText("Маржа сезона = фрахт − (проводка + ледовые надбавки + простои + топливо). Считаем по рейсам", { x: 0.7, y: 4.55, w: 8.6, h: 0.5, fontSize: 14, bold: true, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
src(s, "Источники: Росатом Арктика, КТМ РФ, Reuters/LSEG; формула — рабочая модель для обсуждения");
s.addNotes("[1:45] Теперь про нас. Два танкера, два сезона. Лето — северный завоз: с 1 января 2026 года единый морской оператор — «Росатом Арктика», на 2026 год подтверждено более 264 тысяч тонн, на 65% больше, чем в 2025-м. Окно короткое, график зависит от льда. Зима — вывоз конденсата с Ямал СПГ и Арктик СПГ-2: здесь работает исключительное право российского флага, но главные расходы и риски — ледокольная проводка и ледовый класс. Арктик СПГ-2 под санкциями, отгрузки при этом растут: 3,6 млн т СПГ к 12 октября против 1,4 за весь 2025 год. Формула внизу — рабочая модель, не источник: считаем маржу сезона по каждому рейсу. Вопрос залу: какая часть зимней маржи уходит на проводку и ожидание ледокола?");

// =====================================================================
// 8. Safety: рамка допуска
// =====================================================================
pres.addSection({ title: "Безопасность" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Безопасность" });
s.addText("Safety first: рамка допуска", { placeholder: "title" });
const steps = [
  ["Судно", "Polar Ship Certificate и PWOM (руководство по эксплуатации судна в полярных водах); категории A, B, C"],
  ["Разрешение", "Заявка в Администрацию СМП: судно, маршрут, даты плавания"],
  ["Проводка", "Ледокольная и ледовая лоцманская проводка по правилам Минтранса"],
  ["Экипаж", "Подготовка по Полярному кодексу, STCW: базовый и повышенный уровни"],
];
steps.forEach(([h, t], i) => {
  const x = 0.5 + i * 2.28;
  card(s, x, 1.3, 2.1, 2.7);
  circleNum(s, x + 0.15, 1.45, i + 1, C.accent3);
  s.addText(h, { x: x + 0.15, y: 1.95, w: 1.8, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addText(t, { x: x + 0.15, y: 2.4, w: 1.8, h: 1.55, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  if (i < 3) s.addText("›", { x: x + 2.08, y: 2.45, w: 0.22, h: 0.5, fontSize: 28, bold: true, color: C.accent1, align: "center", valign: "middle", margin: 0, isTextBox: true });
});
card(s, 0.5, 4.2, 9, 0.85, { fill: C.text2 });
s.addText("Восточный сектор зимой без ледокола недоступен: лёд до 3 м. Рамка допуска не про бюрократию, а про выживаемость рейса.", { x: 0.7, y: 4.2, w: 8.6, h: 0.85, fontSize: 14, bold: true, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
src(s, "Источники: Полярный кодекс ИМО/СОЛАС, STCW; Правила плавания в акватории СМП (ПП РФ № 1487); Ведомости");
s.addNotes("[1:30] Четыре замка, которые нельзя открыть в обход. Первый — судно: Polar Ship Certificate и PWOM. PWOM — Polar Water Operational Manual, руководство по эксплуатации судна в полярных водах: документ, который судовладелец готовит под конкретное судно и район плавания на основе оценки рисков, он входит в пакет для получения сертификата. Категории A, B, C — по ледовой прочности. Второй — разрешение Администрации СМП. Третий — проводка, ледокольная и ледовая лоцманская. Четвёртый — экипаж: подготовка по Полярному кодексу, в STCW есть базовый уровень для вахтенных помощников и повышенный для капитанов и старших помощников. Экономия на любом из замков — это не экономия, а рассрочка по гораздо более дорогому платежу. Вопрос капитанам: какой из четырёх замков в реальной практике «скрипит» больше всего?");

// =====================================================================
// 9. Лёд 2026 и спасательный флот
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Безопасность" });
s.addText("Лёд 2026: риск не уменьшился", { placeholder: "title" });
stat(s, 0.5, 1.2, 2.9, 1.6, "14,29", "млн км²: зимний максимум 15.03.2026, почти рекордно низкий", { bigColor: C.steel });
stat(s, 3.55, 1.2, 2.9, 1.6, "5,34", "млн км²: среднее за сентябрь 2026; 16-й минимум из 48 лет (1979–2026)", { bigColor: C.steel });
stat(s, 6.6, 1.2, 2.9, 1.6, "5 + 5", "5 спасательных судов принято в 2025–26 и 5 буксиров до конца 2027", { bigColor: C.accent1 });
card(s, 0.5, 3.0, 9, 1.35);
s.addText([
  { text: "Лето 2026 было холодным и штормовым, а не «чистым ото льда».", options: { bullet: true, breakLine: true } },
  { text: "С 2007 г. сентябрьский лёд почти не убывает: окно открывается, но непредсказуемость остаётся.", options: { bullet: true, breakLine: true } },
  { text: "Спасательный флот растёт, но для огромной акватории его пока мало: помощь может быть далеко, нужна автономность.", options: { bullet: true } },
], { x: 0.7, y: 3.1, w: 8.6, h: 1.2, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true });
question(s, 4.5, "какой сигнал заставит вас отложить или отменить рейс?");
src(s, "Источники: NSIDC, Met Office (end-of-season report, окт. 2026), Минтранс (А. Пошивай, ВЭФ, сентябрь 2026), Северный флот");
s.addNotes("[1:30] Самое частое заблуждение: «льда стало меньше — стало проще». По NSIDC зимний максимум 2026 года, 14,29 млн кв. км, почти рекордно низкий. Сентябрь 2026 по Met Office: среднее 5,34 млн кв. км — 16-й минимум за 48 лет спутниковых наблюдений (1979–2026), то есть льда много меньше, чем в 1980–2000-х, но пятый по величине с 2007 года; лето было холодное и штормовое. Лёд меньше, но менее предсказуем: торосы, дрейф. «5 + 5»: замминистра транспорта Пошивай на ВЭФ сообщил, что в 2025–2026 годах принято пять ледовых спасательных судов, а до конца 2027 года планируется принять ещё пять буксиров. Тезис про спасательный флот: он растёт, но речь о десятке единиц на акваторию в тысячи километров, и отраслевые материалы прямо говорят, что нужно больше аварийно-спасательных формирований. Поэтому помощь рядом не гарантирована, и автономность — наша ответственность. Рост флота — аргумент в пользу будущего, не в пользу сегодняшнего рейса. 14 октября Северный флот учился спасать танкер в Карском море. Вопрос залу: что для вас — стоп-сигнал?");

// =====================================================================
// 10. Флаг и санкции
// =====================================================================
pres.addSection({ title: "Риски" });
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Риски" });
s.addText("Флаг и санкции: коротко", { placeholder: "title" });
const sanc = [
  ["Флаг", "Исключительное право судов под российским флагом на перевозку нефти, газа, СПГ, конденсата и угля по СМП: наш конденсат подпадает. Требование строить такие суда в РФ (с 01.03.2026) пытались отсрочить: статус уточнить."],
  ["Арктик СПГ-2", "3,6 млн т и 51 партия с января по 12.10.2026 (в 2025 г. было 1,4 млн т). Схема через перевалку в Мурманске и на Камчатке: санкционный контекст."],
  ["Для нас", "Арктик СПГ-2 под санкциями: проверка контрагента, страховое покрытие, платёжный канал и порты захода. Решаем до подписания чартера."],
];
sanc.forEach(([h, t], i) => {
  const x = 0.5 + i * 3.05;
  card(s, x, 1.3, 2.9, 3.6, { fill: i === 2 ? C.text2 : C.background2 });
  circleNum(s, x + 0.15, 1.45, i + 1, i === 2 ? C.accent2 : C.accent1);
  s.addText(h, { x: x + 0.7, y: 1.45, w: 2.1, h: 0.42, fontSize: 18, bold: true, color: i === 2 ? C.background1 : C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(t, { x: x + 0.15, y: 2.05, w: 2.6, h: 2.75, fontSize: 14, color: i === 2 ? C.background1 : C.text1, margin: 0, valign: "top", isTextBox: true });
});
src(s, "Источники: Кодекс торгового мореплавания (ст. 4), Interfax, The Moscow Times (Reuters/LSEG); третий блок — практический вывод докладчика");
s.addNotes("[1:00] Коротко, как договаривались. Три вещи. Первое — флаг: по закону перевозка нефти, газа, СПГ, конденсата и угля по СМП — только суда под российским флагом; требование строить такие суда в России с 1 марта 2026 года планировали, в 2025 году речь шла об отсрочке на два года, статус надо проверить в актуальной редакции. Второе — Арктик СПГ-2: отгрузки растут (3,6 млн т к 12 октября против 1,4 млн т за весь 2025 год), но идут через перевалку и под санкциями. Третье — наш практический вывод, это не цитата из источника: любой рейс в этом сегменте начинается с проверки контрагента, страхования и платежей. Не вдаюсь глубоко — вопросы в Q&A.");

// =====================================================================
// 11. Светофор
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Риски" });
s.addText("Светофор решения по рейсу", { placeholder: "title" });
const lights = [
  ["ИДТИ", C.accent6, "Разрешение получено, проводка подтверждена, класс и экипаж подходят, страховка и платежи закрыты, окно навигации ясное."],
  ["ЖДАТЬ", C.accent2, "Нет подтверждённой проводки или даты, нестабильный ледовый прогноз, не закрыт вопрос контрагента или оплаты."],
  ["ОТКАЗАТЬСЯ", C.accent5, "Судно или экипаж не соответствует категории, нет автономности на задержку, страховщик не покрывает рейс."],
];
lights.forEach(([h, col, t], i) => {
  const x = 0.5 + i * 3.05;
  card(s, x, 1.3, 2.9, 3.0);
  s.addShape(pres.shapes.OVAL, { x: x + 0.15, y: 1.45, w: 0.5, h: 0.5, fill: { color: col }, line: { type: "none" }, objectName: "Сигнал " + h });
  s.addText(h, { x: x + 0.8, y: 1.45, w: 2.0, h: 0.5, fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addText(t, { x: x + 0.15, y: 2.1, w: 2.6, h: 2.1, fontSize: 16, color: C.text1, margin: 0, valign: "top", isTextBox: true });
});
card(s, 0.5, 4.45, 9, 0.6, { fill: C.text2 });
s.addText("Деньги считаем после того, как safety дало «зелёный». Это и есть баланс.", { x: 0.7, y: 4.45, w: 8.6, h: 0.6, fontSize: 14, bold: true, color: C.background1, valign: "middle", margin: 0, isTextBox: true });
src(s, "Предложение докладчика для обсуждения; в основе: Полярный кодекс, Правила плавания в акватории СМП");
s.addNotes("[1:30] Попытка свести всё в одну рамку. Это предложение для обсуждения, не норматив. Зелёный — все замки открыты. Жёлтый — чего-то не хватает, и деньги тратятся на ожидание. Красный — безопасность не закрыта, экономика не рассматривается. Главная мысль: порядок. Сначала safety даёт зелёный, потом считаем деньги. Если сначала считать деньги, под любую цифру найдётся причина «рискнуть». Интерактив №3: предложите зал поправить критерии. Что вы бы переместили из жёлтого в красный?");

// =====================================================================
// 12. Итоги и Q&A
// =====================================================================
pres.addSection({ title: "Итоги" });
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Итоги" });
s.addText("Итоги и вопросы", { placeholder: "title" });
const take = [
  ["Поток", "Плато около 37 млн т; Росатом прогнозирует 40+ млн т в 2026 г. Основа — внутренние отгрузки, транзит менее 9%."],
  ["Деньги", "Проводка дорожает: проект ледосбора (20,07 ₽ за тонну экспортного груза) и новый тариф решат экономику рейса."],
  ["Safety", "Допуск, проводка, экипаж, автономность. Лёд меньше, риск не меньше."],
];
take.forEach(([h, t], i) => {
  const x = 0.5 + i * 3.05;
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.3, w: 2.9, h: 2.1, rectRadius: 0.08, fill: { color: C.text2 }, line: { type: "none" }, objectName: "Итог " + h });
  s.addText(h, { x: x + 0.15, y: 1.4, w: 2.6, h: 0.45, fontSize: 20, bold: true, fontFace: "Cambria", color: C.accent4, margin: 0, valign: "middle", isTextBox: true });
  s.addText(t, { x: x + 0.15, y: 1.9, w: 2.6, h: 1.4, fontSize: 14, color: C.background1, margin: 0, valign: "top", isTextBox: true });
});
if (HAS_QR) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.0, y: 3.55, w: 1.5, h: 1.65, rectRadius: 0.08, fill: { color: "FFFFFF" }, line: { type: "none" }, objectName: "QR фон" });
  s.addImage({ path: QRPATH, x: 8.1, y: 3.62, w: 1.3, h: 1.3, altText: "QR-код с контактами докладчика" });
  s.addText("Контакты", { x: 8.0, y: 4.94, w: 1.5, h: 0.22, fontSize: 11, bold: true, color: C.text1, align: "center", valign: "middle", margin: 0, isTextBox: true });
}
s.addText("Вопросы для обсуждения", { x: 0.5, y: 3.6, w: 9, h: 0.4, fontSize: 18, bold: true, fontFace: "Cambria", color: C.background1, margin: 0, isTextBox: true });
s.addText([
  { text: "Где граница выгодности зимнего рейса за конденсатом при росте цены проводки?", options: { bullet: true, breakLine: true } },
  { text: "Северный завоз при едином операторе: рост загрузки или рост зависимости?", options: { bullet: true, breakLine: true } },
  { text: "Что нам нужно от государства: тариф, график, спасательный флот?", options: { bullet: true } },
], { x: 0.5, y: 4.0, w: HAS_QR ? 7.3 : 9, h: 1.1, fontSize: 14, color: C.accent4, margin: 0, valign: "top", paraSpaceAfter: 3, isTextBox: true });
s.addNotes("[0:30 + 15 мин Q&A] Три мысли. Поток держится на внутренних отгрузках, не на транзите; 40+ млн т — это прогноз Росатома на 2026 год, Минтранс осторожнее (38,2). Деньги: ледокольная проводка дорожает; речь о проекте ледосбора на достройку ледоколов — 20,07 рубля за тонну экспортного груза до 2036 года — и о новом тарифе. Safety — порядок: сначала закрыть замки, потом считать. Дальше ваши вопросы. Если пауза — используйте три заготовленных. Приложение с источниками — на следующем слайде.");

// =====================================================================
// 13. Приложение: источники
// =====================================================================
s = pres.addSlide({ masterName: "LIGHT", sectionTitle: "Итоги" });
s.addText("Источники и что проверить", { placeholder: "title" });
s.addText([
  { text: "Государственные", options: { bold: true, color: C.accent3, breakLine: true } },
  { text: "Росатом / «Росатомфлот»: итоги навигации 2025, прогноз 2026 (А. Лихачев)", options: { bullet: true, breakLine: true } },
  { text: "Минтранс: прогноз 2026, аварийно-спасательный флот, проект сбора", options: { bullet: true, breakLine: true } },
  { text: "Минвостокразвития: контейнерные рейсы, тарифная позиция", options: { bullet: true, breakLine: true } },
  { text: "Правительство РФ: ПП № 1487 от 18.09.2020", options: { bullet: true, breakLine: true } },
  { text: "Международные для сверки", options: { bold: true, color: C.accent3, breakLine: true } },
  { text: "Carnegie, CHNL, АСОП, NSIDC, Met Office, Reuters/LSEG", options: { bullet: true } },
], { x: 0.5, y: 1.25, w: 4.4, h: 3.7, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true });
card(s, 5.1, 1.25, 4.4, 3.7, { fill: C.background2 });
s.addText([
  { text: "Перед выступлением перепроверить", options: { bold: true, color: C.accent5, breakLine: true } },
  { text: "Итоги навигации 2026 (идёт сезон)", options: { bullet: true, breakLine: true } },
  { text: "Статус сбора и тарифов (проект, не документ в силе)", options: { bullet: true, breakLine: true } },
  { text: "Отсрочка требования о российской постройке судов", options: { bullet: true, breakLine: true } },
  { text: "Цифры по сухим грузам и ледовые данные ААНИИ", options: { bullet: true, breakLine: true } },
  { text: "Часть цифр взята из пересказов СМИ, а не из первоисточников", options: { bullet: true } },
], { x: 5.3, y: 1.35, w: 4.0, h: 3.5, fontSize: 14, color: C.text1, valign: "top", margin: 0, paraSpaceAfter: 3, isTextBox: true });
s.addNotes("Слайд для приложения, на докладе не показывать. Здесь список источников и пунктов, которые нужно сверить с первоисточниками перед выступлением.");

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("done", OUT);
})();
