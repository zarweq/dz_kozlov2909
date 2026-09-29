const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const md = require("react-icons/md");
const fs = require("fs");

const OUT = process.argv[2] || "LAN.pptx";
const TOPO_IMG = process.argv[3]; // optional Packet Tracer screenshot

const C = {
  navy: "0B1F3A", navy2: "132C52", ink: "1B2430", muted: "5B6675", bg: "F4F7FB", white: "FFFFFF",
  buh: "22B573", sales: "FF8A00", it: "8B5CF6", srv: "00B4D8", core: "E63E62", line: "C9D3E0",
};
const FONT_H = "Cambria", FONT_B = "Calibri", FONT_C = "Courier New";

async function icon(Comp, color, size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const png = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + png.toString("base64");
}

const rub = (n) => n.toLocaleString("ru-RU").replace(/ /g, " ") + " ₽";

const equipment = [
  ["Маршрутизатор Cisco ISR 2911/K9", 1, 65000, "Маршрутизация VLAN, DHCP"],
  ["Коммутатор Cisco Catalyst 2960-24TT-L", 4, 28000, "1 ядро + 3 доступа"],
  ["Рабочая станция (Core i5, 16 ГБ, SSD 512 ГБ) + монитор 24\"", 9, 62000, "по 3 ПК в отдел"],
  ["Сервер HPE ProLiant DL20 Gen10 Plus", 1, 210000, "DNS, Web, файлы"],
  ["МФУ лазерное сетевое", 3, 24000, "по 1 в отдел"],
  ["Шкаф серверный 19\" 18U", 1, 32000, "размещение ядра сети"],
  ["ИБП APC Smart-UPS 1500 ВА", 1, 48000, "защита питания"],
];
const consumables = [
  ["Кабель витая пара UTP Cat.5e, бухта 305 м", 2, 7500],
  ["Коннекторы RJ-45 8P8C (уп. 100 шт.)", 2, 900],
  ["Изолирующие колпачки (уп. 100 шт.)", 2, 350],
  ["Патч-панель 24 порта Cat.5e", 1, 3200],
  ["Розетка RJ-45 настенная", 12, 350],
  ["Кабель-канал 25×16 мм, 2 м", 40, 180],
  ["Патч-корд UTP 1 м", 20, 150],
  ["Консольный кабель Cisco RJ-45 – USB", 1, 1200],
  ["Кримпер + LAN-тестер", 1, 3500],
  ["Стяжки, маркировка, крепёж (комплект)", 1, 1500],
];
const sum = (rows) => rows.reduce((s, r) => s + r[1] * r[2], 0);
const eqTotal = sum(equipment), csTotal = sum(consumables), total = eqTotal + csTotal;

(async () => {
  const I = {
    router: await icon(md.MdRouter, "FFFFFF"),
    sw: await icon(fa.FaNetworkWired, "FFFFFF"),
    pc: await icon(fa.FaDesktop, "FFFFFF"),
    server: await icon(fa.FaServer, "FFFFFF"),
    sitemap: await icon(fa.FaSitemap, "FFFFFF"),
    layer: await icon(fa.FaLayerGroup, "FFFFFF"),
    cogs: await icon(fa.FaCogs, "FFFFFF"),
    cart: await icon(fa.FaShoppingCart, "FFFFFF"),
    check: await icon(fa.FaCheckCircle, C.buh),
    ruble: await icon(fa.FaRubleSign, "FFFFFF"),
    plug: await icon(fa.FaPlug, "FFFFFF"),
    shield: await icon(fa.FaShieldAlt, "FFFFFF"),
  };

  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "ЛВС офиса на 3 отдела — Cisco Packet Tracer";

  const circle = (s, img, color, x, y, d = 0.6) => {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color, width: 0 } });
    const p = d * 0.22;
    s.addImage({ data: img, x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  };
  const title = (s, text, sub) => {
    s.addText(text, { x: 0.5, y: 0.3, w: 9, h: 0.6, fontFace: FONT_H, fontSize: 30, bold: true, color: C.navy, margin: 0, isTextBox: true });
    if (sub) s.addText(sub, { x: 0.5, y: 0.88, w: 9, h: 0.35, fontFace: FONT_B, fontSize: 14, color: C.muted, margin: 0, isTextBox: true });
  };
  const light = () => { const s = pres.addSlide(); s.background = { color: C.white }; return s; };
  const shadow = () => ({ type: "outer", color: "000000", blur: 6, offset: 2, angle: 90, opacity: 0.12 });

  // 1. Title
  {
    const s = pres.addSlide(); s.background = { color: C.navy };
    const cols = [C.buh, C.sales, C.it, C.srv, C.core];
    const pos = [[6.7, 0.8], [7.65, 0.8], [8.6, 0.8], [7.175, 1.75], [8.125, 1.75]];
    [I.router, I.sw, I.pc, I.server, I.sitemap].forEach((img, i) => circle(s, img, cols[i], pos[i][0], pos[i][1], 0.75));
    s.addText("Проектирование ЛВС\nофиса на 3 отдела", { x: 0.6, y: 1.5, w: 6.2, h: 1.6, fontFace: FONT_H, fontSize: 38, bold: true, color: C.white, margin: 0, isTextBox: true });
    s.addText("Топология, настройка в Cisco Packet Tracer и смета на оборудование", { x: 0.6, y: 3.15, w: 6.2, h: 0.7, fontFace: FONT_B, fontSize: 16, color: "CADCFC", margin: 0, isTextBox: true });
    s.addText("Выполнил: Калиберда Сергей", { x: 0.6, y: 4.6, w: 6, h: 0.4, fontFace: FONT_B, fontSize: 14, color: C.srv, bold: true, margin: 0, isTextBox: true });
  }

  // 2. Task
  {
    const s = light(); title(s, "Задание проекта", "Малый офис: бухгалтерия, отдел продаж и IT-отдел");
    const cards = [
      [I.sitemap, C.core, "Топология", "Иерархическая звезда: маршрутизатор → ядро → 3 коммутатора отделов"],
      [I.layer, C.it, "Сегментация", "Каждый отдел в своей VLAN, серверы в отдельной VLAN 99"],
      [I.cogs, C.sales, "Настройка", "Router-on-a-stick, DHCP для всех ПК, trunk-каналы, пароли"],
      [I.cart, C.buh, "Закупка", "Смета оборудования и расходных материалов с ценами"],
    ];
    cards.forEach(([img, col, h, t], i) => {
      const x = 0.5 + i * 2.3;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.55, w: 2.05, h: 3.4, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.12, shadow: shadow() });
      circle(s, img, col, x + 0.25, 1.8, 0.75);
      s.addText(h, { x: x + 0.25, y: 2.75, w: 1.6, h: 0.45, fontFace: FONT_H, fontSize: 18, bold: true, color: C.navy, margin: 0, isTextBox: true });
      s.addText(t, { x: x + 0.25, y: 3.2, w: 1.6, h: 1.55, fontFace: FONT_B, fontSize: 13, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    });
  }

  // 3. Topology
  {
    const s = light(); title(s, "Топология сети");
    const box = (img, col, x, y, label, sub, side = "below", d = 0.62) => {
      circle(s, img, col, x - d / 2, y, d);
      const txt = [{ text: label, options: { bold: true, breakLine: true } }, { text: sub, options: { fontSize: 9, color: C.muted } }];
      const o = { fontFace: FONT_B, fontSize: 11, color: C.ink, margin: 0, isTextBox: true, h: 0.45 };
      if (side === "below") Object.assign(o, { x: x - 0.8, y: y + d + 0.02, w: 1.6, align: "center" });
      if (side === "right") Object.assign(o, { x: x + d / 2 + 0.08, y: y + d / 2 - 0.22, w: 1.7, align: "left" });
      if (side === "left") Object.assign(o, { x: x - d / 2 - 1.78, y: y + d / 2 - 0.22, w: 1.7, align: "right" });
      s.addText(txt, o);
    };
    const ln = (x1, y1, x2, y2, color = C.line, w = 2) => s.addShape(pres.shapes.LINE, { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1) || 0.001, h: Math.abs(y2 - y1) || 0.001, flipH: x2 < x1, line: { color, width: w } });
    // coords (centers)
    const R = [5, 1.05], CORE = [5, 2.05];
    const acc = [[1.9, 3.1, C.buh, "SW-BUH", "VLAN 10"], [5, 3.1, C.sales, "SW-SALES", "VLAN 20"], [8.1, 3.1, C.it, "SW-IT", "VLAN 30"]];
    const SRV = [7.6, 2.05];
    ln(R[0], R[1] + 0.3, CORE[0], CORE[1] + 0.3, C.navy, 3);
    ln(CORE[0], CORE[1] + 0.3, SRV[0], SRV[1] + 0.3, C.srv, 2);
    acc.forEach(([x, y, c]) => ln(CORE[0], CORE[1] + 0.3, x, y + 0.3, C.navy, 2.5));
    acc.forEach(([x, y, c], i) => [-0.75, 0, 0.75].forEach((dx) => ln(x, y + 0.3, x + dx, 4.3 + 0.22, c, 1.5)));
    box(I.router, C.navy, R[0], R[1], "R1 — Cisco 2911", "G0/0.10/.20/.30/.99", "right");
    box(I.sw, C.core, CORE[0], CORE[1], "SW-CORE", "Cisco 2960", "left");
    box(I.server, C.srv, SRV[0], SRV[1], "SERVER", "192.168.99.10");
    acc.forEach(([x, y, c, n, v], i) => {
      box(I.sw, c, x, y, n, v, i === 2 ? "right" : "left");
      [-0.75, 0, 0.75].forEach((dx, k) => {
        circle(s, I.pc, c, x + dx - 0.22, 4.3, 0.44);
        s.addText(`PC${i * 3 + k + 1}`, { x: x + dx - 0.35, y: 4.76, w: 0.7, h: 0.22, fontFace: FONT_B, fontSize: 9, color: C.ink, align: "center", margin: 0, isTextBox: true });
      });
    });
    // legend
    [["Бухгалтерия", C.buh], ["Продажи", C.sales], ["IT-отдел", C.it], ["Серверы", C.srv]].forEach(([t, c], i) => {
      s.addShape(pres.shapes.OVAL, { x: 0.5, y: 1.2 + i * 0.32, w: 0.18, h: 0.18, fill: { color: c }, line: { color: c } });
      s.addText(t, { x: 0.75, y: 1.13 + i * 0.32, w: 1.6, h: 0.3, fontFace: FONT_B, fontSize: 11, color: C.ink, margin: 0, isTextBox: true });
    });
  }

  // 3b. PT screenshot
  if (TOPO_IMG && fs.existsSync(TOPO_IMG)) {
    const s = light(); title(s, "Топология в Cisco Packet Tracer");
    const meta = await sharp(TOPO_IMG).metadata();
    const maxW = 9, maxH = 4.1; let w = maxW, h = (meta.height / meta.width) * w; if (h > maxH) { h = maxH; w = (meta.width / meta.height) * h; }
    s.addImage({ path: TOPO_IMG, x: (10 - w) / 2, y: 1.15, w, h, shadow: shadow() });
  }

  // 4. Addressing
  {
    const s = light(); title(s, "VLAN и план IP-адресации", "Маска /24 (255.255.255.0) во всех сетях, адреса ПК выдаёт DHCP на R1");
    const hdr = ["VLAN", "Отдел", "Сеть", "Шлюз", "DHCP-диапазон", "Устройства"].map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy } } }));
    const row = (c, cells) => cells.map((t, i) => ({ text: t, options: i === 0 ? { bold: true, color: C.white, fill: { color: c } } : { fill: { color: C.bg } } }));
    s.addTable([hdr,
      row(C.buh, ["10", "Бухгалтерия", "192.168.10.0", "192.168.10.1", ".11 – .254", "PC1–PC3"]),
      row(C.sales, ["20", "Отдел продаж", "192.168.20.0", "192.168.20.1", ".11 – .254", "PC4–PC6"]),
      row(C.it, ["30", "IT-отдел", "192.168.30.0", "192.168.30.1", ".11 – .254", "PC7–PC9"]),
      row(C.srv, ["99", "Серверы / упр.", "192.168.99.0", "192.168.99.1", "статика", "SERVER .10, коммутаторы .2–.5"]),
    ], { x: 0.5, y: 1.5, w: 9, colW: [0.8, 1.55, 1.45, 1.45, 1.45, 2.3], fontFace: FONT_B, fontSize: 12, color: C.ink, border: { type: "solid", color: C.white, pt: 2 }, rowH: 0.5, valign: "middle" });
    s.addText([{ text: "Зачем VLAN: ", options: { bold: true, color: C.navy } }, { text: "отделы изолированы на канальном уровне, широковещательный трафик не выходит за пределы отдела, а доступ между отделами контролирует маршрутизатор." }],
      { x: 0.5, y: 4.25, w: 9, h: 0.7, fontFace: FONT_B, fontSize: 13, color: C.ink, margin: 0, isTextBox: true });
  }

  // 5. Cabling
  {
    const s = light(); title(s, "Схема подключения", "Все соединения — медный прямой кабель (Copper Straight-Through)");
    const hdr = ["Устройство", "Порт", "Подключено к", "Порт"].map((t) => ({ text: t, options: { bold: true, color: C.white, fill: { color: C.navy } } }));
    const r = (a, b, c, d, col) => [{ text: a, options: { bold: true, color: col } }, b, c, d];
    s.addTable([hdr,
      r("R1", "G0/0", "SW-CORE", "G0/1", C.navy),
      r("SW-CORE", "Fa0/1", "SW-BUH", "Fa0/24", C.core),
      r("SW-CORE", "Fa0/3", "SW-SALES", "Fa0/24", C.core),
      r("SW-CORE", "Fa0/5", "SW-IT", "Fa0/24", C.core),
      r("SW-CORE", "Fa0/13", "SERVER", "Fa0", C.core),
      r("SW-BUH", "Fa0/1, 3, 5", "PC1–PC3", "Fa0", C.buh),
      r("SW-SALES", "Fa0/1, 3, 5", "PC4–PC6", "Fa0", C.sales),
      r("SW-IT", "Fa0/1, 3, 5", "PC7–PC9", "Fa0", C.it),
    ], { x: 0.5, y: 1.4, w: 5.6, colW: [1.5, 1.1, 1.9, 1.1], fontFace: FONT_B, fontSize: 12, color: C.ink, border: { type: "solid", color: C.line, pt: 0.75 }, rowH: 0.4, valign: "middle", fill: { color: C.white } });
    const stats = [["16", "устройств", C.core], ["14", "кабелей", C.sales], ["4", "VLAN", C.it]];
    stats.forEach(([n, l, c], i) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.6, y: 1.4 + i * 1.2, w: 2.9, h: 1.0, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.1 });
      s.addText(n, { x: 6.8, y: 1.45 + i * 1.2, w: 1.1, h: 0.9, fontFace: FONT_H, fontSize: 40, bold: true, color: c, margin: 0, valign: "middle", isTextBox: true });
      s.addText(l, { x: 7.9, y: 1.45 + i * 1.2, w: 1.5, h: 0.9, fontFace: FONT_B, fontSize: 16, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
    });
  }

  // 6. Router config
  {
    const s = light(); title(s, "Настройка маршрутизатора R1", "Router-on-a-stick: одна физическая линия, четыре логических подынтерфейса");
    const code = [
      "interface GigabitEthernet0/0", " no shutdown",
      "interface GigabitEthernet0/0.10", " encapsulation dot1Q 10", " ip address 192.168.10.1 255.255.255.0",
      "! аналогично .20, .30, .99", "",
      "enable secret cisco", "line console 0", " password cisco", " login", "",
      "ip dhcp excluded-address 192.168.10.1 192.168.10.10",
      "ip dhcp pool BUH", " network 192.168.10.0 255.255.255.0", " default-router 192.168.10.1", " dns-server 192.168.99.10",
      "! пулы SALES и IT — аналогично",
    ].join("\n");
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.4, w: 5.5, h: 3.7, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.1 });
    s.addText(code, { x: 0.7, y: 1.55, w: 5.2, h: 3.45, fontFace: FONT_C, fontSize: 10.5, color: "D6F5FF", margin: 0, valign: "top", isTextBox: true });
    const pts = [[I.layer, C.it, "Подынтерфейсы", "Шлюз для каждой VLAN"], [I.plug, C.buh, "DHCP", "3 пула — ПК получают адрес автоматически"], [I.shield, C.core, "Безопасность", "enable secret, пароль консоли, баннер"]];
    pts.forEach(([img, c, h, t], i) => {
      circle(s, img, c, 6.35, 1.45 + i * 1.2, 0.6);
      s.addText(h, { x: 7.1, y: 1.4 + i * 1.2, w: 2.4, h: 0.35, fontFace: FONT_H, fontSize: 15, bold: true, color: C.navy, margin: 0, isTextBox: true });
      s.addText(t, { x: 7.1, y: 1.75 + i * 1.2, w: 2.4, h: 0.55, fontFace: FONT_B, fontSize: 12, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    });
  }

  // 7. Switch config
  {
    const s = light(); title(s, "Настройка коммутаторов", "Полные конфигурации — в папке configs/");
    const blk = (x, h, col, code) => {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.4, w: 4.35, h: 3.7, fill: { color: C.bg }, line: { color: col, width: 1.5 }, rectRadius: 0.1 });
      s.addText(h, { x: x + 0.2, y: 1.5, w: 3.9, h: 0.4, fontFace: FONT_H, fontSize: 16, bold: true, color: col, margin: 0, isTextBox: true });
      s.addText(code, { x: x + 0.2, y: 1.95, w: 4.0, h: 3.05, fontFace: FONT_C, fontSize: 10, color: C.ink, margin: 0, valign: "top", isTextBox: true });
    };
    blk(0.5, "SW-CORE (ядро)", C.core, ["vlan 10", " name BUH", "vlan 20", " name SALES", "vlan 30", " name IT", "vlan 99", " name SERVERS", "interface g0/1", " switchport mode trunk", "interface range fa0/1,fa0/3,fa0/5", " switchport mode trunk", "interface fa0/13", " switchport access vlan 99"].join("\n"));
    blk(5.15, "SW-BUH / SW-SALES / SW-IT", C.it, ["vlan 10          ! 20 / 30", " name BUH", "interface fa0/24", " switchport mode trunk", "interface range fa0/1 - 23", " switchport mode access", " switchport access vlan 10", " spanning-tree portfast", "interface vlan 99", " ip address 192.168.99.3 ...", "ip default-gateway 192.168.99.1"].join("\n"));
  }

  // photo helper: fit image into box (contain)
  const IMG_DIR = require("path").join(__dirname, "images");
  const credits = JSON.parse(fs.readFileSync(require("path").join(IMG_DIR, "credits.json"), "utf8"));
  const photo = async (s, key, x, y, w, h) => {
    const f = require("path").join(IMG_DIR, key + ".jpg");
    const m = await sharp(f).metadata();
    let iw = w, ih = (m.height / m.width) * w;
    if (ih > h) { ih = h; iw = (m.width / m.height) * h; }
    s.addImage({ path: f, x: x + (w - iw) / 2, y: y + (h - ih) / 2, w: iw, h: ih });
  };

  // 8. Equipment (cards with photos)
  {
    const s = light(); title(s, "Закупка оборудования", "Ориентировочные цены, 2026 г.");
    const keys = ["router", "switch", "pc", "server", "mfp", "rack", "ups"];
    const cols = [C.navy, C.core, C.sales, C.srv, C.it, C.buh, C.core];
    for (let i = 0; i < equipment.length; i++) {
      const [n, q, p] = equipment[i];
      const x = 0.5 + (i % 4) * 2.3, y = 1.3 + Math.floor(i / 4) * 2.05;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 1.9, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.08, shadow: shadow() });
      await photo(s, keys[i], x + 0.1, y + 0.08, 1.9, 0.85);
      s.addText(n, { x: x + 0.1, y: y + 0.97, w: 1.9, h: 0.5, fontFace: FONT_B, fontSize: 9.5, bold: true, color: cols[i], margin: 0, valign: "top", isTextBox: true });
      s.addText([{ text: `${q} × ${rub(p)}`, options: { color: C.muted } }, { text: "   " + rub(q * p), options: { bold: true, color: C.ink } }],
        { x: x + 0.1, y: y + 1.5, w: 1.9, h: 0.3, fontFace: FONT_B, fontSize: 10, margin: 0, valign: "middle", isTextBox: true });
    }
    const x = 0.5 + 3 * 2.3, y = 1.3 + 2.05;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.1, h: 1.9, fill: { color: C.core }, line: { color: C.core }, rectRadius: 0.08 });
    s.addText("Итого\nоборудование", { x: x + 0.15, y: y + 0.2, w: 1.8, h: 0.7, fontFace: FONT_B, fontSize: 14, color: C.white, margin: 0, isTextBox: true });
    s.addText(rub(eqTotal), { x: x + 0.15, y: y + 1.0, w: 1.8, h: 0.6, fontFace: FONT_H, fontSize: 22, bold: true, color: C.white, margin: 0, isTextBox: true });
  }

  // 8b. One slide per equipment item
  {
    const details = [
      { key: "router", color: C.navy, role: "Маршрутизатор", does: [
          ["Соединяет отделы", "маршрутизирует трафик между VLAN 10, 20, 30 и 99 — без него отделы не видят друг друга"],
          ["Раздаёт адреса", "DHCP-сервер: автоматически выдаёт IP, шлюз и DNS всем 9 компьютерам"],
          ["Шлюз по умолчанию", "адрес .1 в каждой сети — через него ПК выходят за пределы своего отдела"],
          ["Выход наружу", "свободные порты G0/1 и G0/2 — для подключения провайдера и интернета"]],
        specs: ["3 × Gigabit Ethernet", "4 слота EHWIC", "Серия ISR G2"],
        net: "R1 · порт G0/0 → SW-CORE · подынтерфейсы G0/0.10, .20, .30, .99" },
      { key: "switch", color: C.core, role: "Коммутатор", does: [
          ["Подключает устройства", "24 порта для ПК, принтеров и сервера, ещё 2 гигабитных порта для связи с ядром"],
          ["Разделяет отделы", "VLAN изолируют трафик бухгалтерии, продаж и IT на одном оборудовании"],
          ["Передаёт VLAN дальше", "trunk-каналы 802.1Q несут трафик всех отделов по одному кабелю"],
          ["Управляется удалённо", "у каждого коммутатора свой IP в VLAN 99 для настройки и мониторинга"]],
        specs: ["24 × Fast Ethernet", "2 × Gigabit uplink", "Управляемый L2"],
        net: "SW-CORE — ядро сети; SW-BUH, SW-SALES, SW-IT — по одному в каждом отделе" },
      { key: "pc", color: C.sales, role: "Рабочая станция", does: [
          ["Рабочее место сотрудника", "учёт и отчётность в бухгалтерии, работа с клиентами в продажах, администрирование в IT"],
          ["Автоматическая настройка", "при включении получает IP-адрес по DHCP от маршрутизатора R1"],
          ["Доступ к сервисам", "открывает файлы и внутренний портал на сервере, печатает на МФУ отдела"],
          ["Запас по мощности", "Core i5, 16 ГБ и SSD хватит на офисные программы на 5+ лет"]],
        specs: ["Core i5 / 16 ГБ", "SSD 512 ГБ", "Монитор 24\""],
        net: "PC1–PC3 — бухгалтерия, PC4–PC6 — продажи, PC7–PC9 — IT-отдел" },
      { key: "server", color: C.srv, role: "Сервер", does: [
          ["Хранит данные", "общие папки отделов и документы в одном защищённом месте"],
          ["DNS", "переводит имена сайтов и сервисов в IP-адреса для всех компьютеров офиса"],
          ["Внутренний веб-портал", "HTTP-сервер для новостей, инструкций и заявок сотрудников"],
          ["Резервные копии", "ночные бэкапы данных — защита от потери информации"]],
        specs: ["Стоечный 1U", "Intel Xeon E-2300", "ECC-память"],
        net: "SERVER · 192.168.99.10 · VLAN 99 · порт SW-CORE Fa0/13" },
      { key: "mfp", color: C.it, role: "МФУ", does: [
          ["Печать", "сетевой лазерный принтер общий для всех сотрудников отдела"],
          ["Сканирование", "сканирует документы сразу в папку отдела на сервере"],
          ["Копирование", "копии договоров и отчётов без отдельного копира"],
          ["Сетевое подключение", "подключается в свободный порт коммутатора и попадает в VLAN своего отдела"]],
        specs: ["Лазерная печать", "Сканер + копир", "Ethernet"],
        net: "По одному МФУ на отдел: к SW-BUH, SW-SALES и SW-IT" },
      { key: "rack", color: C.buh, role: "Серверный шкаф", does: [
          ["Размещает оборудование", "R1, SW-CORE, сервер, патч-панель и ИБП в одном месте"],
          ["Защищает", "запирается на ключ — к оборудованию нет доступа у посторонних"],
          ["Охлаждает", "перфорированные двери и место для вентиляторов отводят тепло"],
          ["Упорядочивает кабели", "кабели от розеток приходят на патч-панель, дальше — короткие патч-корды"]],
        specs: ["19 дюймов", "18U", "Замок + вентиляция"],
        net: "Занято около 7U из 18U — остаётся запас под расширение" },
      { key: "ups", color: C.core, role: "Источник бесперебойного питания", does: [
          ["Защищает от отключений", "при пропадании света питает оборудование от батарей"],
          ["Стабилизирует напряжение", "AVR выравнивает скачки и просадки в электросети"],
          ["Корректное выключение", "даёт время сохранить данные и выключить сервер без потерь"],
          ["Бережёт технику", "фильтрует помехи и продлевает жизнь блокам питания"]],
        specs: ["1500 ВА / 1000 Вт", "Стабилизатор AVR", "Для стойки"],
        net: "Питает ядро сети: R1, SW-CORE и сервер" },
    ];
    for (let i = 0; i < details.length; i++) {
      const d = details[i], [name, q, p] = equipment[i];
      const s = light();
      s.addText(`Оборудование ${i + 1} из ${details.length}`, { x: 0.5, y: 0.28, w: 4, h: 0.25, fontFace: FONT_B, fontSize: 11, bold: true, color: d.color, margin: 0, isTextBox: true });
      s.addText(d.role, { x: 0.5, y: 0.52, w: 9, h: 0.55, fontFace: FONT_H, fontSize: 28, bold: true, color: C.navy, margin: 0, isTextBox: true });
      s.addText(name, { x: 0.5, y: 1.05, w: 9, h: 0.3, fontFace: FONT_B, fontSize: 13, color: C.muted, margin: 0, isTextBox: true });
      // photo card
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 1.5, w: 4.1, h: 2.75, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.1, shadow: shadow() });
      await photo(s, d.key, 0.65, 1.62, 3.8, 2.5);
      // price chip
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y: 4.4, w: 4.1, h: 0.85, fill: { color: d.color }, line: { color: d.color }, rectRadius: 0.1 });
      s.addText(`${q} шт. × ${rub(p)}`, { x: 0.7, y: 4.4, w: 2.0, h: 0.85, fontFace: FONT_B, fontSize: 13, color: C.white, margin: 0, valign: "middle", isTextBox: true });
      s.addText(rub(q * p), { x: 2.5, y: 4.4, w: 1.95, h: 0.85, fontFace: FONT_H, fontSize: 22, bold: true, color: C.white, align: "right", margin: 0, valign: "middle", isTextBox: true });
      // what it does
      s.addText("Что делает", { x: 5.0, y: 1.45, w: 4.5, h: 0.3, fontFace: FONT_H, fontSize: 15, bold: true, color: C.navy, margin: 0, isTextBox: true });
      d.does.forEach(([h, t], k) => {
        const y = 1.82 + k * 0.6;
        s.addShape(pres.shapes.OVAL, { x: 5.0, y: y + 0.05, w: 0.14, h: 0.14, fill: { color: d.color }, line: { color: d.color } });
        s.addText([{ text: h + ": ", options: { bold: true, color: C.ink } }, { text: t, options: { color: C.ink } }],
          { x: 5.25, y, w: 4.25, h: 0.55, fontFace: FONT_B, fontSize: 10.5, margin: 0, valign: "top", isTextBox: true });
      });
      // specs chips
      d.specs.forEach((t, k) => {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0 + k * 1.5, y: 4.27, w: 1.4, h: 0.32, fill: { color: C.bg }, line: { color: d.color, width: 1 }, rectRadius: 0.08 });
        s.addText(t, { x: 5.0 + k * 1.5, y: 4.27, w: 1.4, h: 0.32, fontFace: FONT_B, fontSize: 9, bold: true, color: C.ink, align: "center", valign: "middle", margin: 0, isTextBox: true });
      });
      // role in our network
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.0, y: 4.7, w: 4.5, h: 0.55, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.08 });
      s.addText([{ text: "В нашей сети: ", options: { bold: true, color: "CADCFC" } }, { text: d.net, options: { color: C.white } }],
        { x: 5.15, y: 4.7, w: 4.25, h: 0.55, fontFace: FONT_B, fontSize: 9.5, margin: 0, valign: "middle", isTextBox: true });
    }
  }

  // 9. Consumables (cards with photos)
  {
    const s = light(); title(s, "Расходные материалы", "Для прокладки СКС и подключения рабочих мест");
    const keys = ["cable", "rj45", "boot", "panel", "socket", "duct", "cord", "console", "crimp", "ties"];
    for (let i = 0; i < consumables.length; i++) {
      const [n, q, p] = consumables[i];
      const x = 0.5 + (i % 3) * 3.05, y = 1.35 + Math.floor(i / 3) * 1.0;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 2.9, h: 0.93, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.08, shadow: shadow() });
      await photo(s, keys[i], x + 0.08, y + 0.08, 0.95, 0.77);
      s.addText(n, { x: x + 1.12, y: y + 0.08, w: 1.7, h: 0.5, fontFace: FONT_B, fontSize: 9, color: C.ink, margin: 0, valign: "top", isTextBox: true });
      s.addText([{ text: `${q} × ${rub(p)}`, options: { color: C.muted } }, { text: "  " + rub(q * p), options: { bold: true, color: C.buh } }],
        { x: x + 1.12, y: y + 0.6, w: 1.75, h: 0.25, fontFace: FONT_B, fontSize: 9, margin: 0, valign: "middle", isTextBox: true });
    }
    const x = 0.5 + 1 * 3.05, y = 1.35 + 3 * 1.0;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.95, h: 0.93, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.08 });
    s.addText("Итого расходные материалы", { x: x + 0.25, y, w: 3.4, h: 0.93, fontFace: FONT_B, fontSize: 15, color: C.white, margin: 0, valign: "middle", isTextBox: true });
    s.addText(rub(csTotal), { x: x + 3.6, y, w: 2.1, h: 0.93, fontFace: FONT_H, fontSize: 22, bold: true, color: C.white, align: "right", margin: 0, valign: "middle", isTextBox: true });
  }

  // 10. Budget
  {
    const s = light(); title(s, "Бюджет проекта");
    const labels = ["Рабочие станции", "Сервер", "Коммутаторы", "МФУ", "Маршрутизатор", "Шкаф + ИБП", "Расходники"];
    const vals = [equipment[2], equipment[3], equipment[1], equipment[4], equipment[0]].map((r) => r[1] * r[2]);
    vals.push(equipment[5][1] * equipment[5][2] + equipment[6][1] * equipment[6][2], csTotal);
    s.addChart(pres.charts.DOUGHNUT, [{ name: "Бюджет", labels, values: vals }], {
      x: 0.3, y: 1.0, w: 5.2, h: 4.4, holeSize: 55, chartColors: [C.it, C.srv, C.core, C.sales, C.navy, "F7C948", C.buh],
      showLegend: true, legendPos: "r", legendFontSize: 11, legendFontFace: FONT_B, showValue: false, showPercent: true, dataLabelColor: C.white, dataLabelFontSize: 10, showTitle: false,
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 5.9, y: 1.3, w: 3.6, h: 1.7, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.12 });
    s.addText("Общая стоимость", { x: 6.15, y: 1.45, w: 3.2, h: 0.35, fontFace: FONT_B, fontSize: 14, color: "CADCFC", margin: 0, isTextBox: true });
    s.addText(rub(total), { x: 6.15, y: 1.85, w: 3.2, h: 0.8, fontFace: FONT_H, fontSize: 30, bold: true, color: C.white, margin: 0, isTextBox: true });
    [["Оборудование", eqTotal, C.core], ["Расходные материалы", csTotal, C.buh]].forEach(([l, v, c], i) => {
      s.addShape(pres.shapes.OVAL, { x: 5.95, y: 3.37 + i * 0.6, w: 0.22, h: 0.22, fill: { color: c }, line: { color: c } });
      s.addText(l, { x: 6.3, y: 3.28 + i * 0.6, w: 1.9, h: 0.4, fontFace: FONT_B, fontSize: 13, color: C.ink, margin: 0, isTextBox: true });
      s.addText(rub(v), { x: 8.0, y: 3.28 + i * 0.6, w: 1.5, h: 0.4, fontFace: FONT_B, fontSize: 13, bold: true, color: C.ink, align: "right", margin: 0, isTextBox: true });
    });
  }

  // 11. Verification
  {
    const s = light(); title(s, "Проверка работы сети", "Команда ping с рабочих станций в Packet Tracer");
    const tests = [
      ["PC9 (IT)", "PC1 — бухгалтерия 192.168.10.11", C.it],
      ["PC9 (IT)", "PC5 — продажи 192.168.20.12", C.it],
      ["PC9 (IT)", "SERVER 192.168.99.10", C.it],
      ["Trunk-каналы", "802.1Q, все линки up", C.core],
      ["Все ПК", "получили IP по DHCP", C.srv],
    ];
    tests.forEach(([a, b, c], i) => {
      const y = 1.4 + i * 0.72;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.5, y, w: 5.6, h: 0.58, fill: { color: C.bg }, line: { color: C.bg }, rectRadius: 0.08 });
      s.addShape(pres.shapes.OVAL, { x: 0.65, y: y + 0.17, w: 0.24, h: 0.24, fill: { color: c }, line: { color: c } });
      s.addText([{ text: a + "  →  ", options: { bold: true } }, { text: b }], { x: 1.05, y, w: 4.0, h: 0.58, fontFace: FONT_B, fontSize: 13, color: C.ink, margin: 0, valign: "middle", isTextBox: true });
      s.addImage({ data: I.check, x: 5.1, y: y + 0.12, w: 0.34, h: 0.34 });
      s.addText("OK", { x: 5.5, y, w: 0.5, h: 0.58, fontFace: FONT_B, fontSize: 13, bold: true, color: C.buh, margin: 0, valign: "middle", isTextBox: true });
    });
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.4, w: 3.0, h: 3.46, fill: { color: C.navy }, line: { color: C.navy }, rectRadius: 0.1 });
    s.addText("C:\\>ping 192.168.10.11\n\nRequest timed out.\nReply from 192.168.10.11:\n bytes=32 time<1ms TTL=127\nReply from 192.168.10.11:\n bytes=32 time<1ms TTL=127\n\nSent = 4, Received = 3\n(1-й пакет — ARP-запрос)",
      { x: 6.7, y: 1.55, w: 2.7, h: 3.2, fontFace: FONT_C, fontSize: 10, color: "D6F5FF", margin: 0, valign: "top", isTextBox: true });
  }

  // 12. Conclusion
  {
    const s = pres.addSlide(); s.background = { color: C.navy };
    s.addText("Итоги", { x: 0.6, y: 0.5, w: 8, h: 0.8, fontFace: FONT_H, fontSize: 38, bold: true, color: C.white, margin: 0, isTextBox: true });
    const res = [
      [I.sitemap, C.core, "Спроектирована ЛВС на 3 отдела: 9 ПК, сервер, 4 коммутатора, маршрутизатор"],
      [I.layer, C.it, "Отделы разделены по VLAN, маршрутизация между ними — на R1"],
      [I.plug, C.buh, "Адреса раздаются по DHCP, связь между всеми узлами проверена ping"],
      [I.ruble, C.sales, "Составлена смета: " + rub(total) + " на оборудование и материалы"],
    ];
    res.forEach(([img, c, t], i) => {
      circle(s, img, c, 0.6, 1.55 + i * 0.9, 0.6);
      s.addText(t, { x: 1.4, y: 1.55 + i * 0.9, w: 8, h: 0.6, fontFace: FONT_B, fontSize: 16, color: C.white, margin: 0, valign: "middle", isTextBox: true });
    });
  }

  // 13. Image credits
  {
    const s = light(); title(s, "Источники изображений", "Фотографии оборудования — Wikimedia Commons, свободные лицензии");
    const labels = { router: "Маршрутизатор", switch: "Коммутаторы", pc: "Рабочая станция", server: "Сервер", mfp: "МФУ", rack: "Шкаф", ups: "ИБП",
      cable: "Кабель UTP", rj45: "Коннекторы RJ-45", boot: "Колпачки", panel: "Патч-панель", socket: "Розетка RJ-45", duct: "Кабель-канал",
      cord: "Патч-корд", console: "Консольный кабель", crimp: "Кримпер и тестер", ties: "Стяжки" };
    const clean = (t) => t.replace(/\s+/g, " ").replace(/Original uploader.*$/i, "").replace(/Later versi.*$/i, "").trim().slice(0, 40);
    const items = Object.keys(labels).map((k) => [
      { text: labels[k] + ": ", options: { bold: true, color: C.navy } },
      { text: `${credits[k].file.replace(/^File:/, "")} — ${clean(credits[k].artist)}, ${credits[k].license}`, options: { color: C.ink } },
    ]);
    const half = Math.ceil(items.length / 2);
    [items.slice(0, half), items.slice(half)].forEach((col, ci) => {
      const runs = [];
      col.forEach((it, i) => { runs.push(it[0]); runs.push({ ...it[1], options: { ...it[1].options, breakLine: i < col.length - 1 } }); });
      s.addText(runs, { x: 0.5 + ci * 4.6, y: 1.35, w: 4.4, h: 3.3, fontFace: FONT_B, fontSize: 8.5, margin: 0, valign: "top", paraSpaceAfter: 4, isTextBox: true });
    });
    s.addText("Страница каждого файла: commons.wikimedia.org/wiki/File:<имя файла>", { x: 0.5, y: 4.9, w: 9, h: 0.3, fontFace: FONT_B, fontSize: 10, italic: true, color: C.muted, margin: 0, isTextBox: true });
  }

  await pres.writeFile({ fileName: OUT });
  console.log("written", OUT, "total", total, eqTotal, csTotal);
})();
