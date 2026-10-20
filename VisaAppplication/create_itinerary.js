import fs from "node:fs";
import path from "node:path";
import {
  AlignmentType,
  Document,
  Footer,
  Header,
  HeadingLevel,
  Packer,
  PageNumber,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const outputPath = process.argv[2];
if (!outputPath) throw new Error("Usage: node create_itinerary.js /absolute/path/output.docx");

const T = String.raw;
const palette = {
  dark: "1F3A5F",
  primary: "2E5A88",
  light: "7A8AA0",
  border: "C9D4E0",
  fill: "EDF2F8",
};

const font = {
  ascii: "Times New Roman",
  hAnsi: "Times New Roman",
  cs: "Times New Roman",
  eastAsia: "SimSun",
};

const run = (text, options = {}) =>
  new TextRun({ text, font, size: 21, ...options });

// English run: smaller, gray, italic
const runEn = (text) =>
  new TextRun({ text, font, size: 18, italics: true, color: palette.primary });

const para = (children, options = {}) =>
  new Paragraph({
    spacing: { after: 120, line: 276 },
    ...options,
    children: Array.isArray(children) ? children : [children],
  });

const title = "申根签证行程单（法国）";
const titleEn = "Travel Itinerary for Schengen Visa (France)";

// ---------- applicant info (bilingual) ----------
// each row: [labelZh, labelEn, valueParagraphs: [[zh,en], ...]]
const infoRows = [
  [T`申请人`, T`Applicants`, [
    [T`孟子豪 Zihao Meng（护照号 EB7540007）；朱雪艳 Xueyan Zhu（护照号 EM3984549）`,
     T`Zihao MENG (Passport No. EB7540007); Xueyan ZHU (Passport No. EM3984549)`],
  ]],
  [T`签证类型`, T`Visa Type`, [
    [T`申根短期旅游签证（申请国：法国）`, T`Schengen short-stay tourist visa (applied via France)`],
  ]],
  [T`行程时间`, T`Travel Period`, [
    [T`2026 年 11 月 20 日 — 12 月 2 日（共 13 天）`, T`November 20 – December 2, 2026 (13 days)`],
  ]],
  [T`行程路线`, T`Route`, [
    [T`上海 → 巴黎 → 尼斯 → 昂蒂布 → 佛罗伦萨 → 罗马 → 上海`,
     T`Shanghai → Paris → Nice → Antibes → Florence → Rome → Shanghai`],
  ]],
  [T`国际航班`, T`International Flights`, [
    [T`去程 11/20：CA1566 上海虹桥 T2 22:30 → 北京首都 T3 00:55(+1)；CA875 北京首都 T3 02:45 → 巴黎戴高乐 T1 07:00（11/21）`,
     T`Outbound Nov 20: CA1566 Shanghai Hongqiao T2 22:30 → Beijing Capital T3 00:55(+1); CA875 Beijing T3 02:45 → Paris CDG T1 07:00 (Nov 21)`],
    [T`回程 12/1：CA940 罗马菲乌米奇诺 T3 19:50 → 北京首都 T3 13:00(+1)；12/2：CA1515 北京首都 T3 15:45 → 上海虹桥 T2 17:55`,
     T`Return Dec 1: CA940 Rome Fiumicino T3 19:50 → Beijing T3 13:00(+1); Dec 2: CA1515 Beijing T3 15:45 → Shanghai Hongqiao T2 17:55`],
    [T`航空公司：中国国际航空；订座记录号（PNR）：HRD7DD`,
     T`Airline: Air China; PNR: HRD7DD`],
  ]],
];

// ---------- itinerary data ----------
// [dateZh, dateEn, cityZh, cityEn, [ [zh,en], ... ], [hotelZh, hotelEn]]
const days = [
  [
    T`11/20（周五）`, T`20 Nov (Fri)`,
    T`上海 → 巴黎`, T`Shanghai → Paris`,
    [
      [T`晚间自上海虹桥国际机场 T2 乘 CA1566（22:30）飞往北京，转机后乘 CA875（02:45）飞往巴黎。`,
       T`Evening: fly from Shanghai Hongqiao Int'l Airport T2 to Beijing on CA1566 (dep. 22:30), then connect to CA875 (dep. 02:45) to Paris.`],
    ],
    [T`飞机上`, T`On board flight`],
  ],
  [
    T`11/21（周六）`, T`21 Nov (Sat)`,
    T`巴黎`, T`Paris`,
    [
      [T`约 07:00 抵达巴黎戴高乐机场 T1，办理入境后前往酒店（车程约 3 小时），入住后稍事休息。`,
       T`Approx. 07:00 arrival at Paris Charles de Gaulle Airport T1; transfer to hotel (approx. 3 hours) and check in.`],
      [T`上午：玛黑区（Le Marais）逛街；中午在玛黑区用餐。`,
       T`Morning: stroll and shopping in Le Marais; lunch in the Marais district.`],
      [T`下午：奥赛博物馆（Musée d'Orsay，已预约）。`,
       T`Afternoon: Musée d'Orsay (pre-booked).`],
      [T`傍晚：埃菲尔铁塔（Tour Eiffel）；19:00 塞纳河游船。`,
       T`Evening: Eiffel Tower; 19:00 Seine river cruise.`],
    ],
    [T`科普斯特巴黎凡尔赛门酒店（Kopster Hotel Paris Porte de Versailles）`, T`Kopster Hotel Paris Porte de Versailles`],
  ],
  [
    T`11/22（周日）`, T`22 Nov (Sun)`,
    T`巴黎`, T`Paris`,
    [
      [T`上午：卢浮宫（Musée du Louvre，已预约，预留一上午）。`,
       T`Morning: Louvre Museum (pre-booked; allow a full morning).`],
      [T`中午：卢浮宫附近用餐。`,
       T`Lunch near the Louvre.`],
      [T`下午：杜乐丽花园（Jardin des Tuileries）、协和广场（Place de la Concorde）、海军府（Hôtel de la Marine）。`,
       T`Afternoon: Tuileries Garden, Place de la Concorde, Hôtel de la Marine.`],
      [T`傍晚：巴黎歌剧院（Palais Garnier）参观。`,
       T`Evening: visit of Palais Garnier (Paris Opera).`],
    ],
    [T`同前（巴黎）`, T`Same as above (Paris)`],
  ],
  [
    T`11/23（周一）`, T`23 Nov (Mon)`,
    T`巴黎`, T`Paris`,
    [
      [T`上午：卢森堡公园（Jardin du Luxembourg，约 2 小时）。`,
       T`Morning: Luxembourg Gardens (approx. 2 hours).`],
      [T`中午：拉丁区用餐。`,
       T`Lunch in the Latin Quarter.`],
      [T`下午：莎士比亚书店（Shakespeare and Company）、巴黎圣母院（Cathédrale Notre-Dame de Paris）。`,
       T`Afternoon: Shakespeare and Company bookstore, Notre-Dame de Paris Cathedral.`],
      [T`傍晚：蒙马特高地（Montmartre，圣心堂）。`,
       T`Evening: Montmartre hill (Sacré-Cœur Basilica).`],
    ],
    [T`巴黎里昂车站巴士底假日酒店（Holiday Inn Paris – Gare de Lyon Bastille, IHG）`, T`Holiday Inn Paris – Gare de Lyon Bastille (IHG)`],
  ],
  [
    T`11/24（周二）`, T`24 Nov (Tue)`,
    T`巴黎 → 尼斯`, T`Paris → Nice`,
    [
      [T`06:03 自巴黎里昂车站（Paris Gare de Lyon）乘火车（OUIGO 7829），12:38 抵达尼斯（Nice Ville），车程约 6 小时。`,
       T`06:03 train from Paris Gare de Lyon (OUIGO 7829), arriving Nice Ville at 12:38 (approx. 6 hours).`],
      [T`下午：入住酒店后游览尼斯老城区（Vieux Nice）、英国人漫步大道（Promenade des Anglais）、城堡山（Colline du Château）俯瞰天使湾。`,
       T`Afternoon: check in, then explore Vieux Nice, the Promenade des Anglais, and Castle Hill overlooking the Baie des Anges.`],
    ],
    [T`拉彼鲁兹酒店（Hôtel La Pérouse, Nice）`, T`Hôtel La Pérouse, Nice`],
  ],
  [
    T`11/25（周三）`, T`25 Nov (Wed)`,
    T`尼斯 → 昂蒂布 → 尼斯`, T`Nice → Antibes → Nice`,
    [
      [T`上午：乘区域火车前往昂蒂布（Antibes，约 20 分钟，班次密集），游览昂蒂布老城与海岸。`,
       T`Morning: regional train to Antibes (approx. 20 min, frequent departures); explore the old town and coastline.`],
      [T`下午：视体力前往戛纳（Cannes）游览，傍晚返回尼斯。`,
       T`Afternoon: optional visit to Cannes; return to Nice in the evening.`],
    ],
    [T`尼斯机场诺富特全套房酒店（Novotel Suites Nice Aeroport Arenas）`, T`Novotel Suites Nice Aeroport Arenas`],
  ],
  [
    T`11/26（周四）`, T`26 Nov (Thu)`,
    T`尼斯 → 佛罗伦萨`, T`Nice → Florence`,
    [
      [T`自尼斯机场 T1 乘 FlixBus 大巴前往佛罗伦萨，抵达 Villa Costanza 车站后乘电车（约 30 分钟）进入市中心。`,
       T`FlixBus coach from Nice Airport Terminal 1 to Florence; on arrival at Villa Costanza station, take the tram to the city center (approx. 30 min).`],
      [T`下午：入住酒店后游览领主广场（Piazza della Signoria）、老桥（Ponte Vecchio）、圣十字街区（Santa Croce）。`,
       T`Afternoon: check in, then visit Piazza della Signoria, Ponte Vecchio, and the Santa Croce district.`],
    ],
    [T`洛贾托迪赛维帝酒店（Hotel Loggiato Dei Serviti）`, T`Hotel Loggiato Dei Serviti`],
  ],
  [
    T`11/27（周五）`, T`27 Nov (Fri)`,
    T`佛罗伦萨`, T`Florence`,
    [
      [T`上午：圣母百花大教堂（Cattedrale di Santa Maria del Fiore）。`,
       T`Morning: Florence Cathedral (Santa Maria del Fiore).`],
      [T`中午：圣十字街区用餐。`,
       T`Lunch in the Santa Croce district.`],
      [T`下午：领主广场、老桥（桥南岸可拍阿诺河倒影），沿阿诺河漫步。`,
       T`Afternoon: Piazza della Signoria and Ponte Vecchio (riverside reflections from the south bank); walk along the Arno.`],
    ],
    [T`奇马托利 B&B 民宿（B&B Cimatori）`, T`B&B Cimatori`],
  ],
  [
    T`11/28（周六）`, T`28 Nov (Sat)`,
    T`佛罗伦萨`, T`Florence`,
    [
      [T`上午：乌菲兹美术馆（Galleria degli Uffizi）。`,
       T`Morning: Uffizi Gallery.`],
      [T`中午：市区用餐。`,
       T`Lunch in the city center.`],
      [T`傍晚：米开朗基罗广场（Piazzale Michelangelo）看日落，俯瞰全城。`,
       T`Evening: sunset at Piazzale Michelangelo with panoramic views over the city.`],
    ],
    [T`佛罗伦萨英迪格酒店（Hotel Indigo Florence, IHG）`, T`Hotel Indigo Florence (IHG)`],
  ],
  [
    T`11/29（周日）`, T`29 Nov (Sun)`,
    T`佛罗伦萨 → 罗马`, T`Florence → Rome`,
    [
      [T`乘 Italo 高铁前往罗马，10:19 抵达罗马 Termini 车站。`,
       T`Italo high-speed train to Rome, arriving Roma Termini at 10:19.`],
      [T`中午：车站附近用餐。`,
       T`Lunch near Termini station.`],
      [T`下午：罗马斗兽场（Colosseo，已预约）、帕拉丁山（Palatine Hill，自 Vignola 门进入）、古罗马广场（Foro Romano），合计约 3 小时。`,
       T`Afternoon: Colosseum (pre-booked), Palatine Hill (enter via Portale del Vignola), Roman Forum — approx. 3 hours in total.`],
    ],
    [T`罗马酒店（名称及地址待补充）`, T`Hotel in Rome (name & address to be confirmed)`],
  ],
  [
    T`11/30（周一）`, T`30 Nov (Mon)`,
    T`罗马`, T`Rome`,
    [
      [T`上午：博尔盖塞美术馆（Galleria Borghese，已预约）。`,
       T`Morning: Borghese Gallery (pre-booked).`],
      [T`中午：市区用餐。`,
       T`Lunch in the city center.`],
      [T`下午：万神殿（Pantheon）、圣伊纳爵教堂（Chiesa di Sant'Ignazio）、纳沃纳广场（Piazza Navona）、特雷维喷泉（Fontana di Trevi）、西班牙广场（Piazza di Spagna）。`,
       T`Afternoon: Pantheon, Church of Sant'Ignazio, Piazza Navona, Trevi Fountain, Spanish Steps.`],
      [T`傍晚：苹丘（Passeggiata del Pincio）看日落。`,
       T`Evening: sunset at Passeggiata del Pincio.`],
    ],
    [T`同前（罗马）`, T`Same as above (Rome)`],
  ],
  [
    T`12/1（周二）`, T`1 Dec (Tue)`,
    T`罗马 → 上海`, T`Rome → Shanghai`,
    [
      [T`上午：梵蒂冈博物馆与西斯廷教堂（Musei Vaticani & Cappella Sistina，已预约）。`,
       T`Morning: Vatican Museums & Sistine Chapel (pre-booked).`],
      [T`中午：圣天使城堡（Castel Sant'Angelo）、圣天使桥（Ponte Sant'Angelo）附近用餐。`,
       T`Lunch near Castel Sant'Angelo and Ponte Sant'Angelo.`],
      [T`下午：圣彼得大教堂（Basilica di San Pietro）。`,
       T`Afternoon: St. Peter's Basilica.`],
      [T`15:30 自罗马 Termini 车站乘 Leonardo Express，16:30 抵达菲乌米奇诺机场（FCO）T3；19:50 乘 CA940 返回北京。`,
       T`15:30 Leonardo Express from Roma Termini, arriving Fiumicino Airport (FCO) T3 at 16:30; 19:50 depart on CA940 to Beijing.`],
    ],
    [T`飞机上`, T`On board flight`],
  ],
  [
    T`12/2（周三）`, T`2 Dec (Wed)`,
    T`抵达上海`, T`Arrive in Shanghai`,
    [
      [T`13:00 抵达北京首都 T3，转机后乘 CA1515（15:45），17:55 抵达上海虹桥 T2，行程结束。`,
       T`13:00 arrival at Beijing Capital T3; connect to CA1515 (dep. 15:45), arriving Shanghai Hongqiao T2 at 17:55. End of trip.`],
    ],
    [T`—`, T`—`],
  ],
];

const colW = [1150, 1250, 3988, 3250]; // sum 9638
const tableWidth = colW.reduce((a, b) => a + b, 0);

const cellFromPairs = (pairs, i, options = {}) =>
  new TableCell({
    children: pairs.map(([zh, en]) => [
      para([run(zh)], { spacing: { after: 20, line: 252 } }),
      para([runEn(en)], { spacing: { after: 60, line: 240 } }),
    ]).flat(),
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    width: { size: colW[i], type: WidthType.DXA },
    ...options,
  });

const headerCell = (zh, en, i) =>
  new TableCell({
    children: [
      para([run(zh, { bold: true }), runEn(en ? " " + en : "")], { spacing: { after: 0 } }),
    ],
    shading: { type: ShadingType.CLEAR, fill: palette.fill },
    margins: { top: 80, bottom: 80, left: 100, right: 100 },
    width: { size: colW[i], type: WidthType.DXA },
  });

const itineraryTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: colW,
  rows: [
    new TableRow({
      tableHeader: true,
      children: [
        headerCell(T`日期`, T`Date`, 0),
        headerCell(T`城市`, T`City`, 1),
        headerCell(T`当日行程`, T`Daily Itinerary`, 2),
        headerCell(T`住宿`, T`Accommodation`, 3),
      ],
    }),
    ...days.map(
      (d) =>
        new TableRow({
          children: [
            cellFromPairs([[d[0], d[1]]], 0),
            cellFromPairs([[d[2], d[3]]], 1),
            cellFromPairs(d[4], 2),
            cellFromPairs([d[5]], 3),
          ],
        })
    ),
  ],
});

const infoTable = new Table({
  width: { size: tableWidth, type: WidthType.DXA },
  columnWidths: [1800, 7838],
  rows: infoRows.map(
    (r) =>
      new TableRow({
        children: [
          new TableCell({
            children: [
              para([run(r[0], { bold: true })], { spacing: { after: 20 } }),
              para([runEn(r[1])], { spacing: { after: 0 } }),
            ],
            shading: { type: ShadingType.CLEAR, fill: palette.fill },
            margins: { top: 80, bottom: 80, left: 100, right: 100 },
            width: { size: 1800, type: WidthType.DXA },
          }),
          cellFromPairs(r[2], 1),
        ],
      })
  ),
});

const notePairs = [
  [T`1. 本行程单航班信息已与机票预订单一致（中国国际航空，订座记录号 HRD7DD）。`,
   T`Flight details are consistent with the flight reservation (Air China, PNR HRD7DD).`],
  [T`2. 城际交通已出票：巴黎—尼斯（OUIGO 火车，11/24）、尼斯—佛罗伦萨（FlixBus 大巴，11/26）、佛罗伦萨—罗马（Italo 高铁，11/29）。`,
   T`Intercity tickets booked: Paris–Nice (OUIGO train, Nov 24), Nice–Florence (FlixBus coach, Nov 26), Florence–Rome (Italo high-speed train, Nov 29).`],
  [T`3. 罗马两晚（11/29–11/30）酒店名称及地址待补充，送签前需与酒店预订单保持一致。`,
   T`The Rome hotel for Nov 29–30 is to be confirmed and must match the hotel reservation before submission.`],
  [T`4. 首入国为法国（巴黎），法国与意大利停留时间相当，按行程申请法国申根签证。`,
   T`First entry is France (Paris); stays in France and Italy are comparable, so the Schengen visa is applied for through France.`],
];

const children = [
  para([run(title, { bold: true, size: 36, color: palette.dark })], {
    alignment: AlignmentType.CENTER,
    spacing: { after: 40 },
  }),
  para([runEn(titleEn)], {
    alignment: AlignmentType.CENTER,
    spacing: { after: 280 },
  }),
  infoTable,
  para([run(T`每日行程 `, { bold: true, size: 26, color: palette.dark }), runEn(T`Daily Itinerary`)], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
  }),
  itineraryTable,
  para([run(T`备注 `, { bold: true, size: 26, color: palette.dark }), runEn(T`Notes`)], {
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
  }),
  ...notePairs.map(([zh, en]) =>
    para([run(zh), runEn(" " + en)], { spacing: { after: 100 } })
  ),
];

const doc = new Document({
  features: { updateFields: false },
  sections: [
    {
      properties: {
        page: {
          size: { width: 11906, height: 16838 },
          margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 },
        },
      },
      headers: {
        default: new Header({
          children: [
            para([run(title, { size: 18, color: palette.light }), runEn(" " + titleEn)], {
              alignment: AlignmentType.CENTER,
              spacing: { after: 0 },
            }),
          ],
        }),
      },
      footers: {
        default: new Footer({
          children: [
            para(new TextRun({ children: [PageNumber.CURRENT], font, size: 18 }), {
              alignment: AlignmentType.CENTER,
            }),
          ],
        }),
      },
      children,
    },
  ],
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync(outputPath, buffer);
