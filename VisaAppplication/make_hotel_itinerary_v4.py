# -*- coding: utf-8 -*-
"""生成 v3：读取 v2 中用户补充的酒店 Confirmation Number，按模板2格式输出 v3
- 连续两天同酒店 → 合并 Accommodation 单元格
- 行程文字中去掉 "pre-booked / 已预订" 等预订状态标注
"""
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

SRC = "酒店行程单_Italy-France_2026-11_v2.xlsx"
OUT = "酒店行程单_Italy-France_2026-11_v4.xlsx"

# ---------- 1. Confirmation Number（来自 Notion「2026/11/20 Italy & France」页） ----------
conf = {
    "kopster": "1786542555",          # Kopster Hotel Paris Porte de Versailles
    "holidayinn": "1786542555",       # Holiday Inn —— 注意：与 Kopster 相同，疑似复制笔误，待用户核实
    "laperouse": "1073655965",        # Hôtel La Pérouse, Nice
    "novotel": "1784924929",          # Novotel Suites Nice Aéroport Arenas
    "loggiato": "2584028666",         # Hotel Loggiato Dei Serviti, Florence
    "cimatori": "1128151787402789",   # B&B Cimatori, Florence
    "indigo": "82880388",             # Hotel Indigo Florence
    "rome": "1788240878",             # Dimora Degli Dei, Rome
}

# ---------- 2. 组装酒店信息（名称/地址/电话/邮箱/确认号） ----------
def hotel_block(key, fallback):
    lines = [l for l in fallback if l.strip()]
    blk = "\n".join(lines)
    if key in conf:
        blk += f"\nConfirmation Number: {conf[key]}"
    return blk

HOTELS = {
    "kopster":    ["Kopster Hotel Paris Porte de Versailles",
                   "Address: 1 Boulevard Victor, 75015 Paris, France",
                   "Phone: +33 1 87 44 29 61",
                   "Email: reservation.pdv@kopsterhotels.com"],
    "holidayinn": ["Holiday Inn Paris - Gare de Lyon Bastille",
                   "Address: 11 Rue de Lyon, 75012 Paris, France",
                   "Phone: +33 1 53 02 20 00",
                   "Email: re@holidayinnparisbastille.com"],
    "laperouse":  ["Hôtel La Pérouse",
                   "Address: 11 Quai Rauba Capeu, 06300 Nice, France",
                   "Phone: +33 4 93 62 34 63",
                   "Email: reservation@hotel-la-perouse.com"],
    "novotel":    ["Novotel Suites Nice Aéroport Arenas",
                   "Address: 125 Boulevard René Cassin, 06200 Nice, France",
                   "Phone: +33 4 92 29 41 00",
                   "Email: h6513@accor.com"],
    "loggiato":   ["Hotel Loggiato Dei Serviti",
                   "Address: Piazza della Santissima Annunziata 3, 50122 Florence, Italy",
                   "Phone: +39 055 289592",
                   "Email: info@loggiatodeiservitihotel.it"],
    "cimatori":   ["B&B Cimatori",
                   "Address: Via Dante Alighieri 14, 50122 Florence, Italy",
                   "Phone: +39 333 522 5561",
                   "Email: info@cimatori.it"],
    "indigo":     ["Hotel Indigo Florence",
                   "Address: Via Jacopo da Diacceto 16/20, 50123 Florence, Italy",
                   "Phone: +39 055 27390",
                   "Email: reception@indigoflorence.com"],
    "rome":       ["Dimora Degli Dei",
                   "Address: Via Del Seminario 87, 00186 Rome, Italy",
                   "Phone: +39 06 6880 9533",
                   "Email: info@pantheondimoradeglidei.com"],
    "flight":     ["On board flight"],
    "home":       ["Home in Shanghai"],
}

def clean(text):
    """去掉 pre-booked / 已预订 等预订状态标注"""
    text = re.sub(r"\s*[,(]?\s*(?:morning|evening)?\s*,?\s*pre-booked\s*\)?", "", text, flags=re.I)
    text = re.sub(r"（[^）]*已预订[^）]*）", "", text)
    text = re.sub(r"\(\s*,?\s*\)", "", text)
    text = re.sub(r"，\s*，", "，", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()

HEADERS = ["Day", "Date", "Country", "City", "Touring spot", "Transportation", "Accommodation"]

ROWS = [
    (1, "2026/11/20\nFriday", "China - France",
     "Shanghai - Paris (France)\n上海-巴黎（法国）",
     "Departure from Shanghai\n上海出发\nShanghai Hongqiao Airport T2\n上海虹桥机场T2",
     "Flight Shanghai - Beijing\nAir China CA1566 22:30 (Hongqiao T2 → Beijing T3)\nFlight Beijing - Paris\nAir China CA875 02:45 (Beijing T3 → Paris CDG T1), arrive 07:00 (Nov 21)",
     "flight"),
    (2, "2026/11/21\nSaturday", "France",
     "Paris\n巴黎",
     "Arrive in Paris & transfer to hotel\n抵达巴黎，前往酒店办理入住\nLe Marais stroll\n玛黑区漫步\nMusée d'Orsay\n奥赛博物馆\nEiffel Tower\n埃菲尔铁塔\nSeine river cruise (19:00)\n塞纳河游船（19:00）",
     "Flight + Public transportation & walk",
     "kopster"),
    (3, "2026/11/22\nSunday", "France",
     "Paris\n巴黎",
     "Louvre Museum (morning, pre-booked)\n卢浮宫（上午，已预订）\nTuileries Garden\n杜乐丽花园\nPlace de la Concorde\n协和广场\nHôtel de la Marine\n海军府\nPalais Garnier (evening)\n巴黎歌剧院（晚）",
     "Public transportation & walk",
     "kopster"),
    (4, "2026/11/23\nMonday", "France",
     "Paris\n巴黎",
     "Luxembourg Gardens (morning)\n卢森堡公园（上午）\nShakespeare and Company\n莎士比亚书店\nNotre-Dame de Paris\n巴黎圣母院\nMontmartre & Sacré-Cœur (evening)\n蒙马特 & 圣心大教堂（晚）",
     "Public transportation & walk",
     "holidayinn"),
    (5, "2026/11/24\nTuesday", "France",
     "Paris – Nice\n巴黎-尼斯",
     "Paris → Nice (OUIGO train)\n巴黎→尼斯（OUIGO高铁）\nHotel check-in\n酒店入住\nVieux Nice & Promenade des Anglais\n尼斯老城 & 英国人漫步大道\nCastle Hill\n城堡山",
     "1. Train OUIGO Paris Gare de Lyon → Nice Ville 06:03-12:38\n2. Public transportation & walk",
     "laperouse"),
    (6, "2026/11/25\nWednesday", "France",
     "Nice → Antibes → Nice\n尼斯-昂蒂布-尼斯",
     "Train to Antibes (approx. 20 min)\n火车前往昂蒂布（约20分钟）\nAntibes old town & coastline\n昂蒂布老城 & 海岸线\nCannes (optional)\n戛纳（可选）\nReturn to Nice\n返回尼斯",
     "1. TER train Nice → Antibes\n2. Public transportation & walk",
     "novotel"),
    (7, "2026/11/26\nThursday", "France - Italy",
     "Nice – Florence\n尼斯-佛罗伦萨",
     "FlixBus Nice Airport T1 → Florence Villa Costanza\n巴士前往佛罗伦萨\nTram to city center & hotel check-in\n电车入城，酒店入住\nPiazza della Signoria\n领主广场\nPonte Vecchio\n老桥\nSanta Croce district\n圣十字区",
     "1. FlixBus Nice → Florence\n2. Tram\n3. Walk",
     "loggiato"),
    (8, "2026/11/27\nFriday", "Italy",
     "Florence\n佛罗伦萨",
     "Florence Cathedral (morning)\n圣母百花大教堂（上午）\nLunch in Santa Croce\n圣十字区午餐\nPiazza della Signoria\n领主广场\nPonte Vecchio\n老桥\nWalk along the Arno\n沿阿诺河漫步",
     "Walk",
     "cimatori"),
    (9, "2026/11/28\nSaturday", "Italy",
     "Florence\n佛罗伦萨",
     "Uffizi Gallery (morning, pre-booked)\n乌菲兹美术馆（上午，已预订）\nLunch in city center\n市中心午餐\nPiazzale Michelangelo (sunset)\n米开朗基罗广场（日落）",
     "Walk / Bus",
     "indigo"),
    (10, "2026/11/29\nSunday", "Italy",
     "Florence → Rome\n佛罗伦萨-罗马",
     "Train to Rome\n高铁前往罗马\nHotel check-in\n酒店入住\nColosseum (pre-booked)\n罗马斗兽场（已预订）\nPalatine Hill\n帕拉蒂尼山\nRoman Forum\n古罗马广场",
     "1. Train Italo Florence → Roma Termini 10:19\n2. Public transportation & walk",
     "rome"),
    (11, "2026/11/30\nMonday", "Italy",
     "Rome\n罗马",
     "Borghese Gallery (morning, pre-booked)\n博尔盖塞美术馆（上午，已预订）\nPantheon\n万神殿\nPiazza Navona\n纳沃纳广场\nTrevi Fountain\n特莱维喷泉\nSpanish Steps\n西班牙台阶\nPincio (sunset)\n苹丘（日落）",
     "Public transportation & walk",
     "rome"),
    (12, "2026/12/01\nTuesday", "Italy - China",
     "Rome – Shanghai\n罗马-上海",
     "Vatican Museums & Sistine Chapel (pre-booked)\n梵蒂冈博物馆 & 西斯廷教堂（已预订）\nCastel Sant'Angelo & Ponte Sant'Angelo\n圣天使城堡 & 圣天使桥\nSt. Peter's Basilica\n圣彼得大教堂\nGo to airport and back to China in the evening\n傍晚前往机场返回中国",
     "1. Public transportation & walk\n2. Leonardo Express Rome Termini → FCO T3 15:30\n3. Flight Rome - Beijing\nAir China CA940 19:50 (Rome FCO T3 → Beijing T3), arrive 13:00 (Dec 2)\nFlight Beijing - Shanghai\nAir China CA1515 15:45 (Beijing T3 → Shanghai Hongqiao T2), arrive 17:55",
     "flight"),
    (13, "2026/12/02\nWednesday", "China",
     "Shanghai\n上海",
     "Arrive Shanghai Hongqiao T2 17:55\n抵达上海虹桥T2 17:55\nBack to Home\n回家",
     "Flight",
     "home"),
]

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Trip Itinerary"

thin = Side(style="thin")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)

ws.merge_cells("A1:G1")
c = ws.cell(row=1, column=1, value="Trip Itinerary")
c.font = Font(bold=True, size=16)
c.alignment = center
ws.row_dimensions[1].height = 28

for col, h in enumerate(HEADERS, 1):
    cell = ws.cell(row=2, column=col, value=h)
    cell.font = Font(bold=True, size=10)
    cell.alignment = center
    cell.border = border
ws.row_dimensions[2].height = 22

for r, (day, date, country, city, touring, transport, hkey) in enumerate(ROWS, 3):
    values = [day, date, country, city, clean(touring), clean(transport),
              hotel_block(hkey, HOTELS[hkey])]
    for col, val in enumerate(values, 1):
        cell = ws.cell(row=r, column=col, value=val)
        cell.font = Font(size=9)
        cell.alignment = center if col <= 4 else left
        cell.border = border
    ws.row_dimensions[r].height = 130

# ---------- 3. 合并连续同酒店单元格（仅 Accommodation 列） ----------
r = 3
while r < 3 + len(ROWS):
    hkey = ROWS[r - 3][6]
    r2 = r
    while r2 + 1 < 3 + len(ROWS) and ROWS[r2 + 1 - 3][6] == hkey:
        r2 += 1
    if r2 > r:
        ws.merge_cells(start_row=r, start_column=7, end_row=r2, end_column=7)
    r = r2 + 1

widths = [5, 13, 12, 17, 40, 30, 38]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = "A3"
ws.print_area = f"A1:G{2 + len(ROWS)}"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

wb.save(OUT)
print("saved:", OUT)
