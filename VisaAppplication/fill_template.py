# -*- coding: utf-8 -*-
"""按《法国签证行程单模板.xlsx》填充真实行程数据，生成送签用行程单。"""
import copy
import openpyxl
from openpyxl.utils import get_column_letter

SRC = r"F:/Project/VisaApplicationByKimi/Visa Appplication/法国签证行程单模板.xlsx"
OUT = r"F:/Project/VisaApplicationByKimi/Visa Appplication/法签行程单_按模板_Italy-France_2026-11.xlsx"

wb = openpyxl.load_workbook(SRC)

# ---------------- Sheet 1: 申请人信息 ----------------
ws1 = wb["申请人信息"]
ws1["B4"] = "Zihao Meng / 孟子豪"
ws1["B5"] = "EB7540007"
ws1["B6"] = "20/ 11 / 2026  – 2 / 12 / 2026"
ws1["B7"] = "France（法国）"
ws1["B8"] = "Paris"
ws1["B9"] = "Roma"
ws1["B10"] = ("Kopster Hotel Paris Porte de Versailles, 1 Boulevard Victor, 75015 Paris（11/21–22）；"
              "Holiday Inn Paris – Gare de Lyon Bastille, 11 Rue de Lyon, 75012 Paris（11/23）")
ws1["B11"] = ("去程：CA1566 上海虹桥T2 22:30 → 北京；CA875 北京 02:45 → 巴黎戴高乐T1 07:00（11/21）。"
              "回程：CA940 罗马菲乌米奇诺T3 19:50 → 北京（12/1）；CA1515 北京 15:45 → 上海虹桥T2 17:55（12/2）。"
              "订座记录号 PNR：HRD7DD")
ws1["B12"] = "Xueyan Zhu / 朱雪艳（Passport No. EM3984549）"

# ---------------- Sheet 2: 每日行程 ----------------
ws2 = wb["每日行程 Itinerary"]

rows = [
    ("2026-11-20", "Fri", "Shanghai → Paris",
     "Evening: depart Shanghai Hongqiao T2 on CA1566 (22:30) to Beijing; connect to CA875 (02:45) to Paris",
     "On board flight", "Flight CA1566 + CA875", "出发日"),
    ("2026-11-21", "Sat", "Paris",
     "Arrive CDG T1 07:00 → transfer to hotel & check-in → Le Marais stroll & lunch → Musée d'Orsay → Eiffel Tower → Seine river cruise (19:00)",
     "Kopster Hotel Paris Porte de Versailles, 1 Boulevard Victor, 75015 Paris", "Flight + Metro", "入境日"),
    ("2026-11-22", "Sun", "Paris",
     "Louvre Museum (morning, pre-booked) → Tuileries Garden & Place de la Concorde → Hôtel de la Marine → Palais Garnier (evening)",
     "Kopster Hotel Paris Porte de Versailles, 1 Boulevard Victor, 75015 Paris", "Metro / Walk", ""),
    ("2026-11-23", "Mon", "Paris",
     "Luxembourg Gardens (morning) → Shakespeare and Company & Notre-Dame → Montmartre & Sacré-Cœur (evening)",
     "Holiday Inn Paris – Gare de Lyon Bastille, 11 Rue de Lyon, 75012 Paris", "Metro", "换酒店"),
    ("2026-11-24", "Tue", "Paris → Nice",
     "OUIGO train 06:03 Paris Gare de Lyon → Nice Ville 12:38 → hotel check-in → Vieux Nice & Promenade des Anglais → Castle Hill",
     "Hôtel La Pérouse, 11 Quai Rauba Capeu, 06300 Nice", "OUIGO train + Walk", "换城市"),
    ("2026-11-25", "Wed", "Nice → Antibes → Nice",
     "Regional train to Antibes (approx. 20 min) → old town & coastline → optional visit to Cannes → return to Nice",
     "Novotel Suites Nice Aeroport Arenas, 125 Boulevard René Cassin, 06200 Nice", "TER train", ""),
    ("2026-11-26", "Thu", "Nice → Florence",
     "FlixBus from Nice Airport T1 → Florence Villa Costanza → tram to city center → Piazza della Signoria & Ponte Vecchio → Santa Croce district",
     "Hotel Loggiato Dei Serviti, Piazza della Santissima Annunziata 3, 50122 Florence", "FlixBus + Tram", "换城市 / 入境意大利"),
    ("2026-11-27", "Fri", "Florence",
     "Florence Cathedral (morning) → lunch in Santa Croce → Piazza della Signoria & Ponte Vecchio → walk along the Arno",
     "B&B Cimatori, Via Dante Alighieri 14, 50122 Florence", "Walk", "换酒店"),
    ("2026-11-28", "Sat", "Florence",
     "Uffizi Gallery (morning) → lunch in city center → Piazzale Michelangelo for sunset",
     "Hotel Indigo Florence, Via Jacopo da Diacceto 16/20, 50123 Florence", "Walk / Bus", "换酒店"),
    ("2026-11-29", "Sun", "Florence → Rome",
     "Italo high-speed train → Roma Termini 10:19 → lunch near station → Colosseum, Palatine Hill & Roman Forum (pre-booked, approx. 3h)",
     "Hotel in Rome (name & address to be confirmed)", "Italo train", "换城市"),
    ("2026-11-30", "Mon", "Rome",
     "Borghese Gallery (morning, pre-booked) → Pantheon → Piazza Navona & Trevi Fountain → Spanish Steps → Pincio sunset",
     "Hotel in Rome (name & address to be confirmed)", "Walk / Metro", ""),
    ("2026-12-01", "Tue", "Rome → Shanghai",
     "Vatican Museums & Sistine Chapel (pre-booked) → Castel Sant'Angelo & Ponte Sant'Angelo → St. Peter's Basilica → 15:30 Leonardo Express to FCO T3 → CA940 dep. 19:50",
     "On board flight", "Metro + Leonardo Express + Flight CA940", "出境日"),
    ("2026-12-02", "Wed", "Arrive in Shanghai",
     "Arrive Beijing Capital T3 13:00 → connect CA1515 (15:45) → arrive Shanghai Hongqiao T2 17:55. End of trip",
     "—", "Flight CA1515", "返回"),
]

# 模板原有示例行 4–11（8 行），需要 13 行：把第 4 行的样式复制到 12–16 行
template_row = 4
for r in range(12, 4 + len(rows)):
    for c in range(1, 8):
        src_cell = ws2.cell(row=template_row, column=c)
        dst_cell = ws2.cell(row=r, column=c)
        dst_cell._style = copy.copy(src_cell._style)
    ws2.row_dimensions[r].height = ws2.row_dimensions[template_row].height

for i, row in enumerate(rows):
    r = 4 + i
    for c, val in enumerate(row, start=1):
        cell = ws2.cell(row=r, column=c, value=val if val != "" else None)
        # 长文本列自动换行（沿用模板第 4 行对齐设置）
        if c in (4, 5):
            a = copy.copy(cell.alignment)
            a.wrap_text = True
            a.vertical = "top"
            cell.alignment = a
        else:
            a = copy.copy(cell.alignment)
            a.vertical = "top"
            cell.alignment = a

# 清理可能残留的多余行（模板只有 8 行示例，已覆盖）
wb.save(OUT)
print("saved:", OUT)

# 回读校验
wb2 = openpyxl.load_workbook(OUT)
w = wb2["每日行程 Itinerary"]
print("rows with data:", w.max_row)
for r in range(4, w.max_row + 1):
    print(w.cell(row=r, column=1).value, "|", w.cell(row=r, column=2).value, "|",
          str(w.cell(row=r, column=3).value)[:22], "|",
          str(w.cell(row=r, column=5).value)[:60])
