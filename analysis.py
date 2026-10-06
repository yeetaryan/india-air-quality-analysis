"""
India Urban Air Quality Analysis (2015-2020)
--------------------------------------------
1. Cleans CPCB daily city-level air quality data with pandas.
2. Builds an Excel dashboard (output/India_AQI_Dashboard.xlsx) in which every
   summary number is a live Excel formula (AVERAGEIFS / COUNTIFS / RANK / INDEX-MATCH)
   on the cleaned data sheet, so the workbook recalculates if the data changes.
3. Saves PNG charts for the README.

Run:  python analysis.py
"""
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

RAW = "data/raw/city_day.csv"
OUT_XLSX = "output/India_AQI_Dashboard.xlsx"
CLEAN_CSV = "data/city_day_clean.csv"

NAAQS_PM25_24H = 60     # India NAAQS, 24-hour PM2.5 limit (ug/m3)
LOCKDOWN = ("2020-03-25", "2020-05-31")   # national lockdown phases 1-4
BASELINE = ("2019-03-25", "2019-05-31")   # same calendar window, previous year

SEASONS = {12: "Winter", 1: "Winter", 2: "Winter",
           3: "Summer", 4: "Summer", 5: "Summer",
           6: "Monsoon", 7: "Monsoon", 8: "Monsoon", 9: "Monsoon",
           10: "Post-Monsoon", 11: "Post-Monsoon"}   # IMD season definitions

# ---------------------------------------------------------------- 1. clean
df = pd.read_csv(RAW, parse_dates=["Date"])
raw_rows = len(df)
assert df.duplicated(["City", "Date"]).sum() == 0, "duplicate city-day rows"

cols = ["PM2.5", "PM10", "NO2", "NOx", "CO", "SO2", "O3", "AQI"]
df = df[["City", "Date"] + cols + ["AQI_Bucket"]].copy()
missing_pct = (df[cols].isna().mean() * 100).round(1)
# Missing readings are left blank (not imputed): AVERAGEIFS/COUNTIFS skip blanks,
# so averages are based only on days a station actually reported.
df["Year"] = df["Date"].dt.year
df["Month"] = df["Date"].dt.month
df["Season"] = df["Month"].map(SEASONS)
df = df.sort_values(["City", "Date"]).reset_index(drop=True)
df.to_csv(CLEAN_CSV, index=False)

print(f"Rows: {raw_rows:,} | Cities: {df.City.nunique()} | "
      f"{df.Date.min():%d %b %Y} to {df.Date.max():%d %b %Y}")
print("Missing % by pollutant:\n", missing_pct.to_string())

# Cities with PM2.5 readings in BOTH lockdown and baseline windows
def window(a, b):
    return df[(df.Date >= a) & (df.Date <= b)]
base_c = set(window(*BASELINE).dropna(subset=["PM2.5"]).City)
lock_c = set(window(*LOCKDOWN).dropna(subset=["PM2.5"]).City)
lockdown_cities = sorted(base_c & lock_c)

city_order = (df.groupby("City")["PM2.5"].mean()
                .sort_values(ascending=False).index.tolist())

# ---------------------------------------------------------------- 2. Excel
FONT = "Arial"
HDR_FILL = PatternFill("solid", fgColor="1F4E78")
HDR_FONT = Font(name=FONT, bold=True, color="FFFFFF")
KPI_FILL = PatternFill("solid", fgColor="DDEBF7")
INPUT_FONT = Font(name=FONT, color="0000FF")
thin = Side(style="thin", color="BFBFBF")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

def header(ws, row, labels, col=1):
    for i, h in enumerate(labels):
        c = ws.cell(row=row, column=col + i, value=h)
        c.font, c.fill, c.alignment, c.border = HDR_FONT, HDR_FILL, Alignment(horizontal="center", wrap_text=True), BOX

def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x

wb = Workbook()
dash = wb.active
dash.title = "Dashboard"

# ---- Data sheet
data = wb.create_sheet("Data")
dcols = ["City", "Date", "Year", "Month", "Season"] + cols + ["AQI_Bucket", "Exceeds_NAAQS"]
header(data, 1, dcols)
n = len(df)
last = n + 1
for r, row in enumerate(df.itertuples(index=False), start=2):
    data.cell(r, 1, row.City)
    data.cell(r, 2, row.Date.to_pydatetime()).number_format = "yyyy-mm-dd"
    data.cell(r, 3, row.Year)
    data.cell(r, 4, row.Month)
    data.cell(r, 5, row.Season)
    for j, c in enumerate(cols):
        v = df.at[r - 2, c]
        if pd.notna(v):
            data.cell(r, 6 + j, float(v))
    if pd.notna(row.AQI_Bucket):
        data.cell(r, 14, row.AQI_Bucket)
    # 1 if the day's PM2.5 exceeded the 24-h NAAQS, blank if no reading
    data.cell(r, 15, f'=IF(F{r}="","",IF(F{r}>Notes!$B$4,1,0))')
data.freeze_panes = "A2"
data.auto_filter.ref = f"A1:O{last}"
widths(data, [16, 12, 7, 7, 13] + [8] * 8 + [13, 14])
for row in data.iter_rows(min_row=2, max_row=last):
    for c in row:
        c.font = Font(name=FONT, size=10)

R = lambda col: f"Data!${col}$2:${col}${last}"
CITY, DATE, MON, SEA, PM, NO2, AQI, BKT, EXC = (R("A"), R("B"), R("D"), R("E"),
                                               R("F"), R("H"), R("M"), R("N"), R("O"))

# ---- Notes / assumptions
notes = wb.create_sheet("Notes")
notes["A1"] = "Assumptions, sources and definitions"
notes["A1"].font = Font(name=FONT, bold=True, size=13)
rows = [
    ("Source", "Central Pollution Control Board (CPCB) daily city data, compiled as 'Air Quality Data in India (2015-2020)' on Kaggle; file city_day.csv"),
    ("Coverage", f"{raw_rows:,} city-day records, {df.City.nunique()} cities, {df.Date.min():%d %b %Y} - {df.Date.max():%d %b %Y}"),
    ("PM2.5 24-h NAAQS (ug/m3)", NAAQS_PM25_24H),
    ("Lockdown window start", pd.Timestamp(LOCKDOWN[0]).to_pydatetime()),
    ("Lockdown window end", pd.Timestamp(LOCKDOWN[1]).to_pydatetime()),
    ("Baseline window start", pd.Timestamp(BASELINE[0]).to_pydatetime()),
    ("Baseline window end", pd.Timestamp(BASELINE[1]).to_pydatetime()),
    ("Missing values", "Left blank, not imputed. Averages use reporting days only. Missing share: "
        + ", ".join(f"{k} {v}%" for k, v in missing_pct.items())),
    ("Seasons (IMD)", "Winter Dec-Feb | Summer Mar-May | Monsoon Jun-Sep | Post-Monsoon Oct-Nov"),
    ("Caveat", "Cities joined the dataset at different times (9 in 2015, 26 by 2020), so all-city yearly averages are not like-for-like; trends are shown per city."),
    ("Blue text", "Input cells - change them and every sheet recalculates."),
]
for i, (k, v) in enumerate(rows, start=2):
    notes.cell(i, 1, k).font = Font(name=FONT, bold=True)
    c = notes.cell(i, 2, v)
    c.font = INPUT_FONT if i in (4, 5, 6, 7, 8) else Font(name=FONT)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    if i in (5, 6, 7, 8):
        c.number_format = "yyyy-mm-dd"
notes.column_dimensions["A"].width = 26
notes.column_dimensions["B"].width = 100

# ---- City ranking
cr = wb.create_sheet("City_Ranking")
header(cr, 1, ["City", "Avg PM2.5 (ug/m3)", "Avg NO2 (ug/m3)", "Avg AQI",
               "Days with PM2.5 reading", "Days above NAAQS", "% days above NAAQS",
               "x NAAQS annual (40)", "Rank (PM2.5)"])
for i, city in enumerate(city_order, start=2):
    cr.cell(i, 1, city)
    cr.cell(i, 2, f"=AVERAGEIFS({PM},{CITY},A{i})").number_format = "0.0"
    cr.cell(i, 3, f'=IFERROR(AVERAGEIFS({NO2},{CITY},A{i}),"")').number_format = "0.0"
    cr.cell(i, 4, f'=IFERROR(AVERAGEIFS({AQI},{CITY},A{i}),"")').number_format = "0"
    cr.cell(i, 5, f'=COUNTIFS({CITY},A{i},{PM},"<>")')
    cr.cell(i, 6, f"=COUNTIFS({CITY},A{i},{EXC},1)")
    cr.cell(i, 7, f"=F{i}/E{i}").number_format = "0.0%"
    cr.cell(i, 8, f"=B{i}/40").number_format = '0.0"x"'
    cr.cell(i, 9, f"=RANK(B{i},$B$2:$B${len(city_order)+1})")
nc = len(city_order) + 1
widths(cr, [20, 14, 14, 10, 14, 14, 14, 14, 10])
cr.freeze_panes = "A2"

# ---- Seasonality
se = wb.create_sheet("Seasonality")
header(se, 1, ["Month", "Month name", "Avg PM2.5 - all cities", "Avg PM2.5 - Delhi", "Avg NO2 - all cities"])
mnames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
for m in range(1, 13):
    r = m + 1
    se.cell(r, 1, m)
    se.cell(r, 2, mnames[m - 1])
    se.cell(r, 3, f"=AVERAGEIFS({PM},{MON},A{r})").number_format = "0.0"
    se.cell(r, 4, f'=AVERAGEIFS({PM},{MON},A{r},{CITY},"Delhi")').number_format = "0.0"
    se.cell(r, 5, f"=AVERAGEIFS({NO2},{MON},A{r})").number_format = "0.0"
header(se, 16, ["Season", "Avg PM2.5 - all cities"])
for i, s in enumerate(["Winter", "Summer", "Monsoon", "Post-Monsoon"], start=17):
    se.cell(i, 1, s)
    se.cell(i, 2, f'=AVERAGEIFS({PM},{SEA},A{i})').number_format = "0.0"
se.cell(22, 1, "Winter vs Monsoon (x)").font = Font(name=FONT, bold=True)
se.cell(22, 2, "=B17/B19").number_format = '0.0"x"'
widths(se, [22, 12, 20, 18, 20])

# ---- Lockdown impact
lk = wb.create_sheet("Lockdown_Impact")
header(lk, 1, ["City", "PM2.5 2019 window", "PM2.5 2020 lockdown", "PM2.5 change %",
               "NO2 2019 window", "NO2 2020 lockdown", "NO2 change %"])
B0, B1, L0, L1 = "Notes!$B$7", "Notes!$B$8", "Notes!$B$5", "Notes!$B$6"
for i, city in enumerate(lockdown_cities, start=2):
    lk.cell(i, 1, city)
    for col, rng in ((2, PM), (5, NO2)):
        lk.cell(i, col, f'=IFERROR(AVERAGEIFS({rng},{CITY},$A{i},{DATE},">="&{B0},{DATE},"<="&{B1}),"")').number_format = "0.0"
        lk.cell(i, col + 1, f'=IFERROR(AVERAGEIFS({rng},{CITY},$A{i},{DATE},">="&{L0},{DATE},"<="&{L1}),"")').number_format = "0.0"
        a, b = get_column_letter(col), get_column_letter(col + 1)
        lk.cell(i, col + 2, f'=IFERROR({b}{i}/{a}{i}-1,"")').number_format = "0.0%"
lr = len(lockdown_cities) + 1
t = lr + 2
lk.cell(t, 1, "Average of cities").font = Font(name=FONT, bold=True)
for col in (2, 3, 5, 6):
    L = get_column_letter(col)
    lk.cell(t, col, f"=AVERAGE({L}2:{L}{lr})").number_format = "0.0"
lk.cell(t, 4, f"=C{t}/B{t}-1").number_format = "0.0%"
lk.cell(t, 7, f"=F{t}/E{t}-1").number_format = "0.0%"
lk.cell(t + 1, 1, "Cities improving (PM2.5)").font = Font(name=FONT, bold=True)
lk.cell(t + 1, 4, f'=COUNTIF(D2:D{lr},"<0")&" of "&COUNT(D2:D{lr})')
lk.cell(t + 1, 7, f'=COUNTIF(G2:G{lr},"<0")&" of "&COUNT(G2:G{lr})')
widths(lk, [20, 16, 18, 15, 15, 17, 13])

# ---- Dashboard
dash["A1"] = "India Urban Air Quality Dashboard (2015-2020)"
dash["A1"].font = Font(name=FONT, bold=True, size=16, color="1F4E78")
dash["A2"] = "Source: CPCB daily city data. All figures are live formulas - see Notes for definitions."
dash["A2"].font = Font(name=FONT, italic=True, size=9)
kpis = [
    ("City-day records", f"=COUNTA({CITY})", "#,##0"),
    ("Cities covered", f"=COUNTA(City_Ranking!A2:A{nc})", "0"),
    ("Most polluted city", f"=INDEX(City_Ranking!A2:A{nc},MATCH(1,City_Ranking!I2:I{nc},0))", "@"),
    ("Its avg PM2.5 (ug/m3)", f"=MAX(City_Ranking!B2:B{nc})", "0.0"),
    ("% city-days above NAAQS", f"=SUM({EXC})/COUNT({PM})", "0.0%"),
    ("% days AQI Poor or worse", f'=(COUNTIF({BKT},"Poor")+COUNTIF({BKT},"Very Poor")+COUNTIF({BKT},"Severe"))/COUNTA({BKT})', "0.0%"),
    ("Winter vs Monsoon PM2.5", "=Seasonality!B22", '0.0"x"'),
    ("Lockdown PM2.5 change", f"=Lockdown_Impact!D{t}", "0.0%"),
    ("Lockdown NO2 change", f"=Lockdown_Impact!G{t}", "0.0%"),
]
for i, (k, f, fmt) in enumerate(kpis):
    r = 4 + i
    a = dash.cell(r, 1, k); a.font = Font(name=FONT, bold=True); a.fill = KPI_FILL; a.border = BOX
    b = dash.cell(r, 2, f); b.number_format = fmt; b.font = Font(name=FONT, bold=True, size=12); b.border = BOX
    b.alignment = Alignment(horizontal="right")
widths(dash, [28, 18])

ch = BarChart(); ch.type = "bar"; ch.style = 10
ch.title = "Average PM2.5 by city (ug/m3)"
ch.add_data(Reference(cr, min_col=2, min_row=1, max_row=nc), titles_from_data=True)
ch.set_categories(Reference(cr, min_col=1, min_row=2, max_row=nc))
ch.y_axis.title = "ug/m3"; ch.legend = None
ch.x_axis.scaling.orientation = "maxMin"
ch.height, ch.width = 16, 16; ch.gapWidth = 40
dash.add_chart(ch, "D3")

lc = LineChart(); lc.title = "Monthly PM2.5 pattern"; lc.style = 12
lc.add_data(Reference(se, min_col=3, max_col=4, min_row=1, max_row=13), titles_from_data=True)
lc.set_categories(Reference(se, min_col=2, min_row=2, max_row=13))
lc.y_axis.title = "ug/m3"; lc.height, lc.width = 8, 16
dash.add_chart(lc, "D30")

bc = BarChart(); bc.title = "Lockdown vs 2019 - PM2.5 change by city"; bc.style = 10
bc.add_data(Reference(lk, min_col=4, min_row=1, max_row=lr), titles_from_data=True)
bc.set_categories(Reference(lk, min_col=1, min_row=2, max_row=lr))
bc.legend = None; bc.height, bc.width = 9, 16; bc.y_axis.number_format = "0%"
dash.add_chart(bc, "M3")

for ws in (cr, se, lk):
    for row in ws.iter_rows(min_row=2):
        for c in row:
            if c.font.name != FONT:
                c.font = Font(name=FONT, bold=c.font.bold)
wb._sheets = [wb[n] for n in ["Dashboard", "City_Ranking", "Seasonality", "Lockdown_Impact", "Data", "Notes"]]
wb.save(OUT_XLSX)
print("Saved", OUT_XLSX)

# ---------------------------------------------------------------- 3. README charts
city_pm = df.groupby("City")["PM2.5"].mean().reindex(city_order)
fig, ax = plt.subplots(figsize=(9, 7))
city_pm[::-1].plot.barh(ax=ax, color="#1F4E78")
ax.axvline(40, color="red", ls="--", label="NAAQS annual limit (40)")
ax.set_xlabel("Average PM2.5 (ug/m3), 2015-2020"); ax.set_ylabel(""); ax.legend()
ax.set_title("Average PM2.5 by city")
plt.tight_layout(); plt.savefig("charts/pm25_by_city.png", dpi=130); plt.close()

mon = df.groupby("Month")["PM2.5"].mean()
dmon = df[df.City == "Delhi"].groupby("Month")["PM2.5"].mean()
fig, ax = plt.subplots(figsize=(9, 4))
ax.plot(mnames, mon.values, marker="o", label="All cities")
ax.plot(mnames, dmon.values, marker="o", label="Delhi")
ax.axhline(NAAQS_PM25_24H, color="red", ls="--", label="24-h NAAQS (60)")
ax.set_ylabel("PM2.5 (ug/m3)"); ax.set_title("Seasonality of PM2.5"); ax.legend()
plt.tight_layout(); plt.savefig("charts/seasonality.png", dpi=130); plt.close()

b = window(*BASELINE).groupby("City")[["PM2.5", "NO2"]].mean().loc[lockdown_cities]
l = window(*LOCKDOWN).groupby("City")[["PM2.5", "NO2"]].mean().loc[lockdown_cities]
chg = ((l / b - 1) * 100).sort_values("PM2.5")
fig, ax = plt.subplots(figsize=(9, 6))
chg.plot.barh(ax=ax)
ax.set_xlabel("% change vs same window in 2019"); ax.set_title("Lockdown impact (25 Mar - 31 May 2020)")
plt.tight_layout(); plt.savefig("charts/lockdown_impact.png", dpi=130); plt.close()
print("Charts saved")
