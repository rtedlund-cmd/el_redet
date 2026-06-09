import io
from datetime import datetime

import pandas as pd
import streamlit as st

st.set_page_config(page_title="QT Debiteringsunderlag", layout="wide")

MONTHS_SE = {
    1: "Januari", 2: "Februari", 3: "Mars", 4: "April", 5: "Maj", 6: "Juni",
    7: "Juli", 8: "Augusti", 9: "September", 10: "Oktober", 11: "November", 12: "December"
}


def parse_swedish_number(value):
    """Tolkar svenska tal, t.ex. '1 234,56' -> 1234.56."""
    if pd.isna(value):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(" ", "").replace("\u00a0", "")
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def read_qt_csv(uploaded_file):
    """
    Läser QT CSV-export.

    Viktigt: Vi tar med ALLA rader där Enhet = kWh.
    Vi filtrerar alltså INTE på att mätarnamnet slutar på EL, eftersom vissa hus
    kan ha förbrukning både på t.ex. 24 och 24EL. Båda ska debiteras.
    """
    raw = uploaded_file.getvalue()
    text = None
    for encoding in ["utf-8-sig", "utf-8", "latin1"]:
        try:
            text = raw.decode(encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        text = raw.decode("utf-8", errors="ignore")

    df = pd.read_csv(io.StringIO(text), sep=",", dtype=str)

    required = {"Objekt-ID", "Mätare", "Startdatum", "Slutdatum", "Förbrukning", "Enhet"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Saknade kolumner: {', '.join(sorted(missing))}")

    df["Objekt-ID"] = df["Objekt-ID"].astype(str).str.strip()
    df["Mätare"] = df["Mätare"].astype(str).str.strip()
    df["Enhet"] = df["Enhet"].astype(str).str.strip()
    df["Förbrukning_kWh"] = df["Förbrukning"].apply(parse_swedish_number)
    df["Startdatum_dt"] = pd.to_datetime(df["Startdatum"], errors="coerce")
    df["Slutdatum_dt"] = pd.to_datetime(df["Slutdatum"], errors="coerce")

    # Ta med alla kWh-rader. Detta är hela poängen med v2.
    rows = df[df["Enhet"].str.lower().eq("kwh")].copy()
    rows["Filnamn"] = uploaded_file.name

    first_date = rows["Startdatum_dt"].dropna().min()
    if pd.isna(first_date):
        first_date = datetime.today()

    month_key = f"{first_date.year}-{first_date.month:02d}"
    month_label = f"{MONTHS_SE[first_date.month]} {first_date.year}"
    rows["MånadNyckel"] = month_key
    rows["Månad"] = month_label

    return rows


def to_excel(monthly_result, summary_result, details_result, file_check_result):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_result.to_excel(writer, sheet_name="Sammanfattning", index=False)
        monthly_result.to_excel(writer, sheet_name="Per månad", index=False)
        details_result.to_excel(writer, sheet_name="Detaljer", index=False)
        file_check_result.to_excel(writer, sheet_name="Kontroll per fil", index=False)

        for sheet_name in writer.sheets:
            ws = writer.sheets[sheet_name]
            ws.freeze_panes = "A2"
            for col in ws.columns:
                col_letter = col[0].column_letter
                max_len = max(len(str(cell.value)) if cell.value is not None else 0 for cell in col)
                ws.column_dimensions[col_letter].width = min(max_len + 2, 45)
    output.seek(0)
    return output


st.title("QT Debiteringsunderlag")
st.write(
    "Ladda upp en eller flera CSV-filer från QT Systems. Appen summerar alla rader där Enhet = kWh "
    "per fastighet och månad. Ange elpris per månad och exportera Excel-underlag."
)

uploaded_files = st.file_uploader(
    "Dra in CSV-filer här",
    type=["csv"],
    accept_multiple_files=True,
)

if not uploaded_files:
    st.info("Ladda upp t.ex. jan.csv, feb.csv och mars.csv för att börja.")
    st.stop()

all_rows = []
errors = []
seen_files = set()

for file in uploaded_files:
    # En enkel varning mot exakt samma filnamn två gånger.
    if file.name in seen_files:
        st.warning(f"Filen verkar uppladdad flera gånger och ignoreras: {file.name}")
        continue
    seen_files.add(file.name)

    try:
        all_rows.append(read_qt_csv(file))
    except Exception as exc:
        errors.append(f"{file.name}: {exc}")

if errors:
    st.error("Några filer kunde inte läsas:\n" + "\n".join(errors))

if not all_rows:
    st.stop()

details = pd.concat(all_rows, ignore_index=True)

# Kontroll per fil så man direkt ser totalsumma från varje QT-export.
file_check = (
    details.groupby(["Filnamn", "MånadNyckel", "Månad"], as_index=False)
    .agg(Antal_rader=("Förbrukning_kWh", "size"), Total_kWh=("Förbrukning_kWh", "sum"))
    .sort_values(["MånadNyckel", "Filnamn"])
)

months = (
    details[["MånadNyckel", "Månad"]]
    .drop_duplicates()
    .sort_values("MånadNyckel")
)

st.subheader("1. Ange elpris")
price_map = {}
cols = st.columns(min(len(months), 4) or 1)
for i, row in enumerate(months.itertuples(index=False)):
    with cols[i % len(cols)]:
        price_map[row.MånadNyckel] = st.number_input(
            f"{row.Månad} kr/kWh",
            min_value=0.0,
            value=0.0,
            step=0.01,
            format="%.4f",
        )

details["Pris kr/kWh"] = details["MånadNyckel"].map(price_map).fillna(0.0)
details["Kostnad kr"] = details["Förbrukning_kWh"] * details["Pris kr/kWh"]

monthly = (
    details.groupby(["Objekt-ID", "MånadNyckel", "Månad"], as_index=False)
    .agg({"Förbrukning_kWh": "sum", "Pris kr/kWh": "max", "Kostnad kr": "sum"})
    .sort_values(["Objekt-ID", "MånadNyckel"])
    .rename(columns={"Objekt-ID": "Fastighet", "Förbrukning_kWh": "kWh"})
)

summary = (
    monthly.groupby("Fastighet", as_index=False)
    .agg({"kWh": "sum", "Kostnad kr": "sum"})
    .sort_values("Fastighet")
)

# Avrundade kopior för visning och export. Rådata behåller tre decimaler för kWh.
monthly_display = monthly.copy()
summary_display = summary.copy()
details_display = details.copy()
file_check_display = file_check.copy()

for frame in [monthly_display, summary_display, details_display, file_check_display]:
    for col in ["kWh", "Förbrukning_kWh", "Total_kWh"]:
        if col in frame.columns:
            frame[col] = frame[col].round(3)
    for col in ["Pris kr/kWh", "Kostnad kr"]:
        if col in frame.columns:
            frame[col] = frame[col].round(2)

st.subheader("2. Sammanfattning per fastighet")
st.dataframe(summary_display, use_container_width=True, hide_index=True)

st.subheader("3. Per månad")
st.dataframe(monthly_display, use_container_width=True, hide_index=True)

with st.expander("Kontroll per fil"):
    st.dataframe(file_check_display, use_container_width=True, hide_index=True)

with st.expander("Visa importerade rader"):
    st.dataframe(details_display, use_container_width=True, hide_index=True)

excel_file = to_excel(monthly_display, summary_display, details_display, file_check_display)
st.download_button(
    "Ladda ner Excel-underlag",
    data=excel_file,
    file_name="debiteringsunderlag_qt.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
