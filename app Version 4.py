# app.py
# Streamlit-App: EVA – Logistik-Kostenvergleichs-Chatbot
# Starten mit: streamlit run app.py

"""
Erwartete Excel-Struktur, passend zu deiner Beispiel-Datei:

Sheet: Grundpreis
- Carrier
- Versandart
- Gewicht max.
- Abmessungen L×B×H  oder  Abmessungen / Regel
- Grundpreis
- Tracking
- Haftung
- Laufzeit
- Empfängertyp
- Auslandsversand

Sheet: Versicherung
- Regel
- Warenwert
- Versicherung
- Aufpreis
- Gesamtkosten-Formel

Sheet: Zusatzservices
- Service
- Bedingung
- Aufpreis

Sheet: Sonderregeln nach Warenart
- Warenart
- DHL erlaubt / DPD erlaubt / GLS erlaubt  oder allgemeine Regelspalten
- Versicherung
- Spezialregel
- Chatbot-Aktion

Sheet: Entscheidungsregeln
- Regel
- Bedingung
- Aktion

Hinweis:
Der Code ist bewusst robust geschrieben. Wenn deine Excel-Datei leicht andere Spaltennamen hat,
versucht EVA die passenden Spalten automatisch zu erkennen. Wenn wichtige Spalten fehlen,
zeigt EVA eine klare Fehlermeldung.
"""

import re
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import streamlit as st
from openai import OpenAI


# ------------------------------------------------------------
# Streamlit-Seitenlayout
# ------------------------------------------------------------
st.set_page_config(
    page_title="EVA – Versandkostenvergleich",
    page_icon="📦",
    layout="wide",
)

st.title("📦 EVA – Logistik-Kostenvergleichs-Chatbot")
st.caption("EVA vergleicht Carrier-Kosten auf Basis deiner Excel-Daten, Versicherungsregeln und Entscheidungslogik.")

# ------------------------------------------------------------
# OpenAI / ChatGPT-Integration
# ------------------------------------------------------------
def get_openai_client():
    """Erstellt den OpenAI-Client nur, wenn ein API-Key vorhanden ist."""
    try:
        api_key = st.secrets.get("OPENAI_API_KEY")
    except Exception:
        api_key = None

    if not api_key:
        return None
    return OpenAI(api_key=api_key)


openai_client = get_openai_client()


# ------------------------------------------------------------
# Hilfsfunktionen: Text, Zahlen, Parsing
# ------------------------------------------------------------
def normalize_text(value) -> str:
    """Bereinigt Textwerte aus Excel und UI."""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_column_name(col: str) -> str:
    """Normalisiert Spaltennamen für flexible Spaltenerkennung."""
    col = normalize_text(col).lower()
    col = col.replace("×", "x")
    col = col.replace("/", " ")
    col = col.replace(".", "")
    col = re.sub(r"\s+", " ", col)
    return col.strip()


def parse_euro(value) -> Optional[float]:
    """Wandelt deutsche Preisangaben wie '6,99 €' oder 'ab 17,00 €' in float um."""
    if pd.isna(value):
        return None
    text = str(value).lower().strip()
    if text in ["", "-", "nan", "individuell", "sonderfall"]:
        return None
    text = text.replace("€", "").replace("eur", "").replace("ab", "")
    text = text.replace(".", "").replace(",", ".")
    match = re.search(r"\d+(?:\.\d+)?", text)
    return float(match.group()) if match else None


def parse_number(value) -> Optional[float]:
    """Extrahiert die erste Zahl aus Text, z. B. '31,5 kg' -> 31.5."""
    if pd.isna(value):
        return None
    text = str(value).replace(",", ".")
    match = re.search(r"\d+(?:\.\d+)?", text)
    return float(match.group()) if match else None


def parse_max_weight(value) -> Optional[float]:
    """Liest maximales Gewicht aus Spalten wie 'Gewicht max.' oder 'bis 20 kg'."""
    return parse_number(value)


def parse_dimensions(value) -> Tuple[Optional[float], Optional[float], Optional[float], Optional[float]]:
    """
    Liest Abmessungen aus verschiedenen Formaten.
    Rückgabe: (max_l, max_b, max_h, max_single_dimension)

    Beispiele:
    - '120×60×60 cm' -> (120, 60, 60, None)
    - 'bis 70 cm' -> (None, None, None, 70)
    """
    if pd.isna(value):
        return None, None, None, None

    text = str(value).lower().replace(",", ".").replace("×", "x")
    numbers = [float(x) for x in re.findall(r"\d+(?:\.\d+)?", text)]

    if "x" in text and len(numbers) >= 3:
        return numbers[0], numbers[1], numbers[2], None

    if "bis" in text and len(numbers) >= 1:
        return None, None, None, numbers[0]

    return None, None, None, None


def get_longest_side(length: float, width: float, height: float) -> float:
    return max(length, width, height)


def format_euro(value: Optional[float]) -> str:
    if value is None or pd.isna(value):
        return "nicht verfügbar"
    return f"{value:,.2f} €".replace(",", "X").replace(".", ",").replace("X", ".")


# ------------------------------------------------------------
# Excel-Import und Spaltenerkennung
# ------------------------------------------------------------
def load_excel(uploaded_file) -> Dict[str, pd.DataFrame]:
    """Lädt alle Excel-Sheets in ein Dictionary."""
    try:
        sheets = pd.read_excel(uploaded_file, sheet_name=None)
        cleaned = {}
        for name, df in sheets.items():
            df = df.copy()
            df.columns = [normalize_text(c) for c in df.columns]
            df = df.dropna(how="all")
            cleaned[normalize_text(name)] = df
        return cleaned
    except Exception as exc:
        st.error(f"Die Excel-Datei konnte nicht gelesen werden: {exc}")
        return {}


def find_sheet(sheets: Dict[str, pd.DataFrame], keywords: List[str]) -> Optional[pd.DataFrame]:
    """Findet ein Sheet anhand von Schlüsselwörtern im Sheetnamen."""
    for sheet_name, df in sheets.items():
        normalized = normalize_column_name(sheet_name)
        if any(keyword in normalized for keyword in keywords):
            return df
    return None


def find_column(df: pd.DataFrame, possible_names: List[str]) -> Optional[str]:
    """Findet eine Spalte anhand möglicher deutscher/englischer Namen."""
    normalized_map = {normalize_column_name(c): c for c in df.columns}
    possible = [normalize_column_name(x) for x in possible_names]

    for p in possible:
        if p in normalized_map:
            return normalized_map[p]

    for norm_col, original_col in normalized_map.items():
        if any(p in norm_col for p in possible):
            return original_col

    return None


def require_columns(df: pd.DataFrame, required: Dict[str, List[str]], sheet_name: str) -> Dict[str, str]:
    """Prüft notwendige Spalten und zeigt verständliche Fehlermeldungen."""
    found = {}
    missing = []
    for key, variants in required.items():
        col = find_column(df, variants)
        if col is None:
            missing.append(f"{key}: erwartet z. B. {', '.join(variants)}")
        else:
            found[key] = col

    if missing:
        st.error(f"Im Sheet '{sheet_name}' fehlen wichtige Spalten.")
        with st.expander("Erwartete Spalten anzeigen"):
            for item in missing:
                st.write("- " + item)
            st.write("Vorhandene Spalten:", list(df.columns))
    return found


# ------------------------------------------------------------
# Eingabevalidierung und Gewicht
# ------------------------------------------------------------
def validate_inputs(length, width, height, raw_weight, sender, receiver, goods_type, goods_value) -> List[str]:
    errors = []
    if length <= 0 or width <= 0 or height <= 0:
        errors.append("Länge, Breite und Höhe müssen größer als 0 sein.")
    if raw_weight <= 0:
        errors.append("Das Rohgewicht muss größer als 0 kg sein.")
    if goods_value < 0:
        errors.append("Der Warenwert darf nicht negativ sein.")
    if not sender.strip():
        errors.append("PLZ/Stadt des Absenders ist eine Pflichtangabe.")
    if not receiver.strip():
        errors.append("PLZ/Stadt des Empfängers ist eine Pflichtangabe.")
    if not goods_type.strip():
        errors.append("Die Warenart ist eine Pflichtangabe.")
    return errors


def calculate_volumetric_weight(length: float, width: float, height: float, divisor: float = 5000) -> float:
    """Standardformel: L × B × H / 5000."""
    return round((length * width * height) / divisor, 2)


def calculate_billable_weight(raw_weight: float, volumetric_weight: float) -> float:
    return max(raw_weight, volumetric_weight)


# ------------------------------------------------------------
# Carrier- und Preislogik
# ------------------------------------------------------------
def prepare_base_price_table(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    required = {
        "carrier": ["Carrier"],
        "service": ["Versandart", "Service", "Produkt"],
        "max_weight": ["Gewicht max", "Gewicht max.", "Max Gewicht"],
        "dimensions": ["Abmessungen LxBxH", "Abmessungen L×B×H", "Abmessungen / Regel", "Paketgröße"],
        "price": ["Grundpreis", "Preis"],
        "tracking": ["Tracking"],
        "liability": ["Haftung"],
        "delivery_time": ["Laufzeit"],
        "international": ["Auslandsversand"],
    }
    cols = require_columns(df, required, "Grundpreis")
    if not cols:
        return df, cols

    prepared = df.copy()
    prepared["_price"] = prepared[cols["price"]].apply(parse_euro)
    prepared["_max_weight"] = prepared[cols["max_weight"]].apply(parse_max_weight)
    dim_values = prepared[cols["dimensions"]].apply(parse_dimensions)
    prepared["_max_l"] = dim_values.apply(lambda x: x[0])
    prepared["_max_b"] = dim_values.apply(lambda x: x[1])
    prepared["_max_h"] = dim_values.apply(lambda x: x[2])
    prepared["_max_single_dim"] = dim_values.apply(lambda x: x[3])
    prepared["_carrier"] = prepared[cols["carrier"]].astype(str).str.strip()
    prepared["_service"] = prepared[cols["service"]].astype(str).str.strip()

    return prepared, cols


def row_fits_dimensions(row, length: float, width: float, height: float) -> bool:
    """Prüft, ob Paketmaße in die Zeile passen."""
    dims = sorted([length, width, height], reverse=True)

    if pd.notna(row.get("_max_l")) and pd.notna(row.get("_max_b")) and pd.notna(row.get("_max_h")):
        max_dims = sorted([row["_max_l"], row["_max_b"], row["_max_h"]], reverse=True)
        return all(d <= m for d, m in zip(dims, max_dims))

    if pd.notna(row.get("_max_single_dim")):
        return get_longest_side(length, width, height) <= row["_max_single_dim"]

    return True


def is_international(receiver: str) -> bool:
    """Sehr einfache Heuristik: Deutschland/deutsche PLZ = national, sonst international."""
    text = receiver.lower().strip()
    if "deutschland" in text or "germany" in text:
        return False
    if re.match(r"^\d{5}\b", text):
        return False
    return True


def carrier_allows_international(row, cols: Dict[str, str], international: bool) -> bool:
    if not international:
        return True
    col = cols.get("international")
    if not col or col not in row:
        return True
    value = normalize_text(row[col]).lower()
    return value in ["ja", "yes", "true", "1", "j"]


def find_best_service_for_carrier(
    base_df: pd.DataFrame,
    cols: Dict[str, str],
    carrier: str,
    length: float,
    width: float,
    height: float,
    billable_weight: float,
    receiver: str,
) -> Optional[pd.Series]:
    """Findet für einen Carrier den günstigsten passenden Service."""
    df = base_df[base_df["_carrier"].str.lower() == carrier.lower()].copy()
    if df.empty:
        return None

    international = is_international(receiver)
    candidates = []
    for _, row in df.iterrows():
        price_ok = pd.notna(row["_price"])
        weight_ok = pd.notna(row["_max_weight"]) and billable_weight <= row["_max_weight"]
        dim_ok = row_fits_dimensions(row, length, width, height)
        intl_ok = carrier_allows_international(row, cols, international)
        if price_ok and weight_ok and dim_ok and intl_ok:
            candidates.append(row)

    if not candidates:
        return None

    result = pd.DataFrame(candidates).sort_values("_price", ascending=True).iloc[0]
    return result


# ------------------------------------------------------------
# Versicherung und Zusatzservices
# ------------------------------------------------------------
def get_liability_limit(liability_text: str) -> Optional[float]:
    """Extrahiert Haftungsgrenze aus Text, z. B. 'bis 500 €'."""
    return parse_number(liability_text)


def calculate_insurance(carrier: str, goods_value: float, base_price: float, liability_text: str, insurance_df: Optional[pd.DataFrame]) -> Dict:
    """Berechnet Versicherungsstatus und Prämie. Nutzt Excel-Regeln, soweit möglich."""
    liability_limit = get_liability_limit(liability_text)
    needs_insurance = False
    recommended = False
    premium = 0.0
    status = "Keine Zusatzversicherung nötig"
    note = "Die Standardhaftung reicht nach den hinterlegten Daten aus."

    if liability_limit is not None and goods_value > liability_limit:
        recommended = True
        needs_insurance = True
        status = "Versicherung empfohlen oder notwendig"
        note = f"Der Warenwert liegt über der Standardhaftung von ca. {format_euro(liability_limit)}."

    carrier_upper = carrier.upper()

    # DHL-Regeln aus typischer Excel-Struktur
    if carrier_upper == "DHL":
        if goods_value <= 500:
            premium = 0.0
            status = "Keine Zusatzversicherung nötig"
        elif goods_value <= 2500:
            premium = 6.99
            recommended = True
            needs_insurance = True
            status = "Transportversicherung bis 2.500 € empfohlen"
        elif goods_value <= 25000:
            premium = 19.99
            recommended = True
            needs_insurance = True
            status = "Transportversicherung bis 25.000 € empfohlen"
        else:
            premium = np.nan
            needs_insurance = True
            status = "Nicht zulässig / manuelle Prüfung"
            note = "Der Warenwert überschreitet die in der DHL-Regel hinterlegte Grenze."

    elif carrier_upper == "DPD":
        if goods_value <= 520:
            premium = 0.0
            status = "Standardhaftung enthalten"
        elif goods_value <= 10000:
            premium = round(max(5.0, goods_value * 0.01), 2)
            recommended = True
            needs_insurance = True
            status = "Höherversicherung empfohlen/erforderlich"
            note = "Für DPD ist in der Excel-Datei häufig 'individuell' hinterlegt. EVA nutzt deshalb eine einfache Schätzung von 1 % des Warenwerts, mindestens 5 €."
        else:
            premium = np.nan
            needs_insurance = True
            status = "Manuelle Prüfung nötig"

    elif carrier_upper == "GLS":
        if goods_value <= 750:
            premium = 0.0
            status = "Standardhaftung vermutlich ausreichend"
        elif goods_value <= 5000:
            premium = round(max(5.0, goods_value * 0.01), 2)
            recommended = True
            needs_insurance = True
            status = "Zusatzversicherung empfohlen"
        else:
            premium = np.nan
            needs_insurance = True
            status = "Sonderprüfung erforderlich"

    total_without = base_price
    total_with = None if pd.isna(premium) else base_price + premium

    return {
        "insurance_relevant": bool(needs_insurance or recommended),
        "insurance_status": status,
        "insurance_premium": premium,
        "total_without_insurance": total_without,
        "total_with_insurance": total_with,
        "insurance_note": note,
    }


def get_goods_rule(goods_type: str, rules_df: Optional[pd.DataFrame]) -> Dict:
    """Sucht Warenart-Regel in 'Sonderregeln nach Warenart'."""
    if rules_df is None or rules_df.empty:
        return {}

    goods_col = find_column(rules_df, ["Warenart", "Art", "Produktart"])
    if not goods_col:
        return {}

    target = goods_type.lower().strip()
    for _, row in rules_df.iterrows():
        excel_goods = normalize_text(row.get(goods_col, "")).lower()
        if target == excel_goods or target in excel_goods or excel_goods in target:
            return {c: row.get(c, "") for c in rules_df.columns}

    return {}


def goods_type_blocks_carrier(goods_type: str, goods_value: float, carrier: str, goods_rule: Dict) -> Tuple[bool, str]:
    """Einfache Sperrlogik für kritische Warenarten und bekannte Regeln."""
    gt = goods_type.lower().strip()
    carrier_upper = carrier.upper()

    critical_over_500 = ["schmuck", "uhr", "bargeld"]
    if any(x in gt for x in critical_over_500) and goods_value > 500 and carrier_upper in ["DHL", "DPD"]:
        return True, f"{carrier} wird für {goods_type} über 500 € nach der Entscheidungslogik abgelehnt."

    if "gefahrgut" in gt:
        return True, f"{goods_type} ist ein Sonderfall/Gefahrgut und muss manuell geprüft werden."

    # Falls Excel explizit eine Spalte wie 'DHL erlaubt' enthält
    for key, value in goods_rule.items():
        key_norm = normalize_column_name(key)
        if carrier.lower() in key_norm and "erlaubt" in key_norm:
            if normalize_text(value).lower() in ["nein", "no", "nicht erlaubt"]:
                return True, f"{carrier} ist laut Warenart-Regel nicht erlaubt."

    return False, ""


def select_additional_services(goods_type: str, goods_value: float, service_df: Optional[pd.DataFrame]) -> List[Dict]:
    """Empfiehlt Zusatzservices abhängig von Warenart und Warenwert."""
    recommendations = []
    gt = goods_type.lower()

    def add(service, reason, estimated_price=0.0):
        recommendations.append({
            "Zusatzleistung": service,
            "Grund": reason,
            "geschätzter Aufpreis": estimated_price,
        })

    if any(x in gt for x in ["laptop", "smartphone", "tablet", "kamera", "elektronik"]):
        add("Sendungsverfolgung / Tracking", "Elektronik hat meist höheren Warenwert und sollte transparent verfolgt werden.", 0.0)
        add("Zusatzversicherung", "Bei Elektronik ist eine Versicherung besonders empfehlenswert.", 0.0)

    if any(x in gt for x in ["buch", "bücher"]):
        add("Günstige Waren-/Päckchen-Option prüfen", "Bücher sind oft klein und leicht.", 0.0)

    if goods_value > 500:
        add("Transportversicherung", "Der Warenwert liegt über einer typischen Standardhaftungsgrenze.", 0.0)

    if not recommendations:
        add("Standardversand", "Für diese Warenart reicht meistens ein normaler Paketservice.", 0.0)

    return recommendations


# ------------------------------------------------------------
# Gesamte Berechnung
# ------------------------------------------------------------
def calculate_all_carriers(
    sheets: Dict[str, pd.DataFrame],
    length: float,
    width: float,
    height: float,
    raw_weight: float,
    sender: str,
    receiver: str,
    goods_type: str,
    goods_value: float,
) -> Tuple[pd.DataFrame, List[Dict], float, float]:
    base_df_raw = find_sheet(sheets, ["grundpreis", "ex"])
    insurance_df = find_sheet(sheets, ["versicherung"])
    service_df = find_sheet(sheets, ["zusatz"])
    goods_rules_df = find_sheet(sheets, ["sonderregeln", "warenart"])

    if base_df_raw is None:
        st.error("Kein Sheet für Grundpreise gefunden. Erwarteter Sheetname z. B. 'Grundpreis'.")
        return pd.DataFrame(), [], 0, 0

    base_df, cols = prepare_base_price_table(base_df_raw)
    if not cols:
        return pd.DataFrame(), [], 0, 0

    # Für die Projektlogik wird das Rohgewicht als Abrechnungsgewicht verwendet.
    # Die Maße werden separat geprüft, um passende Paketgrößen auszuwählen.
    volumetric_weight = 0.0
    billable_weight = raw_weight

    carriers = sorted([c for c in base_df["_carrier"].dropna().unique() if c and c.lower() != "carrier"])
    goods_rule = get_goods_rule(goods_type, goods_rules_df)
    additional_services = select_additional_services(goods_type, goods_value, service_df)

    rows = []
    rejected = []

    for carrier in carriers:
        blocked, block_reason = goods_type_blocks_carrier(goods_type, goods_value, carrier, goods_rule)
        if blocked:
            rejected.append({"Carrier": carrier, "Grund": block_reason})
            continue

        service_row = find_best_service_for_carrier(
            base_df=base_df,
            cols=cols,
            carrier=carrier,
            length=length,
            width=width,
            height=height,
            billable_weight=billable_weight,
            receiver=receiver,
        )

        if service_row is None:
            rejected.append({
                "Carrier": carrier,
                "Grund": "Kein passender Tarif für Gewicht, Maße oder Zielort gefunden.",
            })
            continue

        base_price = float(service_row["_price"])
        liability = normalize_text(service_row.get(cols.get("liability", ""), ""))
        insurance = calculate_insurance(carrier, goods_value, base_price, liability, insurance_df)

        total_with = insurance["total_with_insurance"]
        if total_with is None or pd.isna(total_with):
            rejected.append({
                "Carrier": carrier,
                "Grund": insurance["insurance_status"],
            })
            continue

        rows.append({
            "Carrier": carrier,
            "Versandart": normalize_text(service_row.get(cols["service"], "")),
            "Abrechnungsgewicht kg": round(billable_weight, 2),
            "Grundpreis": round(base_price, 2),
            "Versicherungsprämie": round(float(insurance["insurance_premium"]), 2),
            "Gesamtpreis ohne Versicherung": round(float(insurance["total_without_insurance"]), 2),
            "Gesamtpreis mit Versicherung": round(float(total_with), 2),
            "Versicherung": insurance["insurance_status"],
            "Versicherungshinweis": insurance["insurance_note"],
            "Tracking": normalize_text(service_row.get(cols.get("tracking", ""), "")),
            "Haftung": liability,
            "Laufzeit": normalize_text(service_row.get(cols.get("delivery_time", ""), "")),
        })

    results_df = pd.DataFrame(rows)
    if not results_df.empty:
        results_df = results_df.sort_values("Gesamtpreis mit Versicherung", ascending=True)

    return results_df, rejected, volumetric_weight, billable_weight


# ------------------------------------------------------------
# Ergebnisdarstellung
# ------------------------------------------------------------
def display_results(results_df: pd.DataFrame, rejected: List[Dict], volumetric_weight: float, billable_weight: float):
    st.subheader("📊 Ergebnis der Kostenberechnung")

    col1, col2 = st.columns(2)
    col1.metric("Abrechnungsgewicht", f"{billable_weight:.2f} kg")
    if not results_df.empty:
        best = results_df.sort_values("Gesamtpreis mit Versicherung", ascending=True).iloc[0]
        col2.metric("Günstigste passende Option", f"{best['Carrier']} – {best['Versandart']}", format_euro(best["Gesamtpreis mit Versicherung"]))
    else:
        col2.metric("Günstigste passende Option", "Keine passende Option")

    if results_df.empty:
        st.warning("EVA konnte keinen passenden Carrier finden. Bitte prüfe Gewicht, Maße, Warenwert oder Excel-Regeln.")
    else:
        display_df = results_df.copy()
        money_cols = ["Grundpreis", "Versicherungsprämie", "Gesamtpreis ohne Versicherung", "Gesamtpreis mit Versicherung"]
        for col in money_cols:
            display_df[col] = display_df[col].apply(format_euro)

        st.markdown("### Transparente Kostenaufschlüsselung je Carrier")
        st.dataframe(display_df, use_container_width=True)

        st.markdown("### Preisranking")
        price_ranking = results_df.sort_values("Gesamtpreis mit Versicherung", ascending=True)[
            ["Carrier", "Versandart", "Gesamtpreis mit Versicherung", "Versicherung", "Laufzeit"]
        ].copy()
        price_ranking["Gesamtpreis mit Versicherung"] = price_ranking["Gesamtpreis mit Versicherung"].apply(format_euro)
        st.dataframe(price_ranking, use_container_width=True)

        st.markdown("### Finale Empfehlung")
        best = results_df.sort_values("Gesamtpreis mit Versicherung", ascending=True).iloc[0]
        with st.chat_message("assistant", avatar="🤖"):
            st.write(
                f"Ich empfehle **{best['Carrier']} – {best['Versandart']}**. "
                f"Diese Option ist die günstigste passende Versandoption auf Basis der hinterlegten Excel-Daten. "
                f"Der berechnete Gesamtpreis mit Versicherung beträgt **{format_euro(best['Gesamtpreis mit Versicherung'])}**. "
                f"Die Empfehlung basiert auf den Tarifdaten, der Versicherungslogik, den Paketmaßen und den hinterlegten Carrier-Regeln."
            )

    if rejected:
        with st.expander("Abgelehnte oder nicht passende Carrier anzeigen"):
            st.dataframe(pd.DataFrame(rejected), use_container_width=True)


# ------------------------------------------------------------
# Sidebar: Upload-Bereich
# ------------------------------------------------------------
with st.sidebar:
    st.header("📁 Excel-Upload")
    uploaded_file = st.file_uploader(
        "Excel-Datei hochladen",
        type=["xlsx", "xls"],
        help="Lade die Datei mit Grundpreisen, Versicherungen, Zusatzservices und Regeln hoch.",
    )

    st.info(
        "Tipp: Deine Datei sollte mindestens ein Sheet 'Grundpreis' enthalten. "
        "Weitere empfohlene Sheets: Versicherung, Zusatzservices, Sonderregeln nach Warenart, Entscheidungsregeln."
    )


# ------------------------------------------------------------
# Hauptbereich: App-Logik
# ------------------------------------------------------------
if uploaded_file is None:
    st.warning("Bitte lade zuerst deine Excel-Datei hoch, damit EVA die Carrier-Daten lesen kann.")
    st.stop()

sheets = load_excel(uploaded_file)

if not sheets:
    st.stop()

with st.expander("Geladene Excel-Sheets anzeigen"):
    st.write(list(sheets.keys()))
    for sheet_name, df in sheets.items():
        st.write(f"**{sheet_name}** – {len(df)} Zeilen, {len(df.columns)} Spalten")

# Warenarten dynamisch aus Excel vorschlagen
rules_df = find_sheet(sheets, ["sonderregeln", "warenart"])
goods_options = ["Kleidung", "Bücher", "Dokumente", "Elektronik", "Smartphone", "Laptop", "Tablet", "Kamera", "Ersatzteile", "Industrieteile", "Schmuck", "Uhr", "Gefahrgut", "Sonstiges"]
if rules_df is not None:
    goods_col = find_column(rules_df, ["Warenart"])
    if goods_col:
        extracted = [normalize_text(x) for x in rules_df[goods_col].dropna().unique() if normalize_text(x)]
        goods_options = sorted(set(goods_options + extracted))

st.subheader("📝 Sendungsdaten eingeben")

with st.form("eva_input_form"):
    c1, c2, c3 = st.columns(3)
    with c1:
        length = st.number_input("Länge in cm *", min_value=0.0, value=40.0, step=1.0)
        raw_weight = st.number_input("Rohgewicht in kg *", min_value=0.0, value=4.0, step=0.1)
        sender = st.text_input("PLZ/Stadt Absender *", value="95028 Hof, Deutschland")
    with c2:
        width = st.number_input("Breite in cm *", min_value=0.0, value=30.0, step=1.0)
        goods_value = st.number_input("Warenwert in € *", min_value=0.0, value=1800.0, step=10.0)
        receiver = st.text_input("PLZ/Stadt Empfänger *", value="10115 Berlin, Deutschland")
    with c3:
        height = st.number_input("Höhe in cm *", min_value=0.0, value=10.0, step=1.0)
        goods_type = st.selectbox("Warenart *", options=goods_options, index=goods_options.index("Laptop") if "Laptop" in goods_options else 0)

    submitted = st.form_submit_button("EVA berechnen lassen")

if submitted:
    errors = validate_inputs(length, width, height, raw_weight, sender, receiver, goods_type, goods_value)
    if errors:
        st.error("Bitte korrigiere folgende Eingaben:")
        for e in errors:
            st.write("- " + e)
    else:
        results_df, rejected, volumetric_weight, billable_weight = calculate_all_carriers(
            sheets=sheets,
            length=length,
            width=width,
            height=height,
            raw_weight=raw_weight,
            sender=sender,
            receiver=receiver,
            goods_type=goods_type,
            goods_value=goods_value,
        )
        additional_services = select_additional_services(goods_type, goods_value, find_sheet(sheets, ["zusatz"]))

        st.session_state["latest_results"] = results_df
        st.session_state["latest_rejected"] = rejected
        st.session_state["latest_additional_services"] = additional_services
        st.session_state["latest_inputs"] = {
            "Länge cm": length,
            "Breite cm": width,
            "Höhe cm": height,
            "Rohgewicht kg": raw_weight,
            "Abrechnungsgewicht kg": billable_weight,
            "Absender": sender,
            "Empfänger": receiver,
            "Warenart": goods_type,
            "Warenwert €": goods_value,
        }

        display_results(results_df, rejected, volumetric_weight, billable_weight)

        st.markdown("### Empfohlene Zusatzleistungen")
        st.dataframe(pd.DataFrame(additional_services), use_container_width=True)


# ------------------------------------------------------------
# KI-Erklärung: Kontext für ChatGPT
# ------------------------------------------------------------
def build_eva_ai_context() -> str:
    """Baut einen sicheren Kontext aus der letzten EVA-Berechnung auf."""
    results_df = st.session_state.get("latest_results")
    rejected = st.session_state.get("latest_rejected", [])
    additional_services = st.session_state.get("latest_additional_services", [])
    inputs = st.session_state.get("latest_inputs", {})

    if results_df is None or not isinstance(results_df, pd.DataFrame) or results_df.empty:
        result_text = "Es liegt noch keine erfolgreiche Berechnung vor."
        best_text = "Noch keine Empfehlung vorhanden."
    else:
        best = results_df.sort_values("Gesamtpreis mit Versicherung", ascending=True).iloc[0]
        result_text = results_df.to_string(index=False)
        best_text = (
            f"Beste Option: {best['Carrier']} – {best['Versandart']} "
            f"mit Gesamtpreis inkl. Versicherung: {format_euro(best['Gesamtpreis mit Versicherung'])}."
        )

    rejected_text = pd.DataFrame(rejected).to_string(index=False) if rejected else "Keine abgelehnten Carrier."
    services_text = pd.DataFrame(additional_services).to_string(index=False) if additional_services else "Keine Zusatzleistungen gespeichert."

    return f"""
Du bist EVA, ein KI-gestützter Logistik-Assistent für einen Versandkostenvergleich.

Wichtige Regeln:
- Die Preisberechnung wurde bereits regelbasiert auf Basis der Excel-Daten durchgeführt.
- Erfinde keine neuen Preise, Tarife, Haftungsgrenzen oder Carrier-Regeln.
- Nutze nur die unten angegebenen Ergebnisse und erkläre sie verständlich.
- Wenn Daten fehlen, sage klar, dass EVA dafür zuerst eine Berechnung oder passende Excel-Daten braucht.
- Antworte auf Deutsch, kurz, professionell und verständlich für eine Projektpräsentation.

Eingaben der letzten Berechnung:
{inputs}

Berechnungsergebnisse:
{result_text}

{best_text}

Abgelehnte / nicht passende Carrier:
{rejected_text}

Empfohlene Zusatzleistungen:
{services_text}

Projektpositionierung:
EVA kombiniert eine regelbasierte Berechnungslogik mit generativer KI zur Erklärung der Ergebnisse und zur nutzerfreundlichen Kommunikation über die Chatbot-Oberfläche.
"""


def answer_with_chatgpt(user_question: str) -> str:
    """Beantwortet Nutzerfragen mit ChatGPT; nutzt Fallback, wenn kein API-Key vorhanden ist."""
    if openai_client is None:
        return (
            "ChatGPT ist lokal noch nicht aktiviert, weil kein OPENAI_API_KEY in Streamlit Secrets gefunden wurde. "
            "Die EVA-Berechnung funktioniert trotzdem. Für die KI-Erklärung musst du den API-Key in Streamlit Cloud unter Settings → Secrets eintragen."
        )

    try:
        response = openai_client.responses.create(
            model="gpt-4.1-mini",
            input=[
                {"role": "system", "content": build_eva_ai_context()},
                {"role": "user", "content": user_question},
            ],
            max_output_tokens=500,
        )
        return response.output_text
    except Exception as exc:
        return (
            "Die KI-Antwort konnte gerade nicht erzeugt werden. "
            f"Technischer Hinweis: {exc}"
        )


# ------------------------------------------------------------
# Chatbot-Interface
# ------------------------------------------------------------
st.divider()
st.subheader("💬 Chat mit EVA")

if "eva_messages" not in st.session_state:
    st.session_state.eva_messages = [
        {
            "role": "assistant",
            "content": "Hallo, ich bin EVA. Lade deine Excel-Datei hoch, gib die Sendungsdaten ein und ich vergleiche die Carrier für dich.",
        }
    ]

for msg in st.session_state.eva_messages:
    with st.chat_message(msg["role"], avatar="🤖" if msg["role"] == "assistant" else "👤"):
        st.write(msg["content"])

user_question = st.chat_input("Frage EVA etwas, z. B. 'Warum ist Versicherung nötig?' oder 'Welche Spalten braucht meine Excel?' ")

if user_question:
    st.session_state.eva_messages.append({"role": "user", "content": user_question})
    with st.chat_message("user", avatar="👤"):
        st.write(user_question)

    answer = answer_with_chatgpt(user_question)

    st.session_state.eva_messages.append({"role": "assistant", "content": answer})
    with st.chat_message("assistant", avatar="🤖"):
        st.write(answer)
