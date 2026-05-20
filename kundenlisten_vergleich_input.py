import io
import re
import csv
from collections import OrderedDict

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Kundenlisten Vergleich", layout="wide")


STANDARD_SPALTEN = ["SAP", "CSB", "Kundenname", "Straße", "Postleitzahl", "Ort"]

SPALTEN_KANDIDATEN = {
    "SAP": [
        "sap",
        "sap nr",
        "sap-nr",
        "sap nummer",
        "sap-nr debitoren",
        "sap nr debitoren",
        "sap debitoren",
        "sap-nr. debitoren",
    ],
    "CSB": [
        "csb",
        "kd nr",
        "kd-nr",
        "kd.-nr",
        "kundennummer",
        "kunden nr",
        "kunden-nr",
        "kdnr",
    ],
    "Kundenname": [
        "kundenname",
        "name",
        "kunde",
        "kunden name",
    ],
    "Straße": [
        "straße",
        "strasse",
        "strasse",
        "str",
        "str.",
        "anschrift",
    ],
    "Postleitzahl": [
        "postleitzahl",
        "plz",
    ],
    "Ort": [
        "ort",
        "stadt",
    ],
}


def text_wert(wert):
    if pd.isna(wert):
        return ""
    text = str(wert).replace("\u00a0", " ").strip()
    text = re.sub(r"\s+", " ", text)
    return text


def nummer_wert(wert):
    text = text_wert(wert)
    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]
    text = re.sub(r"\s+", "", text)
    return text


def vergleich_text(wert):
    return text_wert(wert).upper()


def kopf_normalisieren(wert):
    text = text_wert(wert).lower()
    text = text.replace("\n", " ")
    text = text.replace("ß", "ss")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def kandidaten_normalisieren(kandidaten):
    return {kopf_normalisieren(x) for x in kandidaten}


def finde_spalte(df, ziel):
    vorhandene = {kopf_normalisieren(spalte): spalte for spalte in df.columns}
    kandidaten = kandidaten_normalisieren(SPALTEN_KANDIDATEN[ziel])

    for kandidat in kandidaten:
        if kandidat in vorhandene:
            return vorhandene[kandidat]

    # etwas toleranter für SAP-Debitoren und Kd.-Nr.
    for norm, original in vorhandene.items():
        if ziel == "SAP" and "sap" in norm:
            return original
        if ziel == "CSB" and (norm == "csb" or "kd nr" in norm or "kunden nr" in norm):
            return original
        if ziel == "Straße" and ("strasse" in norm or "str" == norm):
            return original
        if ziel == "Postleitzahl" and ("plz" in norm or "postleitzahl" in norm):
            return original
        if ziel == "Kundenname" and norm in {"name", "kundenname", "kunde"}:
            return original
        if ziel == "Ort" and norm == "ort":
            return original

    return None


def lese_csv(datei):
    roh = datei.getvalue()

    letzte_fehler = None
    for encoding in ["utf-8-sig", "utf-8", "cp1252", "latin1"]:
        try:
            text = roh.decode(encoding)
            probe = text[:8192]
            try:
                dialect = csv.Sniffer().sniff(probe, delimiters=";,|\t,")
                trenner = dialect.delimiter
            except Exception:
                trenner = ";"
            return pd.read_csv(io.StringIO(text), sep=trenner, dtype=str)
        except Exception as fehler:
            letzte_fehler = fehler

    raise ValueError(f"CSV konnte nicht gelesen werden: {letzte_fehler}")


def lese_excel(datei):
    excel = pd.ExcelFile(datei)

    bevorzugte_blaetter = []
    for blatt in excel.sheet_names:
        if "liste" in blatt.lower():
            bevorzugte_blaetter.append(blatt)
    bevorzugte_blaetter.extend([blatt for blatt in excel.sheet_names if blatt not in bevorzugte_blaetter])

    beste_df = None
    beste_blatt = None
    beste_treffer = -1

    for blatt in bevorzugte_blaetter:
        df = pd.read_excel(excel, sheet_name=blatt, dtype=str)
        treffer = sum(1 for ziel in STANDARD_SPALTEN if finde_spalte(df, ziel) is not None)
        if treffer > beste_treffer:
            beste_treffer = treffer
            beste_df = df
            beste_blatt = blatt

    if beste_df is None or beste_treffer < 4:
        raise ValueError("In der Excel-Datei wurden nicht genug passende Spalten gefunden.")

    st.caption(f"Excel-Blatt verwendet: {beste_blatt}")
    return beste_df


def lese_datei(datei):
    name = datei.name.lower()
    if name.endswith(".csv"):
        return lese_csv(datei)
    if name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls"):
        return lese_excel(datei)
    raise ValueError("Bitte CSV oder Excel hochladen.")


def baue_standard_df(df, quelle):
    zuordnung = OrderedDict()
    fehlend = []

    for ziel in STANDARD_SPALTEN:
        spalte = finde_spalte(df, ziel)
        if spalte is None:
            fehlend.append(ziel)
        else:
            zuordnung[ziel] = spalte

    if fehlend:
        raise ValueError(
            f"{quelle}: Diese Spalten wurden nicht gefunden: {', '.join(fehlend)}"
        )

    out = pd.DataFrame()
    out["SAP"] = df[zuordnung["SAP"]].map(nummer_wert)
    out["CSB"] = df[zuordnung["CSB"]].map(nummer_wert)
    out["Kundenname"] = df[zuordnung["Kundenname"]].map(text_wert)
    out["Straße"] = df[zuordnung["Straße"]].map(text_wert)
    out["Postleitzahl"] = df[zuordnung["Postleitzahl"]].map(nummer_wert)
    out["Ort"] = df[zuordnung["Ort"]].map(text_wert)

    ohne_sap = out[out["SAP"] == ""].copy()
    ohne_sap.insert(0, "Quelle", quelle)

    out = out[out["SAP"] != ""].copy()
    out = out.reset_index(drop=True)

    return out, ohne_sap, zuordnung


def vergleichsschluessel(df):
    temp = df.copy()
    temp["_CSB"] = temp["CSB"].map(nummer_wert)
    temp["_Kundenname"] = temp["Kundenname"].map(vergleich_text)
    temp["_Straße"] = temp["Straße"].map(vergleich_text)
    temp["_Postleitzahl"] = temp["Postleitzahl"].map(nummer_wert)
    temp["_Ort"] = temp["Ort"].map(vergleich_text)
    return temp


def mache_vergleich(kisoft, original):
    k_saps = set(kisoft["SAP"])
    o_saps = set(original["SAP"])

    nur_k_saps = sorted(k_saps - o_saps, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))
    nur_o_saps = sorted(o_saps - k_saps, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))
    gemeinsame_saps = sorted(k_saps & o_saps, key=lambda x: (0, int(x)) if str(x).isdigit() else (1, str(x)))

    nur_k = kisoft[kisoft["SAP"].isin(nur_k_saps)].sort_values("SAP").copy()
    nur_o = original[original["SAP"].isin(nur_o_saps)].sort_values("SAP").copy()

    k_erster = kisoft.drop_duplicates("SAP", keep="first").set_index("SAP")
    o_erster = original.drop_duplicates("SAP", keep="first").set_index("SAP")

    k_v = vergleichsschluessel(k_erster.reset_index()).set_index("SAP")
    o_v = vergleichsschluessel(o_erster.reset_index()).set_index("SAP")

    abweichungen = []
    gleiche = 0

    prueffelder = ["CSB", "Kundenname", "Straße", "Postleitzahl", "Ort"]

    for sap in gemeinsame_saps:
        abweichende_felder = []
        for feld in prueffelder:
            vergleichsspalte = "_" + feld
            if str(k_v.at[sap, vergleichsspalte]) != str(o_v.at[sap, vergleichsspalte]):
                abweichende_felder.append(feld)

        if not abweichende_felder:
            gleiche += 1
            continue

        zeile = {
            "SAP": sap,
            "Abweichende Felder": ", ".join(abweichende_felder),
        }
        for feld in prueffelder:
            zeile[f"Kisoft {feld}"] = k_erster.at[sap, feld]
            zeile[f"Original {feld}"] = o_erster.at[sap, feld]
        abweichungen.append(zeile)

    abweichungen_df = pd.DataFrame(abweichungen)

    duplikate_sap = pd.concat(
        [
            kisoft[kisoft.duplicated("SAP", keep=False)].assign(Quelle="Kisoft"),
            original[original.duplicated("SAP", keep=False)].assign(Quelle="Original"),
        ],
        ignore_index=True,
    )
    if not duplikate_sap.empty:
        duplikate_sap = duplikate_sap[["Quelle"] + STANDARD_SPALTEN].sort_values(["SAP", "Quelle"])

    duplikate_csb = pd.concat(
        [
            kisoft[(kisoft["CSB"] != "") & kisoft.duplicated("CSB", keep=False)].assign(Quelle="Kisoft"),
            original[(original["CSB"] != "") & original.duplicated("CSB", keep=False)].assign(Quelle="Original"),
        ],
        ignore_index=True,
    )
    if not duplikate_csb.empty:
        duplikate_csb = duplikate_csb[["Quelle"] + STANDARD_SPALTEN].sort_values(["CSB", "Quelle", "SAP"])

    kennzahlen = {
        "Kisoft Zeilen mit SAP": len(kisoft),
        "Original Zeilen mit SAP": len(original),
        "Gemeinsame SAP-Nummern": len(gemeinsame_saps),
        "Komplett gleiche gemeinsame SAP-Nummern": gleiche,
        "Gemeinsame SAP mit abweichenden Daten": len(abweichungen_df),
        "Nur in Kisoft - eindeutige SAP": len(nur_k_saps),
        "Nur in Kisoft - Zeilen": len(nur_k),
        "Nur im Original - eindeutige SAP": len(nur_o_saps),
        "Nur im Original - Zeilen": len(nur_o),
        "Doppelte SAP-Zeilen": len(duplikate_sap),
        "Doppelte CSB-Zeilen": len(duplikate_csb),
    }

    return kennzahlen, nur_k, nur_o, abweichungen_df, duplikate_sap, duplikate_csb


def excel_download(kennzahlen, nur_k, nur_o, abweichungen, duplikate_sap, duplikate_csb, ohne_sap, kisoft, original):
    ausgabe = io.BytesIO()

    with pd.ExcelWriter(ausgabe) as writer:
        pd.DataFrame({"Prüfung": list(kennzahlen.keys()), "Wert": list(kennzahlen.values())}).to_excel(
            writer, sheet_name="Übersicht", index=False
        )
        nur_k.to_excel(writer, sheet_name="Nur in Kisoft", index=False)
        nur_o.to_excel(writer, sheet_name="Nur im Original", index=False)
        abweichungen.to_excel(writer, sheet_name="Abweichungen", index=False)
        duplikate_sap.to_excel(writer, sheet_name="Duplikate SAP", index=False)
        duplikate_csb.to_excel(writer, sheet_name="Duplikate CSB", index=False)
        ohne_sap.to_excel(writer, sheet_name="Zeilen ohne SAP", index=False)
        kisoft.to_excel(writer, sheet_name="Kisoft bereinigt", index=False)
        original.to_excel(writer, sheet_name="Original bereinigt", index=False)

    ausgabe.seek(0)
    return ausgabe.getvalue()


st.title("Kundenlisten Vergleich")
st.write(
    "Lade die Kisoft-Kundenliste und die Original-Kundenliste hoch. "
    "Der Vergleich läuft über die SAP-Nummer."
)

links, rechts = st.columns(2)
with links:
    kisoft_datei = st.file_uploader("Kisoft Datei", type=["csv", "xlsx", "xlsm", "xls"], key="kisoft")
with rechts:
    original_datei = st.file_uploader("Original Datei", type=["csv", "xlsx", "xlsm", "xls"], key="original")

if kisoft_datei and original_datei:
    try:
        kisoft_roh = lese_datei(kisoft_datei)
        original_roh = lese_datei(original_datei)

        kisoft, kisoft_ohne_sap, kisoft_mapping = baue_standard_df(kisoft_roh, "Kisoft")
        original, original_ohne_sap, original_mapping = baue_standard_df(original_roh, "Original")

        kennzahlen, nur_k, nur_o, abweichungen, duplikate_sap, duplikate_csb = mache_vergleich(kisoft, original)
        ohne_sap = pd.concat([kisoft_ohne_sap, original_ohne_sap], ignore_index=True)

        st.subheader("Ergebnis")
        metrik_spalten = st.columns(5)
        metrik_spalten[0].metric("Nur in Kisoft", kennzahlen["Nur in Kisoft - eindeutige SAP"])
        metrik_spalten[1].metric("Nur im Original", kennzahlen["Nur im Original - eindeutige SAP"])
        metrik_spalten[2].metric("Abweichungen", kennzahlen["Gemeinsame SAP mit abweichenden Daten"])
        metrik_spalten[3].metric("Gemeinsam", kennzahlen["Gemeinsame SAP-Nummern"])
        metrik_spalten[4].metric("Gleich", kennzahlen["Komplett gleiche gemeinsame SAP-Nummern"])

        st.caption("Erkannte Spalten Kisoft: " + ", ".join([f"{k} = {v}" for k, v in kisoft_mapping.items()]))
        st.caption("Erkannte Spalten Original: " + ", ".join([f"{k} = {v}" for k, v in original_mapping.items()]))

        excel = excel_download(
            kennzahlen,
            nur_k,
            nur_o,
            abweichungen,
            duplikate_sap,
            duplikate_csb,
            ohne_sap,
            kisoft,
            original,
        )

        st.download_button(
            "Excel-Auswertung herunterladen",
            data=excel,
            file_name="kundenlisten_unterschiede.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
            [
                "Abweichungen",
                "Nur in Kisoft",
                "Nur im Original",
                "Duplikate SAP",
                "Duplikate CSB",
                "Zeilen ohne SAP",
            ]
        )

        with tab1:
            st.dataframe(abweichungen, use_container_width=True, hide_index=True)
        with tab2:
            st.dataframe(nur_k, use_container_width=True, hide_index=True)
        with tab3:
            st.dataframe(nur_o, use_container_width=True, hide_index=True)
        with tab4:
            st.dataframe(duplikate_sap, use_container_width=True, hide_index=True)
        with tab5:
            st.dataframe(duplikate_csb, use_container_width=True, hide_index=True)
        with tab6:
            st.dataframe(ohne_sap, use_container_width=True, hide_index=True)

    except Exception as fehler:
        st.error(str(fehler))
else:
    st.info("Bitte beide Dateien hochladen.")
