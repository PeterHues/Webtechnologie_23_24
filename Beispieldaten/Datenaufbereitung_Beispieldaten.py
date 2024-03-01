#%%
import pandas as pd
import streamlit as st
import plotly.express as px
import datetime
import numpy as np

# ---- Daten einlesen und aufbereiten ----
# ----- Schritt 1: Excel-Datei mit Beispieldaten einlesen
df = pd.read_excel(
    io="C:/Webtechnologie_23_24/Beispieldaten/Beispieldaten.xlsx",
    engine="openpyxl",
    sheet_name="Tabelle1",
    usecols="A:Q",
    na_values= ["",'#N/A', '#N/A N/A', '#NA', '-1.#IND', '-1.#QNAN', '-NaN', '-nan', '1.#IND', '1.#QNAN', '<NA>', 'N/A', 'NULL', 'NaN', 'None', 'n/a', 'nan', 'null'],
    keep_default_na=False,
    dtype={'Jahr': str}
)

# ----- Schritt 2: Spaltenueberschriften aendern
header_Beispieldaten = ["Produktgruppe1", "Produktgruppe1_Name", "Produktgruppe2",	"Produktgruppe2_Name",	"Produktgruppe3",	"Produktgruppe3_Name",
                        "Materialnummer",	"Materialname",	"Region_Kunde",	"Länderkürzel_Kunde",	"Land_Kunde",	"Kundennummer",	"Kundenname",
                        "Geschaeftsjahr",	"Absatz",	"Umsatz",	"Deckungsbeitrag"]

df.columns = header_Beispieldaten


# ----- Schritt 3: in den Kennzahlen Absatz, Umsatz, Deckungsbeitrag leere Zellen durch 0 ersetzen
df_vorbereitet = df
df_vorbereitet.loc[:, "Absatz":"Deckungsbeitrag"] = df_vorbereitet.loc[:, "Absatz":"Deckungsbeitrag"].fillna(0)
#%%

#Platzhalter Schritt 4: Im Orginalskript erfolgen hier die ABC-Analyse und die Klassifizierung des Deckungsbeitrages.
#Die Ergebnisse werden in das Dataframe result_df geschrieben. Da dieser Schritt hier ausbleibt, wird lediglich df_vorbereitet in result_df geschrieben.

result_df = df_vorbereitet

# ---- Formattierungen am Ende vornehmen ----
#Materialnummern als string und achtstellig formattieren. Somit kann man sie später leichter suchen, falls sie achtstellig formattiert sind
result_df.loc[:,"Materialnummer"] = result_df["Materialnummer"].apply(lambda x: str(x).zfill(8))

#Produkthierarchie 2 als string und zweistellig formattieren. Sieht schöner aus
result_df.loc[:,"Produktgruppe1"] = result_df["Produktgruppe1"].apply(lambda x: str(x).zfill(2))
result_df.loc[:,"Produktgruppe2"] = result_df["Produktgruppe2"].apply(lambda x: str(x).zfill(2))
# %%
result_df.to_feather("C:/Webtechnologie_23_24/Beispieldaten//Input_Beispieldaten.feather")
print("Beispieldaten erstellt")
# %%
