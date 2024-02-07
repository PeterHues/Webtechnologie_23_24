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
#Liste mit den eindeutigen Geschäftsjahren erstellen. Somit kann für jedes Geschäftsjahr die ABCD und die HMLN-Klassifizierung vorgenommen werden 

Jahreszahlen = sorted(df_vorbereitet["Geschaeftsjahr"].unique(), reverse=False)
#%%
# ---- ABC-Kategorien ermitteln ----
def filter_dataframe(df, value1, value2, value3):
    
    list_dfs_function = []

    for Jahr in Jahreszahlen:

        # ---- Dataframe für ABC-Analyse vorbereiten ----    
        #Das Dataframe zunaechst nach den uebergebenen Produktgruppen 1-3 filtern. Somit werden nur vergleichbare Materialien aus einer Produktgruppe in der ABC-Analyse betrachtet
        df_gefiltert_PRDH3 = df[(df["Produktgruppe1"] == value1) & (df["Produktgruppe2"] == value2) & (df["Produktgruppe3"] == value3)]

           
        #Schritt 1: nach  Produktgruppe und den Materialnummern gruppieren und die Spalte "Absatz" addieren.
        #Die Produktgruppe ist nur noch dabei, damit man zur Kontrolle die Produktgruppen filtern kann
        df_filtered_PRDH3_grouped = df_gefiltert_PRDH3.query('Geschaeftsjahr == @Jahr').groupby(["Produktgruppe1", "Produktgruppe2", "Produktgruppe3", "Materialnummer", "Materialname", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()  

        # ---- ABC-Analyse durchfuehren ----
        #Sortierung der Spalte 'Absatz' absteigend
        df_filtered_PRDH3_grouped = df_filtered_PRDH3_grouped.sort_values(by='Absatz', ascending=False)
        
        # Schritt 1: Berechnen der kumulierten Summe
        df_filtered_PRDH3_grouped['cumulated_sum'] = df_filtered_PRDH3_grouped['Absatz'].cumsum()

        # Schritt 2: Berechnen der Gesamtsumme
        gesamtsumme = df_filtered_PRDH3_grouped['Absatz'].sum()

        # Schritt 3: Berechnen des kumulierten Prozentsatzes
        df_filtered_PRDH3_grouped['cumulated_percentages'] = (df_filtered_PRDH3_grouped['cumulated_sum'] / gesamtsumme) * 100

        # Schritt 4: Kategorisierung in A, B, C und D
        df_filtered_PRDH3_grouped["ABCD"] = ""
        j = 0
        while(j < len(df_filtered_PRDH3_grouped)):
            if(df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("Absatz")] == 0):
                df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("ABCD")] = "D" 
            elif(df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("cumulated_percentages")] > 90):
                df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("ABCD")] = "C"
            elif(df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("cumulated_percentages")] > 80):
                df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("ABCD")] = "B"
            elif(df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("cumulated_percentages")] > 0):
                df_filtered_PRDH3_grouped.iloc[j, df_filtered_PRDH3_grouped.columns.get_loc("ABCD")] = "A"   
            j += 1
        
        
        list_dfs_function.append(df_filtered_PRDH3_grouped)

    # Schritt 6: Das Rueckgabe-Dataframe erstellen, in dem nur die Materialnummern und die ABCD-Kategorien enthalten sind
    #Das Rueckgabe-Dataframe wird zunaechst in eine Liste von Dataframes (list_of_dfs) geschrieben und anschließend im Dataframe result_df_ABCD entpackt
    #Die Produktgruppen 1-3 muessen mit uebergeben werden, da in der Realitaet manche Materialnummern mehreren Produktgruppen zugeordnet sind. Somit kaeme das Material, wenn man es nur ueber die Materialnummer joint,
    #mehrmals in seinen jeweiligen Produktgruppen vor
    rueckgabe_df = pd.concat(list_dfs_function, ignore_index=True)
    rueckgabe_df = rueckgabe_df[['Produktgruppe1', 'Produktgruppe2', 'Produktgruppe3', "Materialnummer", "Geschaeftsjahr", "ABCD"]]

    return rueckgabe_df

list_of_dfs = []

#Liste mit allen eindeutigen Produktgruppen 1-3 erstellen
#fuer jede eindeutige Kombination von Produktgruppe 1-3 werden in der Schleife fuer jedes Jahr die ABCD-Kategorien fuer die jeweiligen Materialien ermittelt
list_all_PRDHs = df_vorbereitet.groupby(['Produktgruppe1', 'Produktgruppe2', 'Produktgruppe3']).first().reset_index()
list_all_PRDHs = list_all_PRDHs.loc[:, ['Produktgruppe1', 'Produktgruppe2', 'Produktgruppe3']]

i = 0
while(i < len(list_all_PRDHs)):
    zwischen_df = filter_dataframe(df_vorbereitet, value1=list_all_PRDHs.iloc[i, list_all_PRDHs.columns.get_loc("Produktgruppe1")], value2=list_all_PRDHs.iloc[i, list_all_PRDHs.columns.get_loc("Produktgruppe2")], value3=list_all_PRDHs.iloc[i, list_all_PRDHs.columns.get_loc("Produktgruppe3")])
    list_of_dfs.append(zwischen_df)
    i += 1

#Result_df_ABCD enthaelt alle Materialien mit ihren ABCD-Kategorien
result_df_ABCD = pd.concat(list_of_dfs, ignore_index=True)


#%%
# ---- HMLN ----
# ---- HMLN-Kategorien ermitteln ----
#Schritt 1: Deckungsbeitrag (DB) pro KG ermitteln
df_HMLN = df_vorbereitet.groupby(["Materialnummer", "Geschaeftsjahr"], as_index=False)[["Absatz", "Deckungsbeitrag"]].sum()

df_HMLN['DB_per_KG'] = (df_HMLN['Deckungsbeitrag'] / df_HMLN['Absatz'])

#Schritt 2: Spalte fuer DB-Kategorie erstellen
df_HMLN["HMLN"] = ""

#Schritt 3: den verschiedenen DB pro Kilogramm die Kategorie zuweisen. 
#Anmerkung: falls eine Zahl keine Verkaufsmengen hatte (Absatz == 0), aber dennoch einen DB erzielt hat, ergibt die Berechnung des DB pro KG "inf2". 
#Der DB pro KG ergibt "nan", wenn sowohl Absatz als auch Deckungsbeitrag 0 sind
#diese Deckungsbeitraege fuer "inf" und "nan" werden alle der Kategorie N zugeordnet. 

i = 0
while(i < len(df_HMLN)):
    if(np.isinf(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")])):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "N" 
    elif(np.isnan(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")])):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "N" 
    elif(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")] <= 0.0):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "N" 
    elif(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")] <= 0.15):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "L"
    elif(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")] <= 0.3):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "M"
    elif(df_HMLN.iloc[i, df_HMLN.columns.get_loc("DB_per_KG")] > 0.3):
        df_HMLN.iloc[i, df_HMLN.columns.get_loc("HMLN")] = "H"   
    i += 1

#enthaelt die HMLN-Kategorie je Materialnummer und Geschaeftsjahr. Die einzelnen Produktgruppen werden nicht beruecksichtigt, da der Deckungsbeitrag absolut fuer das 
#Material ermittelt wird und nicht in Abhaengigkeit der Produktgruppe
result_df_HMLN = df_HMLN.loc[:, ["Materialnummer", "Geschaeftsjahr", "HMLN"]]

#%%
#result_df erstellen und die wichtigsten Spalten behalten - vor allem irrelevante Kennzahlen werden entfernt
result_df = pd.merge(df_vorbereitet, pd.merge(result_df_ABCD, result_df_HMLN, on=['Materialnummer', 'Geschaeftsjahr'], how='left'), on=['Produktgruppe1', 'Produktgruppe2', 'Produktgruppe3', 'Materialnummer', 'Geschaeftsjahr'], how='left')

result_df = result_df[["Produktgruppe1", "Produktgruppe1_Name", "Produktgruppe2",	"Produktgruppe2_Name",	"Produktgruppe3",	"Produktgruppe3_Name",
                        "Materialnummer",	"Materialname",	"Region_Kunde",	"Länderkürzel_Kunde",	"Land_Kunde",	"Kundennummer",	"Kundenname",
                        "Geschaeftsjahr",	"Absatz",	"Umsatz",	"Deckungsbeitrag", "ABCD", "HMLN"]]

# %%

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
