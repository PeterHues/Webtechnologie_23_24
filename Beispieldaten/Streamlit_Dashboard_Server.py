#%%
import pandas as pd
import streamlit as st
import plotly.express as px
import datetime
import numpy as np
import xlsxwriter
import io
import hmac

#Passwort abfragen
def check_password():
    """Returns `True` if the user had the correct password."""

    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if hmac.compare_digest(st.session_state["password"], st.secrets["password"]):
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Don't store the password.
        else:
            st.session_state["password_correct"] = False

    # Return True if the password is validated.
    if st.session_state.get("password_correct", False):
        return True

    # Show input for password.
    st.text_input(
        "Password", type="password", on_change=password_entered, key="password"
    )
    if "password_correct" in st.session_state:
        st.error("😕 Password incorrect")
    return False


if not check_password():
    st.stop()  # Do not continue if check_password is not True.



# ---- Input-Data einlesen ----
input_data = pd.read_feather("Beispieldaten/Input_Beispieldaten.feather")

working_data = input_data

#%%
# ---- Streamlit Seiteneinstellungen anpassen
st.set_page_config(page_title="Absatz Dashboard", #Name des Browser-Tabs
                   page_icon=":bar_chart:", #Symbol des Browsertab
                   layout="wide" #gesamten Screen benutzen
)


st.title(":bar_chart: Absatz Dashboard")
st.markdown("---")




#%%
# ---- SIDEBAR erstellen ----
#Die Sidebar- und Eingabefilter werden als erstes erstellt, damit mit diesen Filter das df_selection erstellt werden kann.
st.sidebar.header("Bitte Hier Filtern:")

#Auswahlfilter für Produktgruppe 1 erstellen
#Elemente sind die Produktgruppen in der Spalte Produktgruppe1
pg1 = st.sidebar.multiselect(
    "Produktgruppe 1 wählen",
    options=sorted(working_data["Produktgruppe1"].unique(), reverse=False),
    default=sorted(working_data["Produktgruppe1"].unique(), reverse=False)
)

#Auswahlfilter für Produktgruppe 2 erstellen
#Elemente sind die Produktgruppen in der Spalte Produktgruppe2
pg2 = st.sidebar.multiselect(
    "Produktgruppe 2 wählen",
    options=sorted(working_data["Produktgruppe2"].unique(), reverse=False),
    default=sorted(working_data["Produktgruppe2"].unique(), reverse=False)
)

#Auswahlfilter für Geschaeftsjahr erstellen
#Elemente sind die Geschaeftsjahre in der Spalte Geschaeftsjahre
jahr_gefiltert = st.sidebar.multiselect(
    "Bitte Jahr wählen:",
    options=sorted(working_data["Geschaeftsjahr"].unique(), reverse=True),
    default=sorted(working_data["Geschaeftsjahr"].unique(), reverse=True)[:3]
)

#Auswahlfilter für Region erstellen
#Elemente sind die Regionen in der Spalte Region
region = st.sidebar.multiselect(
    "Bitte Region wählen:",
    options=sorted(working_data["Region_Kunde"].unique(), reverse=False),
    default=sorted(working_data["Region_Kunde"].unique(), reverse=False)
)

#das Dataframe entsprechend filtern
df_selection =working_data.query(
    "Produktgruppe1 == @pg1 & Produktgruppe2 == @pg2 & Geschaeftsjahr == @jahr_gefiltert & Region_Kunde == @region"
)

#Spalten Container für subheader
st.subheader("Materialübersicht")

#Spalten-Container einfuegen
left_column, right_column = st.columns(2)

with left_column:
    filter_material = st.text_input('Bitte Materialnummer eingeben:')

with right_column:
    filter_country = st.text_input('Bitte vollständigen Landesnamen eingeben:')

#Filter nur berücksichtigen, wenn in die Textfelder etwas eingegeben wurde. Sonst enthalten die Filter als Wert None
#und ein leeres Dataframe wäre die Folge
if filter_material:
    df_selection = df_selection[df_selection['Materialnummer'].str.contains(filter_material)]

if filter_country:
    df_selection = df_selection[df_selection['Land_Kunde'] == filter_country]

#Dataframe als Tabelle in Streamlit darstellen und die Spalten Geschaeftsjahr, Kundennummer, Absatz, Umsatz und Deckungsbeitrag als Integer darstellen
st.dataframe(
    df_selection,
    column_config={
        "Geschaeftsjahr": st.column_config.NumberColumn(
            format="%i"
        ),
        "Kundennummer": st.column_config.NumberColumn(
        format="%i"
        ),
        "Absatz": st.column_config.NumberColumn(
        format="%i"
        ),
        "Umsatz": st.column_config.NumberColumn(
        format="%i"
        ),
        "Deckungsbeitrag": st.column_config.NumberColumn(
        format="%i"
        ),
    },
    hide_index=True,
    use_container_width=True
)

#%%

# ---- Dpwnloadbutton erstellen ----
col1, col2 = st.columns([0.9,0.1])
#Puffer in Arbeitsspeicher für Excel-Datei vorhalten
buffer = io.BytesIO()

#Dataframe Filterkriterien erstellen - enthält die ausgewählten Filterkriterien, welche in der Excel-Tabelle in den Spalten A und B eingefügt werden
filterkriterien = pd.DataFrame(data={
    "Filter": ["Produktgruppe 1", "Produktgruppe 2", "Geschäftsjahr", "Region", "Materialnummer", "Land Kunde"],
    "Gesetzte Werte": [str(pg1), 
                       str(sorted(df_selection["Produktgruppe2"].unique(), reverse=False)), 
                       str(jahr_gefiltert), 
                       str(region),
                       str(filter_material),
                       str(filter_country)]
})

#Falls ein Filterkriterium nicht ausgewählt wurde, bspw. keine Materialnummer speziell gesucht wird, soll "-" in der Zelle stehen
i = 0
while(i < len(filterkriterien)):
    if(filterkriterien.iloc[i, 1]== ""):
        filterkriterien.iloc[i, 1] = "-"
    i = i+1

#Einen Pandas Excel-Writer mit XlsxWriter als Engine erstellen
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:

    #Dataframes in XlsxWriter-Objekte konvertieren
    filterkriterien.to_excel(writer, sheet_name='Tabelle1', startrow=0, index=False)
    df_selection.to_excel(writer, sheet_name='Tabelle1', startrow=len(filterkriterien)+3, index=False)

    #Aus den xlsxwriter-Objekten das workbook und Tabellenblatt "Tabelle1" extrahieren
    workbook = writer.book
    worksheet = writer.sheets["Tabelle1"]

    #Länge und Breite des df_selection in die Variablen max_row und max_col schreiben
    #werden im add_table()-Befehl verwendet
    (max_row, max_col) = df_selection.shape

    #Liste mit Überschriften erstellen, welche in add_table() verwendet werden können
    column_settings = [{"header": column} for column in df_selection.columns]

    #das Dataframe df_selection als Tabelle darstellen
    #Aufbau des add_table()-Befehls: Startzeile, Startspalte, Endzeile, Endspalte
    worksheet.add_table(len(filterkriterien)+3,0, max_row+len(filterkriterien)+3, max_col-1, {"columns": column_settings,
                                                                                              "style": 'Table Style Light 1'})
    
    #das Format mit den Rahmenlinien erstellen
    rahmenlinien = workbook.add_format({'border': 1})

    #das Format iterativ mit den Werten aus dem Dataframe filterkriterien in das Excel-Tabellenblatt schreiben
    i = 0
    while i < len(filterkriterien):
        worksheet.write(i+1, 0, filterkriterien.iloc[i,0], rahmenlinien)
        worksheet.write(i+1, 1, filterkriterien.iloc[i,1], rahmenlinien)
        i = i+1

    #alle Gridlines ausschalten
    worksheet.hide_gridlines(2)

    #automatische Spaltenbreiten anpassen
    worksheet.autofit()




#Downloadbutton einfügen, welcher die Excel-Datei aus dem RAM-Puffer verwendet
with col2:
    st.download_button(
        label=":inbox_tray: Download Excel-Arbeitsmappe",
        data=buffer,
        file_name="Absatzdaten.xlsx",
        mime="application/vnd.ms-excel",
        use_container_width=True
    )

st.markdown("---")

#Spaltencontainer für die Diagramme der jeweiligen Produktgruppen einfügen
col_Plastic_AS1, col_Plastic_AS2, col_Plastic_AS3 = st.columns(3)

col_Plastic_ABS_Plastic1, col_Plastic_ABS_Plastic2, col_Plastic_ABS_Plastic3 = st.columns(3)

col_PET_PO1, col_PET_PO2, col_PET_PO3 = st.columns(3)

col_PET_HP1, col_PET_HP2, col_PET_HP3 = st.columns(3)

col_PET_Special1, col_PET_Special2, col_PET_Special3 = st.columns(3)

#Farben für Säulendiagramm manuell festlegen
color_absatz = '#9ecae1'  # Hellblau
color_umsatz = '#fdae61'  # Pfirsich


#Funktion fuer Balkendiagramme erstellen
def Saeulendiagramme_erstellen(df):

    Saeulendiagramm = px.bar(df, 
        x='Geschaeftsjahr', 
        y=['Absatz', 'Umsatz'], 
        color_discrete_sequence=[color_absatz, color_umsatz],
        labels={'value': 'Menge', 'variable' : 'Legende'},
        barmode='group',
        template="plotly_white") 
    
    Saeulendiagramm.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=(dict(showgrid=False)),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.1,
            xanchor="center",
            x=0.5
        )   
    )

    return Saeulendiagramm


#Funktion fuer Liniendiagramme erstellen
def Liniendiagramme_erstellen(df):
    
    liniendiagramm = px.line(df, x='Geschaeftsjahr', y='Absatz', color='Produktgruppe3', symbol="Produktgruppe3")
    return liniendiagramm
        

#Funktion fuer Tortendiagramme erstellen
def Tortendiagramm_erstellen(df):
    
    tortendiagramm = px.pie(labels=df["Region_Kunde"], 
                                    values=df["Absatz"], 
                                    names=df["Region_Kunde"],
                )
    
    return tortendiagramm
    
    
#Diagramme Plastic AS
with col_Plastic_AS1:
    balkendiagramm_Plastic_AS = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    # Saeulendiagramm erstellen
    fig_Plastic_AS = Saeulendiagramme_erstellen(balkendiagramm_Plastic_AS)

    st.subheader("Absatz- und Umsatzentwicklung der Produktgruppe 2 Plastic AS")
    st.plotly_chart(fig_Plastic_AS, use_container_width=True)

with col_Plastic_AS2:
    liniendiagramm_Plastic_AS = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_Plastic_AS = Liniendiagramme_erstellen(liniendiagramm_Plastic_AS)
    
    st.subheader("Absatzentwicklung der Produktgruppen 3 von Plastic AS")
    st.plotly_chart(fig_line_Plastic_AS, use_container_width=True)

with col_Plastic_AS3:
    pie_chart_regionen_Plastic_AS = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Region_Kunde"], as_index=False)[["Absatz"]].sum()

    fig_pie_Plastic_AS = Tortendiagramm_erstellen(pie_chart_regionen_Plastic_AS)


    st.subheader("Absatzverteilung der Produktgruppe 2 Plastic AS pro Region")
    st.plotly_chart(fig_pie_Plastic_AS, use_container_width=True)




#Diagramme Plastic ABS_Plastic
with col_Plastic_ABS_Plastic1:
    balkendiagramm_Plastic_ABS_Plastic = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    fig_Plastic_ABS_Plastic = Saeulendiagramme_erstellen(balkendiagramm_Plastic_ABS_Plastic)

    st.subheader("Absatz- und Umsatzentwicklung der Produktgruppe 2 Plastic ABS_Plastic")
    st.plotly_chart(fig_Plastic_ABS_Plastic, use_container_width=True)


with col_Plastic_ABS_Plastic2:
    liniendiagramm_Plastic_ABS_Plastic = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_Plastic_ABS_Plastic = Liniendiagramme_erstellen(liniendiagramm_Plastic_ABS_Plastic)
    
    st.subheader("Absatzentwicklung der Produktgruppen 3 von Plastic ABS_Plastic")
    st.plotly_chart(fig_line_Plastic_ABS_Plastic, use_container_width=True)

with col_Plastic_ABS_Plastic3:
    pie_chart_regionen_Plastic_ABS_Plastic = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Region_Kunde"], as_index=False)[["Absatz"]].sum()

    fig_pie_Plastic_ABS_Plastic = Tortendiagramm_erstellen(pie_chart_regionen_Plastic_ABS_Plastic)


    st.subheader("Absatzverteilung der Produktgruppe 2 Plastic ABS_Plastic pro Region")
    st.plotly_chart(fig_pie_Plastic_ABS_Plastic, use_container_width=True)






#Diagramme PET Propylenoxid
with col_PET_PO1:
    balkendiagramm_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    fig_PET_PO = Saeulendiagramme_erstellen(balkendiagramm_PET_PO)

    st.subheader("Absatz- und Umsatzentwicklung der Produktgruppe 2 PET Propylenoxid")
    st.plotly_chart(fig_PET_PO, use_container_width=True)


with col_PET_PO2:
    liniendiagramm_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_PO = Liniendiagramme_erstellen(liniendiagramm_PET_PO)
    
    st.subheader("Absatzentwicklung der Produktgruppen 3 von PET Propylenoxid")
    st.plotly_chart(fig_line_PET_PO, use_container_width=True)

with col_PET_PO3:
    pie_chart_regionen_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Region_Kunde"], as_index=False)[["Absatz"]].sum()

    fig_pie_PET_PO = Tortendiagramm_erstellen(pie_chart_regionen_PET_PO)


    st.subheader("Absatzverteilung der Produktgruppe 2 PET Propylenoxid pro Region")
    st.plotly_chart(fig_pie_PET_PO, use_container_width=True)





#Diagramme PET HighPerformance
with col_PET_HP1:
    balkendiagramm_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    # Balkendiagramm erstellen
    fig_PET_HP = Saeulendiagramme_erstellen(balkendiagramm_PET_HP)

    st.subheader("Absatz- und Umsatzentwicklung der Produktgruppe 2 PET HighPerformance")
    st.plotly_chart(fig_PET_HP, use_container_width=True)

with col_PET_HP2:
    liniendiagramm_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_HP = Liniendiagramme_erstellen(liniendiagramm_PET_HP)
    
    st.subheader("Absatzentwicklung der Produktgruppen 3 von PET HighPerformance")
    st.plotly_chart(fig_line_PET_HP, use_container_width=True)

with col_PET_HP3:
    pie_chart_regionen_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Region_Kunde"], as_index=False)[["Absatz"]].sum()

    fig_pie_PET_HP = Tortendiagramm_erstellen(pie_chart_regionen_PET_HP)


    st.subheader("Absatzverteilung der Produktgruppe 2 PET HighPerformance pro Region")
    st.plotly_chart(fig_pie_PET_HP, use_container_width=True)




#Diagramme PET Special-PETs
with col_PET_Special1:
    balkendiagramm_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()
    fig_PET_Special = Saeulendiagramme_erstellen(balkendiagramm_PET_Special)

    st.subheader("Absatz- und Umsatzentwicklung der Produktgruppe 2 PET Special-PET")
    st.plotly_chart(fig_PET_Special, use_container_width=True)


with col_PET_Special2:
    liniendiagramm_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_Special = Liniendiagramme_erstellen(liniendiagramm_PET_Special)
    
    st.subheader("Absatzentwicklung der Produktgruppen 3 von PET Special-PET")
    st.plotly_chart(fig_line_PET_Special, use_container_width=True)

with col_PET_Special3:
    pie_chart_regionen_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Region_Kunde"], as_index=False)[["Absatz"]].sum()

    fig_pie_PET_Special = Tortendiagramm_erstellen(pie_chart_regionen_PET_Special)


    st.subheader("Absatzverteilung der Produktgruppe 2 PET Special-PET pro Region")
    st.plotly_chart(fig_pie_PET_Special, use_container_width=True)

#%%