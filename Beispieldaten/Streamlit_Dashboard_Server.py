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
# ---- SIDEBAR ----
st.sidebar.header("Bitte Hier Filtern:")
pg1 = st.sidebar.multiselect(
    "Produktgruppe 1 wählen",
    options=sorted(working_data["Produktgruppe1"].unique(), reverse=False),
    default=sorted(working_data["Produktgruppe1"].unique(), reverse=False)
)

pg2 = st.sidebar.multiselect(
    "Produktgruppe 2 wählen",
    options=sorted(working_data["Produktgruppe2"].unique(), reverse=False),
    default=sorted(working_data["Produktgruppe2"].unique(), reverse=False)
)

jahr_gefiltert = st.sidebar.multiselect(
    "Bitte Jahr wählen:",
    options=sorted(working_data["Geschaeftsjahr"].unique(), reverse=True),
    default=sorted(working_data["Geschaeftsjahr"].unique(), reverse=True)
)

region = st.sidebar.multiselect(
    "Bitte Region wählen:",
    options=sorted(working_data["Region_Kunde"].unique(), reverse=False),
    default=sorted(working_data["Region_Kunde"].unique(), reverse=False)
)

df_selection =working_data.query(
    "Produktgruppe1 == @pg1 & Produktgruppe2 == @pg2 & Geschaeftsjahr == @jahr_gefiltert & Region_Kunde == @region"
)

#Spalten Container fuer subheader
st.subheader("Materialübersicht")

#Spalten-Container einfuegen
left_column, right_column = st.columns(2)

with left_column:
    filter_material = st.text_input('Bitte Materialnummer eingeben:')

with right_column:
    filter_country = st.text_input('Bitte vollständigen Landesnamen eingeben:')

if filter_material:
    df_selection = df_selection[df_selection['Materialnummer'].str.contains(filter_material)]

if filter_country:
    df_selection = df_selection[df_selection['Land_Kunde'] == filter_country]


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

col1, col2 = st.columns([0.9,0.1])
buffer = io.BytesIO()

filterkriterien = pd.DataFrame(data={
    "Filter": ["Produktgruppe 1", "Produktgruppe 2", "Geschäftsjahr", "Region", "Materialnummer", "Land Kunde"],
    "Gesetzte Werte": [str(sorted(df_selection["Produktgruppe1"].unique(), reverse=False)), 
                       str(sorted(df_selection["Produktgruppe2"].unique(), reverse=False)), 
                       str(sorted(df_selection["Geschaeftsjahr"].unique(), reverse=True)), 
                       str(sorted(df_selection["Region_Kunde"].unique(), reverse=False)),
                       str(filter_material),
                       str(filter_country)]
})

i = 0
while(i < len(filterkriterien)):
    if(filterkriterien.iloc[i, 1]== ""):
        filterkriterien.iloc[i, 1] = "-"
    i = i+1


# Create a Pandas Excel writer using XlsxWriter as the engine.
with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:

    filterkriterien.to_excel(writer, sheet_name='Tabelle1', startrow=0, index=False)
    #Convert the dataframe to an XlsxWriter Excel object
    df_selection.to_excel(writer, sheet_name='Tabelle1', startrow=len(filterkriterien)+3, index=False)

    # Get the xlsxwriter objects from the dataframe writer object
    #The Workbook and Worksheet objects can be used to access other XlsxWriter feature
    workbook = writer.book
    worksheet = writer.sheets["Tabelle1"]

    #get dimensions of dataframe
    (max_row, max_col) = df_selection.shape

    #create list of column-headers to use in add_table()
    column_settings = [{"header": column} for column in df_selection.columns]

    #add excel table structure. Pandas will add the data
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


    worksheet.hide_gridlines(2)

    worksheet.autofit()





with col2:
    st.download_button(
        label=":inbox_tray: Download Excel workbook",
        data=buffer,
        file_name="Absatzdaten.xlsx",
        mime="application/vnd.ms-excel",
        use_container_width=True
    )

st.markdown("---")

col_ISO_MDI1, col_ISO_MDI2, col_ISO_MDI3 = st.columns(3)

col_ISO_TDI1, col_ISO_TDI2, col_ISO_TDI3 = st.columns(3)

col_PET_PO1, col_PET_PO2, col_PET_PO3 = st.columns(3)

col_PET_HP1, col_PET_HP2, col_PET_HP3 = st.columns(3)

col_PET_Special1, col_PET_Special2, col_PET_Special3 = st.columns(3)


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
                                    values=df["Umsatz"], 
                                    names=df["Region_Kunde"],
                                    #color=['#9ecae1', '#fdae61', '#a1d99b', '#ffe34d', '#d73027']
                )
    
    return tortendiagramm
    
    
#Diagramme ISO MDI
with col_ISO_MDI1:
    balkendiagramm_ISO_MDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    # Saeulendiagramm erstellen
    fig_ISO_MDI = Saeulendiagramme_erstellen(balkendiagramm_ISO_MDI)

    st.subheader("Absatz- und Umsatzentwicklung ISO MDI Ebene PG2")
    st.plotly_chart(fig_ISO_MDI, use_container_width=True)

with col_ISO_MDI2:
    liniendiagramm_ISO_MDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_ISO_MDI = Liniendiagramme_erstellen(liniendiagramm_ISO_MDI)
    
    st.subheader("Absatzentwicklung ISO MDI Ebene PG3")
    st.plotly_chart(fig_line_ISO_MDI, use_container_width=True)

with col_ISO_MDI3:
    pie_chart_regionen_ISO_MDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '11'").groupby(["Region_Kunde"], as_index=False)[["Umsatz"]].sum()

    fig_pie_ISO_MDI = Tortendiagramm_erstellen(pie_chart_regionen_ISO_MDI)


    st.subheader("Umsatzentwicklung ISO MDI Ebene PG2 pro Region")
    st.plotly_chart(fig_pie_ISO_MDI, use_container_width=True)




#Diagramme ISO TDI
with col_ISO_TDI1:
    balkendiagramm_ISO_TDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    fig_ISO_TDI = Saeulendiagramme_erstellen(balkendiagramm_ISO_TDI)

    st.subheader("Absatz- und Umsatzentwicklung ISO TDI Ebene PG2")
    st.plotly_chart(fig_ISO_TDI, use_container_width=True)


with col_ISO_TDI2:
    liniendiagramm_ISO_TDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_ISO_TDI = Liniendiagramme_erstellen(liniendiagramm_ISO_TDI)
    
    st.subheader("Absatzentwicklung ISO TDI Ebene PG3")
    st.plotly_chart(fig_line_ISO_TDI, use_container_width=True)

with col_ISO_TDI3:
    pie_chart_regionen_ISO_TDI = df_selection.query("Produktgruppe1 == '01' & Produktgruppe2 == '21'").groupby(["Region_Kunde"], as_index=False)[["Umsatz"]].sum()

    fig_pie_ISO_TDI = Tortendiagramm_erstellen(pie_chart_regionen_ISO_TDI)


    st.subheader("Umsatzentwicklung ISO TDI Ebene PG2 pro Region")
    st.plotly_chart(fig_pie_ISO_TDI, use_container_width=True)






#Diagramme PET PO
with col_PET_PO1:
    balkendiagramm_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    fig_PET_PO = Saeulendiagramme_erstellen(balkendiagramm_PET_PO)

    st.subheader("Absatz- und Umsatzentwicklung PET PO Ebene PG2")
    st.plotly_chart(fig_PET_PO, use_container_width=True)


with col_PET_PO2:
    liniendiagramm_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_PO = Liniendiagramme_erstellen(liniendiagramm_PET_PO)
    
    st.subheader("Absatzentwicklung PET PO Ebene PG3")
    st.plotly_chart(fig_line_PET_PO, use_container_width=True)

with col_PET_PO3:
    pie_chart_regionen_PET_PO = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '41'").groupby(["Region_Kunde"], as_index=False)[["Umsatz"]].sum()

    fig_pie_PET_PO = Tortendiagramm_erstellen(pie_chart_regionen_PET_PO)


    st.subheader("Umsatzentwicklung PET PO Ebene PG2 pro Region")
    st.plotly_chart(fig_pie_PET_PO, use_container_width=True)





#Diagramme PET HighPerformance
with col_PET_HP1:
    balkendiagramm_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()

    # Balkendiagramm erstellen
    fig_PET_HP = Saeulendiagramme_erstellen(balkendiagramm_PET_HP)

    st.subheader("Absatz- und Umsatzentwicklung PET HighPerformance Ebene PG2")
    st.plotly_chart(fig_PET_HP, use_container_width=True)

with col_PET_HP2:
    liniendiagramm_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_HP = Liniendiagramme_erstellen(liniendiagramm_PET_HP)
    
    st.subheader("Absatzentwicklung PET HighPerformance Ebene PG3")
    st.plotly_chart(fig_line_PET_HP, use_container_width=True)

with col_PET_HP3:
    pie_chart_regionen_PET_HP = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '43'").groupby(["Region_Kunde"], as_index=False)[["Umsatz"]].sum()

    fig_pie_PET_HP = Tortendiagramm_erstellen(pie_chart_regionen_PET_HP)


    st.subheader("Umsatzentwicklung PET HighPerformance Ebene PG2 pro Region")
    st.plotly_chart(fig_pie_PET_HP, use_container_width=True)




#Diagramme PET Special-PETs

with col_PET_Special1:
    balkendiagramm_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Produktgruppe2", "Geschaeftsjahr"], as_index=False)[["Absatz", "Umsatz"]].sum()
    fig_PET_Special = Saeulendiagramme_erstellen(balkendiagramm_PET_Special)

    st.subheader("Absatz- und Umsatzentwicklung PET Special-PET Ebene PG2")
    st.plotly_chart(fig_PET_Special, use_container_width=True)


with col_PET_Special2:
    liniendiagramm_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Produktgruppe2", "Produktgruppe3", "Geschaeftsjahr"], as_index=False)[["Absatz"]].sum()

    fig_line_PET_Special = Liniendiagramme_erstellen(liniendiagramm_PET_Special)
    
    st.subheader("Absatzentwicklung PET Special-PET Ebene PG3")
    st.plotly_chart(fig_line_PET_Special, use_container_width=True)

with col_PET_Special3:
    pie_chart_regionen_PET_Special = df_selection.query("Produktgruppe1 == '04' & Produktgruppe2 == '49'").groupby(["Region_Kunde"], as_index=False)[["Umsatz"]].sum()

    fig_pie_PET_Special = Tortendiagramm_erstellen(pie_chart_regionen_PET_Special)


    st.subheader("Umsatzentwicklung PET Special-PET Ebene PG2 pro Region")
    st.plotly_chart(fig_pie_PET_Special, use_container_width=True)

#%%