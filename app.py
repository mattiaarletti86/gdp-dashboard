import pandas as pd
import streamlit as st

# Configurazione pagina per cellulare
st.set_page_config(page_title="Gestione Casa - Arletti", layout="centered")

st.title("🏡 Spese & Casa - Arletti")
st.write("Tool di previsione costi e gestione bilancio familiare.")

# Caricamento dati da Excel
file_path = "Spese casa -2.xlsx"

@st.cache_data
def load_data():
    df_costi = pd.read_excel(file_path, sheet_name="Costi famiglia")
    df_mobili = pd.read_excel(file_path, sheet_name="Mobili")
    df_casa = pd.read_excel(file_path, sheet_name="Piano acquisto nuova casa")
    return df_costi, df_mobili, df_casa

try:
    df_costi, df_mobili, df_casa = load_data()
except Exception as e:
    st.error(f"Errore nel caricamento del file Excel: {e}")
    st.stop()

# Menu di navigazione a tendina (ottimizzato per mobile)
menu = st.selectbox("Seleziona sezione:", [
    "🎯 Simulatore Risparmio Mobili", 
    "📊 Medie Costi Famiglia", 
    "🪑 Lista Mobili (15k €)", 
    "🏠 Bilancio Nuova Casa"
])

# --- 1. SIMULATORE RISPARMIO ---
if menu == "🎯 Simulatore Risparmio Mobili":
    st.subheader("Simulatore Risparmio Arredi")
    st.write("Calcola il risparmio mensile per arrivare a ottobre 2027.")
    
    obiettivo = st.number_input("Costo totale obiettivo [€]:", value=15000.0, step=500.0)
    mesi = st.slider("Mesi rimanenti:", min_value=1, max_value=36, value=13)
    
    risparmio_mensile = obiettivo / mesi
    st.success(f"💡 Per l'obiettivo di **{obiettivo:,.2f} €** in **{mesi} mesi**, devi risparmiare:\n### **{risparmio_mensile:,.2f} € al mese**")

# --- 2. MEDIE COSTI FAMIGLIA ---
elif menu == "📊 Medie Costi Famiglia":
    st.subheader("Medie Mensili Spese")
    try:
        medie_df = df_costi.iloc[0:9, [0, 14]].dropna()
        medie_df.columns = ["Categoria", "Media (€)"]
        for index, row in medie_df.iterrows():
            st.metric(label=row["Categoria"], value=f"{row['Media (€)']:.2f} €")
    except Exception as e:
        st.error("Errore nella lettura delle medie.")

# --- 3. LISTA MOBILI ---
elif menu == "🪑 Lista Mobili (15k €)":
    st.subheader("Controllo Mobili & Arredi")
    if not df_mobili.empty:
        mobili_clean = df_mobili[["Articolo", "Costo", "NEGOZIO"]].dropna(how="all")
        for index, row in mobili_clean.iterrows():
            st.container()
            st.write(f"**{row['Articolo']}** — 💰 {row['Costo']} €  \n🏪 *Negozio: {row['NEGOZIO']}*")
            st.divider()
    else:
        st.info("Nessun mobile trovato.")

# --- 4. BILANCIO NUOVA CASA ---
elif menu == "🏠 Bilancio Nuova Casa":
    st.subheader("Piano Finanziario Nuova Casa")
    if not df_casa.empty:
        st.dataframe(df_casa.dropna(how="all"), use_container_width=True)
    else:
        st.info("Dati non disponibili.")
