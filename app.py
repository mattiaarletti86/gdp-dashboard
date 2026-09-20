import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Configurazione pagina per cellulare
st.set_page_config(page_title="Gestione Casa & Spese - Arletti", layout="centered")

st.title("🏡 Spese & Casa - Arletti")
st.write("Dashboard finanziaria, previsioni e inserimento costi futuri.")

# Caricamento dati da Excel
file_path = "Spese casa -2.xlsx"

@st.cache_data
def load_data():
    df_costi = pd.read_excel(file_path, sheet_name="Costi famiglia")
    df_mobili = pd.read_excel(file_path, sheet_name="Mobili", skiprows=1)
    df_casa = pd.read_excel(file_path, sheet_name="Piano acquisto nuova casa")
    return df_costi, df_mobili, df_casa

try:
    df_costi, df_mobili, df_casa = load_data()
except Exception as e:
    st.error(f"Errore nel caricamento del file Excel: {e}")
    st.stop()

# Inizializzazione della memoria temporanea per i costi futuri aggiunti dall'utente
if "spese_future" not in st.session_state:
    st.session_state.spese_future = pd.DataFrame(columns=["Mese/Anno", "Categoria", "Importo (€)", "Note"])

# Menu di navigazione a tendina (ottimizzato per mobile)
menu = st.selectbox("Seleziona sezione:", [
    "📊 Dashboard & Grafici Visivi", 
    "➕ Inserisci Costi Futuri", 
    "🎯 Simulatore Risparmio Mobili", 
    "🪑 Lista Mobili (15k €)", 
    "🏠 Bilancio Nuova Casa"
])

# --- 1. DASHBOARD & GRAFICI VISIVI ---
if menu == "📊 Dashboard & Grafici Visivi":
    st.subheader("Panoramica Visiva delle Spese")
    
    try:
        medie_df = df_costi.iloc[0:9, [14, 15]].dropna()
        medie_df.columns = ["Categoria", "Media"]
        medie_df["Media"] = pd.to_numeric(medie_df["Media"])
        
        # Metriche generali in evidenza
        totale_medio = medie_df["Media"].sum()
        st.metric(label="Spesa Media Mensile Totale", value=f"{totale_medio:,.2f} €")
        
        # Grafico a barre interattivo con Streamlit
        st.write("### Spesa Media per Categoria")
        st.bar_chart(medie_df.set_index("Categoria"))
        
        # Se l'utente ha inserito costi futuri, mostriamo anche quelli
        if not st.session_state.spese_future.empty:
            st.write("### 📅 Costi Futuri Aggiunti")
            st.dataframe(st.session_state.spese_future, use_container_width=True)
            
            totale_futuro = st.session_state.spese_future["Importo (€)"].sum()
            st.metric(label="Totale Costi Futuri Programmati", value=f"{totale_futuro:,.2f} €")
            
    except Exception as e:
        st.error(f"Errore nella generazione dei grafici: {e}")

# --- 2. INSERISCI COSTI FUTURI ---
elif menu == "➕ Inserisci Costi Futuri":
    st.subheader("Aggiungi Spesa o Costo Futuro")
    st.write("Pianifica nuove spese suddividendole per categoria.")
    
    with st.form("form_spesa_futura"):
        mese_anno = st.selectbox("Mese di riferimento:", [
            "Ottobre 2026", "Novembre 2026", "Dicembre 2026", 
            "Gennaio 2027", Febbraio 2027 := "Febbraio 2027", "Marzo 2027", 
            "Aprile 2027", "Maggio 2027", "Giugno 2027", "Luglio 2027", 
            "Agosto 2027", "Settembre 2027", "Ottobre 2027"
        ])
        
        categoria = st.selectbox("Categoria di spesa:", [
            "Costo alimentare mensile", 
            "Tempo libero e viaggi", 
            "Utenze", 
            "Scuola e sport", 
            "Trasporti e auto", 
            "Casa e assicurazioni", 
            "Shopping", 
            "Farmacia e cura della persona",
            "Arredi e Extra Nuova Casa"
        ])
        
        importo = st.number_input("Importo previsto [€]:", min_value=0.0, step=50.0, value=100.0)
        note = st.text_input("Note aggiuntive (opzionale):", "")
        
        submitted = st.form_submit_button("Aggiungi alla previsione")
        
        if submitted:
            nuova_riga = pd.DataFrame({
                "Mese/Anno": [mese_anno],
                "Categoria": [categoria],
                "Importo (€)": [importo],
                "Note": [note]
            })
            st.session_state.spese_future = pd.concat([st.session_state.spese_future, nuova_riga], ignore_index=True)
            st.success("Spesa futura aggiunta con successo!")

    if not st.session_state.spese_future.empty:
        st.write("### Elenco Spese Future Inserite")
        st.dataframe(st.session_state.spese_future, use_container_width=True)
        
        if st.button("Pulisci elenco spese future"):
            st.session_state.spese_future = pd.DataFrame(columns=["Mese/Anno", "Categoria", "Importo (€)", "Note"])
            st.rerun()

# --- 3. SIMULATORE RISPARMIO ---
elif menu == "🎯 Simulatore Risparmio Mobili":
    st.subheader("Simulatore Risparmio Arredi")
    st.write("Calcola il risparmio mensile per arrivare a ottobre 2027.")
    
    obiettivo = st.number_input("Costo totale obiettivo [€]:", value=15000.0, step=500.0)
    mesi = st.slider("Mesi rimanenti:", min_value=1, max_value=36, value=13)
    
    risparmio_mensile = obiettivo / mesi
    st.success(f"💡 Per l'obiettivo di **{obiettivo:,.2f} €** in **{mesi} mesi**, devi risparmiare:\n### **{risparmio_mensile:,.2f} € al mese**")

# --- 4. LISTA MOBILI ---
elif menu == "🪑 Lista Mobili (15k €)":
    st.subheader("Controllo Mobili & Arredi")
    if not df_mobili.empty:
        mobili_clean = df_mobili.iloc[:, [0, 1, 3]].dropna(how="all")
        mobili_clean.columns = ["Articolo", "Costo", "Negozio"]
        for index, row in mobili_clean.iterrows():
            st.container()
            st.write(f"**{row['Articolo']}** — 💰 {row['Costo']} €  \n🏪 *Negozio: {row['Negozio']}*")
            st.divider()
    else:
        st.info("Nessun mobile trovato.")

# --- 5. BILANCIO NUOVA CASA ---
elif menu == "🏠 Bilancio Nuova Casa":
    st.subheader("Piano Finanziario Nuova Casa")
    if not df_casa.empty:
        st.dataframe(df_casa.dropna(how="all"), use_container_width=True)
    else:
        st.info("Dati non disponibili.")
    st.write("Calcola il risparmio mensile per arrivare a ottobre 2027.")
    
    obiettivo = st.number_input("Costo totale obiettivo [€]:", value=15000.0, step=500.0)
    mesi = st.slider("Mesi rimanenti:", min_value=1, max_value=36, value=13)
    
    risparmio_mensile = obiettivo / mesi
    st.success(f"💡 Per l'obiettivo di **{obiettivo:,.2f} €** in **{mesi} mesi**, devi risparmiare:\n### **{risparmio_mensile:,.2f} € al mese**")

# --- 2. MEDIE COSTI FAMIGLIA ---
elif menu == "📊 Medie Costi Famiglia":
    st.subheader("Medie Mensili Spese")
    try:
        medie_df = df_costi.iloc[0:9, [14, 15]].dropna()
        medie_df.columns = ["Categoria", "Media (€)"]
        for index, row in medie_df.iterrows():
            st.metric(label=str(row["Categoria"]), value=f"{float(row['Media (€)']):.2f} €")
    except Exception as e:
        st.error(f"Errore nella lettura delle medie: {e}")

# --- 3. LISTA MOBILI ---
elif menu == "🪑 Lista Mobili (15k €)":
    st.subheader("Controllo Mobili & Arredi")
    if not df_mobili.empty:
        # Prende le prime colonne utili: Articolo, Costo, Negozio
        mobili_clean = df_mobili.iloc[:, [0, 1, 3]].dropna(how="all")
        mobili_clean.columns = ["Articolo", "Costo", "Negozio"]
        for index, row in mobili_clean.iterrows():
            st.container()
            st.write(f"**{row['Articolo']}** — 💰 {row['Costo']} €  \n🏪 *Negozio: {row['Negozio']}*")
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
    st.write("Calcola il risparmio mensile per arrivare a ottobre 2027.")
    
    obiettivo = st.number_input("Costo totale obiettivo [€]:", value=15000.0, step=500.0)
    mesi = st.slider("Mesi rimanenti:", min_value=1, max_value=36, value=13)
    
    risparmio_mensile = obiettivo / mesi
    st.success(f"💡 Per l'obiettivo di **{obiettivo:,.2f} €** in **{mesi} mesi**, devi risparmiare:\n### **{risparmio_mensile:,.2f} € al mese**")

# --- 2. MEDIE COSTI FAMIGLIA ---
elif menu == "📊 Medie Costi Famiglia":
    st.subheader("Medie Mensili Spese")
    try:
        medie_df = df_costi.iloc[0:9, [14, 15]].dropna()
        medie_df.columns = ["Categoria", "Media (€)"]
        for index, row in medie_df.iterrows():
            st.metric(label=str(row["Categoria"]), value=f"{float(row['Media (€)']):.2f} €")
    except Exception as e:
        st.error(f"Errore nella lettura delle medie: {e}")

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
