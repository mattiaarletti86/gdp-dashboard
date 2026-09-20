import pandas as pd
import streamlit as st

# Configurazione pagina per cellulare con tema pulito
st.set_page_config(page_title="Gestione Casa - Arletti", layout="centered", page_icon="🏡")

st.markdown("""
<style>
    .main { background-color: #f8fafc; }
    .stMetric { background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 15px; border-radius: 14px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05); border-left: 6px solid #2563eb; }
    div.stButton > button { background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%); color: white; border-radius: 10px; font-weight: bold; border: none; padding: 10px 20px; width: 100%; }
    h1 { color: #1e293b; font-weight: 800; font-size: 1.8rem !important; }
    h2, h3 { color: #334155; }
</style>
""", unsafe_allow_html=True)

st.title("🏡 Spese & Casa - Arletti")
st.markdown("💡 *Controllo bilancio, arredi, costi futuri e rendering stanze.*")

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

# Inizializzazione session state per spese e rendering
if "spese_future" not in st.session_state:
    st.session_state.spese_future = pd.DataFrame(columns=["Mese/Anno", "Categoria", "Importo (€)", "Note"])

if "room_renderings" not in st.session_state:
    st.session_state.room_renderings = {}  # Dizionario per memorizzare i rendering per stanza

menu = st.selectbox("📂 Scegli la sezione:", [
    "📊 Dashboard & Grafici Colori", 
    "➕ Inserisci Costi Futuri", 
    "🖼️ Rendering & Planimetrie Stanze",
    "🎯 Simulatore Risparmio Mobili", 
    "🪑 Lista Mobili (15k €)", 
    "🏠 Bilancio Nuova Casa"
])

st.markdown("---")

# --- 1. DASHBOARD & GRAFICI COLORI ---
if menu == "📊 Dashboard & Grafici Colori":
    st.subheader("📊 Panoramica Spese Medie")
    
    try:
        medie_df = df_costi.iloc[0:9, [14, 15]].dropna()
        medie_df.columns = ["Categoria", "Media"]
        medie_df["Media"] = pd.to_numeric(medie_df["Media"])
        
        totale_medio = medie_df["Media"].sum()
        st.metric(label="💳 Spesa Media Mensile Totale", value=f"{totale_medio:,.2f} €")
        
        st.write("")
        st.write("### 📈 Distribuzione per Categoria")
        st.bar_chart(medie_df.set_index("Categoria"), color="#3b82f6")
        
        if not st.session_state.spese_future.empty:
            st.markdown("---")
            st.subheader("📅 Costi Futuri Programmabili")
            st.dataframe(st.session_state.spese_future, use_container_width=True)
            
            totale_futuro = st.session_state.spese_future["Importo (€)"].sum()
            st.metric(label="📌 Totale Costi Futuri Aggiunti", value=f"{totale_futuro:,.2f} €")
            
    except Exception as e:
        st.error(f"Errore nella generazione dei grafici: {e}")

# --- 2. INSERISCI COSTI FUTURI ---
elif menu == "➕ Inserisci Costi Futuri":
    st.subheader("➕ Pianifica Spesa Futura")
    st.write("Aggiungi e suddividi i costi futuri per categoria.")
    
    with st.form("form_spesa_futura"):
        col1, col2 = st.columns(2)
        with col1:
            mese_anno = st.selectbox("Mese:", [
                "Ottobre 2026", "Novembre 2026", "Dicembre 2026", 
                "Gennaio 2027", "Febbraio 2027", "Marzo 2027", 
                "Aprile 2027", "Maggio 2027", "Giugno 2027", "Luglio 2027", 
                "Agosto 2027", "Settembre 2027", "Ottobre 2027"
            ])
        with col2:
            importo = st.number_input("Importo [€]:", min_value=0.0, step=50.0, value=150.0)
        
        categoria = st.selectbox("Categoria:", [
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
        
        note = st.text_input("Note (es. negozio, descrizione):", "")
        
        submitted = st.form_submit_button("Salva Spesa Futura")
        
        if submitted:
            nuova_riga = pd.DataFrame({
                "Mese/Anno": [mese_anno],
                "Categoria": [categoria],
                "Importo (€)": [importo],
                "Note": [note]
            })
            st.session_state.spese_future = pd.concat([st.session_state.spese_future, nuova_riga], ignore_index=True)
            st.success("✅ Spesa futura aggiunta con successo!")

    if not st.session_state.spese_future.empty:
        st.markdown("### 📋 Elenco Spese Inserite")
        st.dataframe(st.session_state.spese_future, use_container_width=True)
        
        if st.button("🗑️ Svuota elenco costi futuri"):
            st.session_state.spese_future = pd.DataFrame(columns=["Mese/Anno", "Categoria", "Importo (€)", "Note"])
            st.rerun()

# --- 3. RENDERING & PLANIMETRIE STANZE ---
elif menu == "🖼️ Rendering & Planimetrie Stanze":
    st.subheader("🖼️ Rendering & Planimetrie delle Stanze")
    st.write("Carica e visualizza i rendering fotografici o le planimetrie di ogni ambiente.")
    
    stanza = st.selectbox("Seleziona Stanza / Ambiente:", [
        "Cucina", 
        "Salotto", 
        "Ingresso / Armadio ingresso", 
        "Camera Matrimoniale", 
        "Camera Bimbe", 
        "Bagno Piano Terra", 
        "Bagno Ammezzato", 
        "Bagno Piano Primo", 
        "Mansarda", 
        "Ripostiglio / Altro"
    ])
    
    uploaded_files = st.file_uploader(f"Carica immagini / rendering per: {stanza}", type=["png", "jpg", "jpeg"], accept_multiple_files=True)
    
    if uploaded_files:
        if stanza not in st.session_state.room_renderings:
            st.session_state.room_renderings[stanza] = []
        for file in uploaded_files:
            if file not in st.session_state.room_renderings[stanza]:
                st.session_state.room_renderings[stanza].append(file)
                
    # Visualizzazione galleria per la stanza selezionata
    if stanza in st.session_state.room_renderings and st.session_state.room_renderings[stanza]:
        st.markdown(f"### 📷 Immagini salvate per: *{stanza}*")
        for i, img_file in enumerate(st.session_state.room_renderings[stanza]):
            st.image(img_file, caption=f"{stanza} - Immagine {i+1}", use_column_width=True)
            if st.button(f"Elimina immagine {i+1} da {stanza}", key=f"del_{stanza}_{i}"):
                st.session_state.room_renderings[stanza].pop(i)
                st.rerun()
    else:
        st.info(f"Nessun rendering caricato per {stanza.lower()}. Usa il pulsante sopra per caricarne uno direttamente dal cellulare!")

# --- 4. SIMULATORE RISPARMIO ---
elif menu == "🎯 Simulatore Risparmio Mobili":
    st.subheader("🎯 Simulatore Risparmio Arredi")
    st.write("Calcola quanto accantonare al mese per l'obiettivo arredi.")
    
    obiettivo = st.number_input("Costo totale obiettivo [€]:", value=15000.0, step=500.0)
    mesi = st.slider("Mesi rimanenti:", min_value=1, max_value=36, value=13)
    
    risparmio_mensile = obiettivo / mesi
    
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%); padding: 20px; border-radius: 14px; border-left: 6px solid #10b981; text-align: center;">
        <h3 style="color: #065f46; margin: 0;">Obiettivo: {obiettivo:,.2f} € in {mesi} mesi</h3>
        <p style="color: #047857; font-size: 1.1rem; margin-top: 10px;">Risparmio mensile consigliato:</p>
        <h2 style="color: #047857; font-size: 2.2rem; margin: 0;">{risparmio_mensile:,.2f} € / mese</h2>
    </div>
    """, unsafe_allow_html=True)

# --- 5. LISTA MOBILI ---
elif menu == "🪑 Lista Mobili (15k €)":
    st.subheader("🪑 Controllo Mobili & Arredi")
    if not df_mobili.empty:
        mobili_clean = df_mobili.iloc[:, [0, 1, 3]].dropna(how="all")
        mobili_clean.columns = ["Articolo", "Costo", "Negozio"]
        
        for index, row in mobili_clean.iterrows():
            st.markdown(f"""
            <div style="background: white; padding: 14px; border-radius: 12px; margin-bottom: 10px; border: 1px solid #e2e8f0; box-shadow: 0 1px 3px rgba(0,0,0,0.02);">
                <b style="color: #1e293b; font-size: 1.05rem;">{row['Articolo']}</b><br>
                <span style="color: #2563eb; font-weight: bold;">💰 {row['Costo']} €</span> &nbsp;|&nbsp; 
                <span style="color: #64748b; font-style: italic;">🏪 {row['Negozio']}</span>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Nessun mobile trovato.")

# --- 6. BILANCIO NUOVA CASA ---
elif menu == "🏠 Bilancio Nuova Casa":
    st.subheader("🏠 Piano Finanziario Nuova Casa")
    if not df_casa.empty:
        st.dataframe(df_casa.dropna(how="all"), use_command_width=True)
    else:
        st.info("Dati non disponibili.")
