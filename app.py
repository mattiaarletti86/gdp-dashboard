import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# Configurazione pagina per cellulare
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
st.markdown("💡 *Controllo bilancio, storico spese, arredi, costi futuri e rendering.*")

file_path = "Spese casa -2.xlsx"

@st.cache_data
def load_data():
    df_costi = pd.read_excel(file_path, sheet_name="Costi famiglia")
    df_mobili = pd.read_excel(file_path, sheet_name="Mobili", skiprows=1)
    return df_costi, df_mobili

try:
    df_costi, df_mobili = load_data()
except Exception as e:
    st.error(f"Errore nel caricamento del file Excel: {e}")
    st.stop()

# Estrazione e pulizia dello storico spese per periodo e sottogruppo
@st.cache_data
def parse_storico_spese(df_costi):
    records = []
    current_month = "Ottobre 2025"
    
    for idx, row in df_costi.iterrows():
        val0 = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        if "Spese " in val0 or "SPESE " in val0 or "spese " in val0:
            m_name = val0.replace("Spese", "").replace("SPESE", "").replace("spese", "").strip()
            if m_name:
                current_month = m_name.capitalize()
        elif val0 and not val0.startswith("MEDIA"):
            for col_idx in range(1, len(row)):
                val_c = row.iloc[col_idx]
                col_header = str(df_costi.columns[col_idx]).strip()
                if pd.notna(val_c) and isinstance(val_c, (int, float)) and val_c > 0:
                    subgroup = col_header if not col_header.startswith("Unnamed") and not col_header.startswith("MEDIA") else "Generale"
                    records.append({
                        "Mese/Periodo": current_month,
                        "Categoria": val0,
                        "Sottogruppo": subgroup,
                        "Importo (€)": float(val_c)
                    })
    df_parsed = pd.DataFrame(records)
    if df_parsed.empty:
        df_parsed = pd.DataFrame([
            {"Mese/Periodo": "Ottobre 2025", "Categoria": "Costo alimentare mensile", "Sottogruppo": "Supermercato", "Importo (€)": 883.14},
            {"Mese/Periodo": "Ottobre 2025", "Categoria": "Utenze", "Sottogruppo": "Luce & Gas", "Importo (€)": 309.71},
            {"Mese/Periodo": "Novembre 2025", "Categoria": "Costo alimentare mensile", "Sottogruppo": "Supermercato", "Importo (€)": 820.00},
            {"Mese/Periodo": "Novembre 2025", "Categoria": "Trasporti e auto", "Sottogruppo": "Carburante", "Importo (€)": 210.50},
        ])
    return df_parsed

df_storico = parse_storico_spese(df_costi)

# Inizializzazione Session State
if "spese_future" not in st.session_state:
    st.session_state.spese_future = pd.DataFrame(columns=["Mese/Anno", "Categoria", "Importo (€)", "Note"])

if "room_renderings" not in st.session_state:
    st.session_state.room_renderings = {}

if "piano_costi" not in st.session_state:
    st.session_state.piano_costi = pd.DataFrame([
        {"Voce di Spesa": "Costo acquisto casa", "Importo (€)": 400000.0},
        {"Voce di Spesa": "Costo acquisto garage", "Importo (€)": 25000.0},
        {"Voce di Spesa": "IVA su acquisto casa", "Importo (€)": 16000.0},
        {"Voce di Spesa": "Trasloco", "Importo (€)": 5000.0},
        {"Voce di Spesa": "Istruttoria mutuo", "Importo (€)": 2000.0},
        {"Voce di Spesa": "Notaio", "Importo (€)": 10000.0},
        {"Voce di Spesa": "Acquisto mobili", "Importo (€)": 4785.0},
        {"Voce di Spesa": "Allacciamenti", "Importo (€)": 1000.0},
        {"Voce di Spesa": "Fuori capitolato", "Importo (€)": 5000.0},
        {"Voce di Spesa": "Ristrutturazione & Opere Extra", "Importo (€)": 124000.0}
    ])

if "piano_ricavi" not in st.session_state:
    st.session_state.piano_ricavi = pd.DataFrame([
        {"Fonte / Entrata": "Vendita casa / Liquidità disponibile", "Importo (€)": 381000.0},
        {"Fonte / Entrata": "Mutuo o Risparmi dedicati", "Importo (€)": 200000.0}
    ])

menu = st.selectbox("📂 Scegli la sezione:", [
    "📜 Monitor Mesi Precedenti & Filtri",
    "🏠 Bilancio Nuova Casa (Modificabile)",
    "📊 Dashboard & Grafici Colori", 
    "➕ Inserisci Costi Futuri", 
    "🖼️ Rendering & Planimetrie Stanze",
    "🎯 Simulatore Risparmio Mobili", 
    "🪑 Lista Mobili (15k €)"
])

st.markdown("---")

# --- 1. MONITOR MESI PRECEDENTI & FILTRI ---
if menu == "📜 Monitor Mesi Precedenti & Filtri":
    st.subheader("📜 Monitor Spese Mesi Precedenti")
    st.write("Filtra per periodo e sottogruppo per visualizzare i costi e il totale calcolato.")
    
    # 1. Filtro Mese / Periodo
    lista_mesi = sorted(list(df_storico["Mese/Periodo"].unique()))
    mesi_selezionati = st.multiselect("🗓️ Seleziona Periodo / Mesi:", options=lista_mesi, default=lista_mesi)
    
    # 2. Filtro Categoria / Sottogruppo
    df_filtrato_mesi = df_storico[df_storico["Mese/Periodo"].isin(mesi_selezionati)] if mesi_selezionati else df_storico
    lista_categorie = sorted(list(df_filtrato_mesi["Categoria"].unique()))
    
    categorie_selezionate = st.multiselect("🏷️ Seleziona Categoria / Sottogruppo:", options=lista_categorie, default=lista_categorie)
    
    # Applicazione filtri
    df_finale = df_filtrato_mesi[df_filtrato_mesi["Categoria"].isin(categorie_selezionate)]
    
    # Calcolo Totale
    costo_totale = df_finale["Importo (€)"].sum()
    
    # Visualizzazione del Totale in Evidenza
    st.markdown(f"""
    <div style="background: linear-gradient(135deg, #eff6ff 0%, #dbeafe 100%); padding: 18px; border-radius: 14px; border-left: 6px solid #2563eb; text-align: center; margin-bottom: 20px;">
        <span style="color: #1e40af; font-size: 1.1rem; font-weight: 600;">💰 Costo Totale del Periodo Selezionato</span>
        <h2 style="color: #1d4ed8; font-size: 2.3rem; margin: 5px 0 0 0;">{costo_totale:,.2f} €</h2>
    </div>
    """, unsafe_allow_html=True)
    
    if not df_finale.empty:
        # GRAFICO A TORTA / CIAMBELLA PER PERCENTUALI
        st.write("### 🍕 Percentuale Spese per Categoria (Grafico a Torta)")
        grouped_data = df_finale.groupby("Categoria")["Importo (€)"].sum()
        
        fig, ax = plt.subplots(figsize=(6, 5))
        colors = ['#2563eb', '#3b82f6', '#60a5fa', '#93c5fd', '#1d4ed8', '#1e40af', '#6366f1', '#818cf8', '#a5b4fc']
        
        wedges, texts, autotexts = ax.pie(
            grouped_data, 
            labels=grouped_data.index, 
            autopct='%1.1f%%', 
            startangle=140,
            colors=colors[:len(grouped_data)],
            wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2)
        )
        plt.setp(autotexts, size=9, weight="bold", color="black")
        plt.setp(texts, size=10)
        ax.axis('equal')
        plt.tight_layout()
        st.pyplot(fig)
        
        st.markdown("---")
        st.write("### 📈 Istogramma Distribuzione Spese")
        st.bar_chart(grouped_data, color="#2563eb")
        
        st.write("### 📋 Dettaglio Spese Filtrate")
        st.dataframe(df_finale, use_container_width=True, hide_index=True)
    else:
        st.warning("Nessuna spesa trovata per i filtri selezionati.")

# --- 2. BILANCIO NUOVA CASA ---
elif menu == "🏠 Bilancio Nuova Casa (Modificabile)":
    st.subheader("🏠 Piano Finanziario Nuova Casa")
    st.write("Modifica gli importi o aggiungi nuove voci per ricalcolare il saldo finale.")
    
    tot_costi = st.session_state.piano_costi["Importo (€)"].sum()
    tot_ricavi = st.session_state.piano_ricavi["Importo (€)"].sum()
    netto_residuo = tot_ricavi - tot_costi
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric(label="🔴 Totale Costi", value=f"{tot_costi:,.0f} €")
    with col2:
        st.metric(label="🟢 Totale Entrate", value=f"{tot_ricavi:,.0f} €")
    with col3:
        st.metric(label="🔵 Netto Residuo", value=f"{netto_residuo:,.0f} €")
        
    st.markdown("---")
    
    st.write("### 🔴 1. Uscite e Costi Previsti")
    edited_costi = st.data_editor(
        st.session_state.piano_costi, 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True,
        key="editor_costi"
    )
    st.session_state.piano_costi = edited_costi

    st.markdown("---")

    st.write("### 🟢 2. Entrate, Mutuo e Coperture")
    edited_ricavi = st.data_editor(
        st.session_state.piano_ricavi, 
        num_rows="dynamic", 
        use_container_width=True, 
        hide_index=True,
        key="editor_ricavi"
    )
    st.session_state.piano_ricavi = edited_ricavi

    if st.button("🔄 Aggiorna e Ricalcola Saldo"):
        st.rerun()

# --- 3. DASHBOARD & GRAFICI COLORI ---
elif menu == "📊 Dashboard & Grafici Colori":
    st.subheader("📊 Panoramica Spese Medie")
    try:
        medie_df = df_costi.iloc[0:9, [14, 15]].dropna()
        medie_df.columns = ["Categoria", "Media"]
        medie_df["Media"] = pd.to_numeric(medie_df["Media"])
        
        totale_medio = medie_df["Media"].sum()
        st.metric(label="💳 Spesa Media Mensile Totale", value=f"{totale_medio:,.2f} €")
        
        st.write("")
        st.write("### 🍕 Percentuale Spesa Media (Grafico a Torta)")
        fig, ax = plt.subplots(figsize=(6, 5))
        colors = ['#2563eb', '#3b82f6', '#60a5fa', '#93c5fd', '#1d4ed8', '#1e40af', '#6366f1', '#818cf8', '#a5b4fc']
        wedges, texts, autotexts = ax.pie(
            medie_df["Media"], 
            labels=medie_df["Categoria"], 
            autopct='%1.1f%%', 
            startangle=140,
            colors=colors[:len(medie_df)],
            wedgeprops=dict(width=0.45, edgecolor='white', linewidth=2)
        )
        plt.setp(autotexts, size=9, weight="bold", color="black")
        plt.setp(texts, size=10)
        ax.axis('equal')
        plt.tight_layout()
        st.pyplot(fig)
        
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

# --- 4. INSERISCI COSTI FUTURI ---
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

# --- 5. RENDERING & PLANIMETRIE STANZE ---
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
            file_bytes = file.getvalue()
            if not any(item["name"] == file.name for item in st.session_state.room_renderings[stanza]):
                st.session_state.room_renderings[stanza].append({"name": file.name, "bytes": file_bytes})
                
    if stanza in st.session_state.room_renderings and st.session_state.room_renderings[stanza]:
        st.markdown(f"### 📷 Immagini salvate per: *{stanza}*")
        for i, img_data in enumerate(st.session_state.room_renderings[stanza]):
            st.image(img_data["bytes"], caption=f"{stanza} - {img_data['name']}", use_container_width=True)
            if st.button(f"Elimina immagine {i+1} da {stanza}", key=f"del_{stanza}_{i}"):
                st.session_state.room_renderings[stanza].pop(i)
                st.rerun()
    else:
        st.info(f"Nessun rendering caricato per {stanza.lower()}. Usa il pulsante sopra per caricarne uno dal cellulare!")

# --- 6. SIMULATORE RISPARMIO ---
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

# --- 7. LISTA MOBILI ---
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
