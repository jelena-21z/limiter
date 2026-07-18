import streamlit as st
import pandas as pd
import datetime
import plotly.express as px

st.set_page_config(layout="wide", page_title="LIMITER", page_icon="📊")

# STROGO SKRIVANJE PROFILA, MENIJA I FORK-A ZA SVE KORISNIKE NA MOBILNOM I RAČUNARU
st.markdown("""
<style>
.stApp { background-color: #f8f9fa; color: #212529; }
[data-testid="stMetricValue"] { font-size: 2.2rem !important; line-height: 1.2 !important; font-weight: 700; }
[data-testid="stMetricLabel"] { font-size: 1rem !important; color: #495057 !important; font-weight: 600; }
.main-title { font-size: 2.5rem; font-weight: 700; color: #0a2540; margin-bottom: 2px; }
.sub-title { font-size: 1.1rem; color: #495057; margin-bottom: 20px; font-weight: 500; font-style: italic; }
.section-box { padding: 25px; background-color: #ffffff; border-radius: 8px; border: 1px solid #e9ecef; box-shadow: 0 2px 4px rgba(0,0,0,0.02); margin-bottom: 25px; }
.profile-box { padding: 25px; background-color: #ffffff; border-radius: 8px; border: 2px solid #0a2540; box-shadow: 0 4px 6px rgba(0,0,0,0.04); margin-bottom: 25px; }
.group-box { padding: 25px; background-color: #f1f3f5; border-radius: 8px; border: 2px dashed #0a2540; margin-bottom: 25px; }
h3 { color: #0a2540 !important; font-weight: 700 !important; }

/* CSS ZAŠTITA: Briše profil, tri tačkice, fork i Streamlit oznake sa ekrana */
#MainMenu {visibility: hidden !important;}
footer {visibility: hidden !important;}
header {visibility: hidden !important;}
.stDeployButton {display:none !important;}
[data-testid="stToolbar"] {visibility: hidden !important; display: none !important;}
[data-testid="stDecoration"] {display: none !important;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 LIMITER</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-title">Finansijski Monitor: PDV & Paušal Analitika</div>', unsafe_allow_html=True)

# VRAĆENA ČISTA LOKALNA PUTANJA KOJA VAM RADI NA RAČUNARU
fajl_putanja = "firme_podaci.xlsx"
def ucitaj_i_analiziraj(putanja):
    df_raw = pd.read_excel(putanja)
    df_raw.columns = [str(col).strip() for col in df_raw.columns]
    df_long = pd.melt(df_raw, id_vars=['Naziv', 'Tip'], var_name='Mesec_Tekst', value_name='Prihod')
    df_long['Prihod'] = pd.to_numeric(df_long['Prihod'], errors='coerce').fillna(0)
    meseci_prevod = {'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'Maj': 5, 'Jun': 6, 'Jul': 7, 'Avg': 8, 'Sep': 9, 'Okt': 10, 'Nov': 11, 'Dec': 12}
    
    def parsiraj_mesec(tekst):
        try:
            delovi = str(tekst).strip().split()
            return datetime.date(int(delovi), meseci_prevod[delovi], 1)
        except: return None

    df_long['Datum'] = df_long['Mesec_Tekst'].apply(parsiraj_mesec)
    df_long = df_long.dropna(subset=['Datum']).sort_values(by=['Naziv', 'Datum'])
    trenutni_datum = datetime.date(2026, 6, 1)
    datum_pre_12m = datetime.date(2025, 7, 1)
    
    izvestaj = []
    for firma, grupa in df_long.groupby('Naziv'):
        cisti_tip = str(grupa['Tip'].iloc).strip().lower()
        
        je_pausal = 'paus' in cisti_tip or 'pauš' in cisti_tip
        je_udruzenje = 'udruz' in cisti_tip or 'udruž' in cisti_tip
        forma_ispis = "Paušalac" if je_pausal else ("Udruženje" if je_udruzenje else "D.O.O.")
        
        istorija_do_sada = grupa[grupa['Datum'] <= trenutni_datum]
        prosecan_mesecni_prihod = istorija_do_sada[istorija_do_sada['Prihod'] > 0]['Prihod'].mean() if len(istorija_do_sada[istorija_do_sada['Prihod'] > 0]) > 0 else 0
        
        prihod_tekuce_godine = 0
        prihod_12m = 0
        for _, red in grupa.iterrows():
            d = red['Datum']
            p = red['Prihod']
            if d.year == 2026 and d <= trenutni_datum: prihod_tekuce_godine += p
            if datum_pre_12m <= d <= trenutni_datum: prihod_12m += p
                
        proc_pausal = (prihod_tekuce_godine / 6000000) * 100 if je_pausal else 0
        proc_pdv = (prihod_12m / 8000000) * 100
        
        preostalo_meseci, tip_limita = 999, ""
        if prosecan_mesecni_prihod > 0:
            if je_pausal and (6000000 - prihod_tekuce_godine) > 0:
                preostalo_meseci = (6000000 - prihod_tekuce_godine) / prosecan_mesecni_prihod
                tip_limita = "Limit 6 mil."
            if (8000000 - prihod_12m) > 0 and ((8000000 - prihod_12m) / prosecan_mesecni_prihod) < preostalo_meseci:
                preostalo_meseci = (8000000 - prihod_12m) / prosecan_mesecni_prihod
                tip_limita = "Limit PDV"

        if preostalo_meseci == 999: predikcija_poruka = "✅ Stabilno u narednih 6+ meseci"
        elif preostalo_meseci <= 0: predikcija_poruka = f"🚨 Limit dostignut ({tip_limita})!"
        else:
            br_meseci = int(round(preostalo_meseci))
            buduci_mesec = (trenutni_datum.month + br_meseci - 1) % 12 + 1
            buduca_godina = trenutni_datum.year + (trenutni_datum.month + br_meseci - 1) // 12
            predikcija_poruka = f"⚠️ {tip_limita} za {br_meseci} mes. ({buduci_mesec}/{buduca_godina})"
            
        zona = "🔴 Visok rizik" if (proc_pdv >= 80 or proc_pausal >= 80 or preostalo_meseci <= 3) else ("🟡 Srednji rizik" if (60 <= proc_pdv < 80 or 60 <= proc_pausal < 80) else "🟢 Bezbedno")
        izvestaj.append({"Firma": firma, "Pravna Forma": forma_ispis, "Prihod (Tekuća god)": f"{prihod_tekuce_godine:,.2f} RSD" if je_pausal else "-", "Iskorišćenost 6 mil.": f"{proc_pausal:.1f}%" if je_pausal else "-", "Prihod (Poslednjih 12M)": f"{prihod_12m:,.2f} RSD", "Iskorišćenost 8 mil.": f"{proc_pdv:.1f}%", "Prognoza limita": predikcija_poruka, "Status": zona, "Sirov_Prihod_12M": prihod_12m, "Sirov_Prihod_2026": prihod_tekuce_godine})
    return pd.DataFrame(izvestaj), df_long, trenutni_datum
try:
    rezultati_df, df_sve_mesecno, korisceni_mesec = ucitaj_i_analiziraj(fajl_putanja)
    st.success(f"🎉 Podaci uspešno povučeni iz lokalne baze za presek: {korisceni_mesec.strftime('%m/%Y')}")
    
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    k1, k2, k3 = st.columns(3)
    k1.metric("🔴 Visok rizik", len(rezultati_df[rezultati_df["Status"] == "🔴 Visok rizik"]))
    k2.metric("🟡 Srednji rizik", len(rezultati_df[rezultati_df["Status"] == "🟡 Srednji rizik"]))
    k3.metric("🟢 Bezbedno", len(rezultati_df[rezultati_df["Status"] == "🟢 Bezbedno"]))
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("📋 Lista svih firmi sa dva kriterijuma i prognozom")
    izabrana_zona = st.multiselect("🔍 Filtrirajte tabelu:", ["🔴 Visok rizik", "🟡 Srednji rizik", "🟢 Bezbedno"], default=["🔴 Visok rizik", "🟡 Srednji rizik", "🟢 Bezbedno"])
    prikaz_df = rezultati_df[rezultati_df["Status"].isin(izabrana_zona)]
    
    # POPRAVLJENO I OBRISANO: Izbačene su kolone "Sirov_Prihod_12M" i "Sirov_Prihod_2026" i tabela se sama skuplja na mobilnom
    st.dataframe(prikaz_df.drop(columns=["Sirov_Prihod_12M", "Sirov_Prihod_2026"], errors="ignore"), use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="group-box">', unsafe_allow_html=True)
    st.subheader("🏢 Grupe povezanih firmi (Zbirna konsolidacija)")
    izbor_grupe = st.radio("Izaberite grupu povezane poslovne mreže za analizu:", ["Activity Grupa", "Folklorić Grupa"], horizontal=True)
    kljucna_rec = "activity" if "Activity" in izbor_grupe else "folklori"
    firme_u_grupi = rezultati_df[rezultati_df["Firma"].str.lower().str.contains(kljucna_rec)]
    
    if len(firme_u_grupi) > 0:
        zbir_12m = firme_u_grupi[firme_u_grupi["Pravna Forma"] != "Udruženje"]["Sirov_Prihod_12M"].sum()
        zbir_2026 = firme_u_grupi[firme_u_grupi["Pravna Forma"] != "Udruženje"]["Sirov_Prihod_2026"].sum()
        
        # POPRAVLJENO ZA MOBILNI: Spisak firmi je spakovan u dugme na rasklapanje, podrazumevano su sakrivene
        with st.expander("🔍 Prikaži sve firme u ovoj grupi"):
            st.write(f"📂 {', '.join(firme_u_grupi['Firma'].tolist())}")
        
        g1, g2 = st.columns(2)
        g1.metric(" Ukupan prihod grupe (Zadnjih 12M)", f"{zbir_12m:,.2f} RSD")
        g2.metric("📅 Ukupan prihod grupe (Tekuća 2026)", f"{zbir_2026:,.2f} RSD")
        
        st.markdown("### 🧮 Uporedni prikaz poreskih obaveza: Trenutni model vs Jedno D.O.O. lise")
        trenutne_obaveze, broj_pausalaca, broj_udruzenja = 0, 0, 0
        detalji_trenutnog = []
        
        for _, klijent_red in firme_u_grupi.iterrows():
            if klijent_red["Pravna Forma"] == "Paušalac":
                trenutne_obaveze += 35000 * 6
                broj_pausalaca += 1
            elif klijent_red["Pravna Forma"] == "Udruženje":
                broj_udruzenja += 1
            else:
                iznos_d = klijent_red["Sirov_Prihod_2026"] * 0.40 * 0.15
                trenutne_obaveze += iznos_d
                detalji_trenutnog.append(f"D.O.O. ({klijent_red['Firma']}): {iznos_d:,.2f} RSD")
        
        if broj_pausalaca > 0: detalji_trenutnog.append(f"{broj_pausalaca} Paušalca: {broj_pausalaca * 35000 * 6:,.2f} RSD")
        if broj_udruzenja > 0: detalji_trenutnog.append(f"{broj_udruzenja} Udruženje: 0.00 RSD (Izuzeto)")
        
        potencijalni_pdv = zbir_12m * 0.20
        potencijalni_porez_na_dobit = zbir_2026 * 0.40 * 0.15
        ukupni_konsolidovani_trosak = potencijalni_pdv + potencijalni_porez_na_dobit
        
        c1, c2 = st.columns(2)
        with c1: st.info(f"** TRENUTNE UKUPNE OBAVEZE GRUPE:**\n* **Ukupno plaćeno državi:** ~{trenutne_obaveze:,.2f} RSD\n* *Struktura:* {' + '.join(detalji_trenutnog)}")
        with c2: st.warning(f"**🚨 POTENCIJALNE OBAVEZE (Kao jedno D.O.O. lice):**\n* **Zbirni PDV (20%):** ~{potencijalni_pdv:,.2f} RSD\n* **Porez na dobit (15%):** ~{potencijalni_porez_na_dobit:,.2f} RSD\n* *Razlika u trošku:* ~{ukupni_konsolidovani_trosak - trenutne_obaveze:,.2f} RSD više obaveza!")
    else: st.warning("Nema firmi u tabeli koje odgovaraju ovom nazivu.")
    st.markdown('</div>', unsafe_allow_html=True)
    
    st.markdown('<div class="profile-box">', unsafe_allow_html=True)
    izabrana_firma = st.selectbox("Izaberite firmu za detaljan pregled istorije:", sorted(rezultati_df["Firma"].unique()))
    if izabrana_firma:
        # POPRAVLJENO: Dinamički povlačimo trenutni status firme iz gornje tabele i lepimo ga u naslov
        trenutni_status_firme = rezultati_df[rezultati_df["Firma"] == izabrana_firma]["Status"].values[0]
        st.subheader(f"📈 Prihod po mesecima za firmu {izabrana_firma} (Trenutni status: {trenutni_status_firme})")
        
        istorija_firme = df_sve_mesecno[df_sve_mesecno["Naziv"] == izabrana_firma].sort_values(by="Datum")
        istorija_prikaz = istorija_firme[["Mesec_Tekst", "Prihod"]].copy()
        istorija_prikaz["Prihod"] = istorija_prikaz["Prihod"].apply(lambda x: f"{x:,.2f} RSD")
        istorija_prikaz.columns = ["Mesec / Godina", "Ostvareni Prihod"]
        col_tabela, col_grafik = st.columns(2)
        with col_tabela:
            st.dataframe(istorija_prikaz, use_container_width=True, hide_index=True, height=250)
        with col_grafik:
            st.plotly_chart(px.line(istorija_firme, x="Datum", y="Prihod", title=f"Kretanje prihoda kroz vreme", labels={"Prihod": "Prihod (RSD)", "Datum": "Vreme"}, template="plotly_white").update_traces(line_color="#0a2540", line_width=3, mode="lines+markers"), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)
except Exception as e: st.error(f"⚠️ Greška u strukturi podataka: {e}")
