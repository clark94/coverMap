import streamlit as st
import pandas as pd
import folium
from folium import plugins
from streamlit_folium import st_folium
from math import radians, sin, cos, sqrt, atan2
import io
import json
import hashlib
from datetime import datetime, date
import gspread
from google.oauth2.service_account import Credentials

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="CoverMap — Délégués Terrain",
    page_icon="🗺️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600&family=Syne:wght@700;800&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
[data-testid="stSidebar"] { background: #0f172a; border-right: 1px solid #1e293b; }
[data-testid="stSidebar"] * { color: #e2e8f0 !important; }
[data-testid="stSidebar"] label {
    color: #94a3b8 !important; font-size: 0.73rem;
    text-transform: uppercase; letter-spacing: 0.07em; font-weight: 500;
}
.main { background: #f8fafc; }
.hero {
    background: linear-gradient(135deg, #0f172a 0%, #1e3a5f 60%, #0e4166 100%);
    border-radius: 16px; padding: 1.8rem 2.2rem; margin-bottom: 1.4rem;
}
.hero-title { font-family:'Syne',sans-serif; font-size:1.9rem; font-weight:800;
    color:#f0f9ff; margin:0; letter-spacing:-0.02em; }
.hero-sub { color:#7dd3fc; font-size:0.88rem; margin-top:0.35rem; font-weight:300; }
.kpi-row { display:grid; grid-template-columns:repeat(4,1fr); gap:0.9rem; margin-bottom:1.4rem; }
.kpi { background:white; border-radius:12px; padding:1.1rem 1.3rem;
    border:1px solid #e2e8f0; box-shadow:0 1px 3px rgba(0,0,0,0.05); position:relative; overflow:hidden; }
.kpi::after { content:''; position:absolute; bottom:0; left:0; right:0; height:3px; }
.kpi.red::after   { background:#ef4444; }
.kpi.blue::after  { background:#3b82f6; }
.kpi.green::after { background:#10b981; }
.kpi.amber::after { background:#f59e0b; }
.kpi-v { font-family:'Syne',sans-serif; font-size:1.9rem; font-weight:800; color:#0f172a; line-height:1; }
.kpi-l { font-size:0.72rem; color:#64748b; margin-top:0.25rem; text-transform:uppercase; letter-spacing:.06em; font-weight:500; }
.kpi-s { font-size:0.68rem; color:#94a3b8; margin-top:0.1rem; }
.form-card {
    background: white; border: 1px solid #e2e8f0; border-radius: 16px;
    padding: 1.8rem; box-shadow: 0 2px 8px rgba(0,0,0,0.06);
}
.login-box {
    max-width: 420px; margin: 4rem auto;
    background: white; border-radius: 20px;
    padding: 2.5rem; box-shadow: 0 8px 32px rgba(0,0,0,0.10);
    border: 1px solid #e2e8f0;
}
.login-title {
    font-family:'Syne',sans-serif; font-size:1.6rem; font-weight:800;
    color:#0f172a; margin:0 0 0.3rem; text-align:center;
}
.login-sub { color:#64748b; font-size:0.85rem; text-align:center; margin-bottom:2rem; }
.badge-admin { background:#fef3c7; color:#92400e; padding:2px 10px; border-radius:999px; font-size:0.72rem; font-weight:600; }
.badge-delegue { background:#dbeafe; color:#1e40af; padding:2px 10px; border-radius:999px; font-size:0.72rem; font-weight:600; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────
RAYON_KM = 30

def haversine(lat1, lon1, lat2, lon2):
    R = 6371
    dlat = radians(lat2 - lat1)
    dlon = radians(lon2 - lon1)
    a = sin(dlat/2)**2 + cos(radians(lat1)) * cos(radians(lat2)) * sin(dlon/2)**2
    return R * 2 * atan2(sqrt(a), sqrt(1 - a))

def hash_password(password: str) -> str:
    return hashlib.sha256(password.encode()).hexdigest()

# ─────────────────────────────────────────────
# CONNEXION GOOGLE SHEETS — AVEC GESTION D'ERREUR
# ─────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

@st.cache_resource
def get_gsheet_client():
    return None

def get_sheet(sheet_name: str):
    return None

    """Récupère un onglet Google Sheets — avec gestion d'erreur propre."""
    try:
        client = get_gsheet_client()
        spreadsheet_id = st.secrets.get("SPREADSHEET_ID", "")
        if not spreadsheet_id:
            st.error("❌ Secret 'SPREADSHEET_ID' manquant. Configurez-le dans Streamlit Cloud → Settings → Secrets.")
            st.stop()
        spreadsheet = client.open_by_key(spreadsheet_id)
        return spreadsheet.worksheet(sheet_name)
    except Exception as e:
        st.error(f"❌ Impossible d'accéder à l'onglet '{sheet_name}' : {e}")
        st.stop()

# ─────────────────────────────────────────────
# AUTHENTIFICATION
# ─────────────────────────────────────────────
@st.cache_data(ttl=120)
def load_utilisateurs():
    return pd.DataFrame([
        {"email": "admin@test.com", "password_hash": hash_password("924802"), "nom": "Clark", "role": "admin", "actif": "OUI"}
    ])

def verifier_login(email: str, password: str):
    if email.lower() == "admin@test.com" and password == "924802":
        return {
            "email": "admin@test.com",
            "nom": "Clark",
            "role": "admin",
            "actif": "OUI"
        }
    return None

# ─────────────────────────────────────────────
# PAGE DE LOGIN
# ─────────────────────────────────────────────
def page_login():
    st.markdown("""
    <div style='text-align:center;padding:2rem 0 0'>
        <p style='font-family:Syne,sans-serif;font-size:2.2rem;font-weight:800;color:#0f172a;margin:0'>🗺️ CoverMap</p>
        <p style='color:#64748b;font-size:0.9rem;margin:0.3rem 0 2rem'>Analyse terrain délégués</p>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_center, col_right = st.columns([1, 1.2, 1])
    with col_center:
        st.markdown("<div class='login-box'>", unsafe_allow_html=True)
        st.markdown("<p class='login-title'>Connexion</p>", unsafe_allow_html=True)
        st.markdown("<p class='login-sub'>Entrez vos identifiants pour accéder à l'app</p>", unsafe_allow_html=True)

        email    = st.text_input("📧 Email", placeholder="votre.email@entreprise.com", key="login_email")
        password = st.text_input("🔒 Mot de passe", type="password", placeholder="••••••••", key="login_pwd")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("Se connecter →", use_container_width=True, type="primary"):
            if not email or not password:
                st.error("Veuillez remplir tous les champs.")
            else:
                user = verifier_login(email, password)
                if user:
                    st.session_state["user"] = user
                    st.session_state["logged_in"] = True
                    st.rerun()
                else:
                    st.error("❌ Email ou mot de passe incorrect, ou compte désactivé.")

        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("""
    <p style='text-align:center;color:#94a3b8;font-size:0.75rem;margin-top:1.5rem'>
        Problème de connexion ? Contactez votre administrateur.
    </p>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CHARGEMENT DONNÉES
# ─────────────────────────────────────────────
@st.cache_data(ttl=300)
def load_data():
    delegues = pd.DataFrame({
        "nom": ["Clark"],
        "region": ["Paris"],
        "region_group": ["IDF"],
        "latitude": [48.8566],
        "longitude": [2.3522],
        "nb_visites": [10],
    })

    med_france = pd.DataFrame({
        "nom": ["Martin", "Diallo", "Bernard"],
        "prenom": ["Jean", "Awa", "Luc"],
        "specialite_libelle": ["Médecin généraliste", "Cardiologue", "Dermatologue"],
        "commune": ["Paris", "Boulogne-Billancourt", "Saint-Denis"],
        "code_postal": ["75015", "92100", "93200"],
        "latitude": [48.8500, 48.8397, 48.9362],
        "longitude": [2.3000, 2.2399, 2.3574],
    })

    med_reseau_raw = pd.DataFrame({
        "nom": ["Martin", "Awa"],
        "commune": ["Paris", "Boulogne-Billancourt"],
        "specialite_libelle": ["Médecin généraliste", "Cardiologue"],
        "latitude": [48.8500, 48.8397],
        "longitude": [2.3000, 2.2399],
    })

    return delegues, med_france, med_reseau_raw

@st.cache_data(ttl=60)
def load_visites():
    if "visites_locales" not in st.session_state:
        st.session_state["visites_locales"] = []
    data = st.session_state["visites_locales"]
    return pd.DataFrame(data) if data else pd.DataFrame(columns=[
        "id","date","delegue","medecin","specialite","ville",
        "type_visite","affiche_deposee","presentoir","commentaire",
        "latitude","longitude","saisie_le"
    ])

def save_visite(row_dict: dict):
    if "visites_locales" not in st.session_state:
        st.session_state["visites_locales"] = []
    row_dict = row_dict.copy()
    row_dict["id"] = len(st.session_state["visites_locales"]) + 1
    row_dict["saisie_le"] = datetime.now().strftime("%Y-%m-%d %H:%M")
    st.session_state["visites_locales"].append(row_dict)
    load_visites.clear()

@st.cache_data
def calc_couverture(delegues_json, med_france):
    delegues = pd.read_json(io.StringIO(delegues_json))
    med = med_france.copy()
    med["couvert"] = False
    med["delegue_zone"] = ""
    for _, d in delegues.iterrows():
        mask = med.apply(
            lambda r: haversine(d.latitude, d.longitude, r.latitude, r.longitude) <= RAYON_KM, axis=1
        )
        med.loc[mask, "couvert"] = True
        med.loc[mask & (med["delegue_zone"] == ""), "delegue_zone"] = d["nom"]
    return med

def export_excel(delegues, med_couverts, med_non_couverts, med_reseau, visites):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        taux = round(len(med_couverts) / max(len(med_couverts)+len(med_non_couverts), 1) * 100, 1)
        pd.DataFrame({
            "Indicateur": ["Délégués","Médecins France","Couverts","Non couverts","Réseau IDF","Taux couverture","Visites saisies","Date"],
            "Valeur": [len(delegues), len(med_couverts)+len(med_non_couverts),
                       len(med_couverts), len(med_non_couverts), len(med_reseau),
                       f"{taux}%", len(visites), datetime.now().strftime("%d/%m/%Y %H:%M")]
        }).to_excel(writer, sheet_name="Résumé", index=False)
        delegues[["nom","region","region_group","nb_visites"]].rename(columns={
            "nom":"Délégué","region":"Région","region_group":"Groupe","nb_visites":"Visites"
        }).to_excel(writer, sheet_name="Délégués", index=False)
        cols = ["nom","prenom","specialite_libelle","commune","couvert","delegue_zone"]
        cols_ok = [c for c in cols if c in med_couverts.columns]
        med_couverts[cols_ok].to_excel(writer, sheet_name="Médecins couverts", index=False)
        cols_ok2 = [c for c in cols if c in med_non_couverts.columns]
        med_non_couverts[cols_ok2].to_excel(writer, sheet_name="Non couverts", index=False)
        if len(visites) > 0:
            visites.to_excel(writer, sheet_name="Visites terrain", index=False)
    buffer.seek(0)
    return buffer

# ─────────────────────────────────────────────
# PAGE ADMIN — GESTION UTILISATEURS
# ─────────────────────────────────────────────
def page_admin_users():
    st.markdown("### 👥 Gestion des utilisateurs")
    st.info("Mode test local : les comptes sont simulés dans l'application.")

    users = load_utilisateurs().copy()
    display = users[["email","nom","role","actif"]].copy()
    display.columns = ["Email","Nom","Rôle","Actif"]
    st.dataframe(display, use_container_width=True, hide_index=True)
    st.caption("Compte de test disponible : admin@test.com / 924802")

# ══════════════════════════════════════════════
# POINT D'ENTRÉE PRINCIPAL (SANS LOGIN)
# ══════════════════════════════════════════════

# Accès direct sans authentification
if "user" not in st.session_state:
    st.session_state["user"] = {
        "email": "public@covermap.com",
        "nom": "Clark",
        "role": "admin"
    }

if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = True

# ─────────────────────────────────────────────
# APP PRINCIPALE (utilisateur connecté)
# ─────────────────────────────────────────────
user     = st.session_state["user"]
is_admin = user.get("role", "delegue") == "admin"

try:
    delegues, med_france, med_reseau = load_data()
    visites_df = load_visites()
except Exception as e:
    st.error(f"❌ Erreur de chargement des données : {e}")
    st.stop()

# ── SIDEBAR ──
with st.sidebar:
    st.markdown(f"""
    <div style='padding: 1rem 0 0.5rem'>
        <p style='font-family:Syne,sans-serif;font-weight:800;font-size:1.1rem;color:#f0f9ff;margin:0'>🗺️ CoverMap</p>
        <p style='font-size:0.72rem;color:#64748b;margin:0.2rem 0 0.5rem'>Analyse terrain délégués</p>
        <div style='background:#1e293b;border-radius:8px;padding:0.6rem 0.8rem;margin-bottom:1rem'>
            <p style='font-size:0.78rem;color:#7dd3fc;margin:0;font-weight:600'>👤 {user["nom"]}</p>
            <p style='font-size:0.68rem;color:#475569;margin:0'>{user["email"]}</p>
            <span style='font-size:0.65rem;background:{"#fef3c7" if is_admin else "#dbeafe"};
                         color:{"#92400e" if is_admin else "#1e40af"};
                         padding:1px 7px;border-radius:999px;font-weight:600;'>
                {"⭐ Admin" if is_admin else "🏃 Délégué"}
            </span>
        </div>
    </div>""", unsafe_allow_html=True)

    st.markdown("**Couches carte**")
    show_delegues    = st.checkbox("🔴 Délégués",          value=True)
    show_france      = st.checkbox("🔵 Médecins France",   value=True)
    show_reseau      = st.checkbox("🟢 Réseau IDF",        value=True)
    show_zones       = st.checkbox("⭕ Zones 30 KM",       value=True)
    show_couverts    = st.checkbox("✅ Couverts",           value=True)
    show_noncouverts = st.checkbox("❌ Non couverts",       value=True)

    st.divider()
    st.markdown("**Groupe région**")
    groupes    = ["Tous"] + sorted(delegues["region_group"].dropna().unique().tolist())
    groupe_sel = st.selectbox("Groupe", groupes, label_visibility="collapsed")

    st.markdown("**Région**")
    if groupe_sel == "Tous":
        regions = ["Toutes"] + sorted(delegues["region"].dropna().unique().tolist())
    else:
        regions = ["Toutes"] + sorted(delegues[delegues["region_group"] == groupe_sel]["region"].dropna().unique().tolist())
    region_sel = st.selectbox("Région", regions, label_visibility="collapsed")

    st.divider()
    if st.button("🔄 Rafraîchir", use_container_width=True):
        load_data.clear()
        load_visites.clear()
        st.rerun()

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚪 Se déconnecter", use_container_width=True):
        st.session_state.clear()
        st.rerun()

# Filtrage délégués
del_f = delegues.copy()
if groupe_sel != "Tous":
    del_f = del_f[del_f["region_group"] == groupe_sel]
if region_sel != "Toutes":
    del_f = del_f[del_f["region"] == region_sel]

# Calcul couverture
med_avec_couverture = calc_couverture(del_f.to_json(), med_france)
med_couverts     = med_avec_couverture[med_avec_couverture["couvert"] == True].copy()
med_non_couverts = med_avec_couverture[med_avec_couverture["couvert"] == False].copy()
taux_couv        = round(len(med_couverts) / max(len(med_avec_couverture), 1) * 100, 1)

# ── HEADER ──
st.markdown(f"""
<div class="hero">
    <p class="hero-title">🗺️ CoverMap — Analyse Terrain Délégués</p>
    <p class="hero-sub">Carte · Saisie terrain · Tableau filtré · Export rapport</p>
</div>
<div class="kpi-row">
    <div class="kpi red">
        <div class="kpi-v">{len(del_f)}</div>
        <div class="kpi-l">Délégués</div>
        <div class="kpi-s">Sur le terrain</div>
    </div>
    <div class="kpi blue">
        <div class="kpi-v">{len(med_avec_couverture)}</div>
        <div class="kpi-l">Médecins France</div>
        <div class="kpi-s">Base complète</div>
    </div>
    <div class="kpi green">
        <div class="kpi-v">{len(med_couverts)}</div>
        <div class="kpi-l">Couverts</div>
        <div class="kpi-s">{taux_couv}% du total</div>
    </div>
    <div class="kpi amber">
        <div class="kpi-v">{len(visites_df)}</div>
        <div class="kpi-l">Visites saisies</div>
        <div class="kpi-s">Via formulaire terrain</div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── TABS ──
tabs_list = ["🗺️ Carte", "📝 Saisie terrain", "📋 Historique visites", "🔍 Tableau filtré", "📥 Export"]
if is_admin:
    tabs_list.append("⚙️ Utilisateurs")

all_tabs = st.tabs(tabs_list)
tab1, tab2, tab3, tab4, tab5 = all_tabs[:5]
tab_admin = all_tabs[5] if is_admin else None

# ══════════════════════════════════════════════
# TAB 1 — CARTE
# ══════════════════════════════════════════════
with tab1:
    col_map, col_leg = st.columns([4, 1])
    with col_map:
        center_lat = del_f["latitude"].mean() if len(del_f) > 0 else 46.6
        center_lon = del_f["longitude"].mean() if len(del_f) > 0 else 2.3
        m = folium.Map(location=[center_lat, center_lon], zoom_start=6, tiles="CartoDB positron")

        if show_zones:
            for _, d in del_f.iterrows():
                folium.Circle(
                    location=[d.latitude, d.longitude], radius=RAYON_KM*1000,
                    color="#ef4444", fill=True, fill_color="#fca5a5",
                    fill_opacity=0.08, weight=1.5, dash_array="6"
                ).add_to(m)

        if show_france and show_couverts:
            for _, med in med_couverts.iterrows():
                folium.CircleMarker(
                    location=[med.latitude, med.longitude],
                    radius=4, color="#3b82f6", fill=True,
                    fill_color="#3b82f6", fill_opacity=0.8, weight=1,
                    tooltip=folium.Tooltip(
                        f"🔵 {med.get('nom','?')} {med.get('prenom','')}<br>"
                        f"{med.get('specialite_libelle','—')} — {med.get('commune','—')}<br>"
                        f"<small>Couvert par : {med.get('delegue_zone','—')}</small>"
                    )
                ).add_to(m)

        if show_france and show_noncouverts:
            for _, med in med_non_couverts.iterrows():
                folium.CircleMarker(
                    location=[med.latitude, med.longitude],
                    radius=4, color="#ef4444", fill=True,
                    fill_color="#ef4444", fill_opacity=0.8, weight=1,
                    tooltip=folium.Tooltip(
                        f"🔴 {med.get('nom','?')} {med.get('prenom','')}<br>"
                        f"{med.get('specialite_libelle','—')} — {med.get('commune','—')}<br>"
                        f"<small>Hors zone — non couvert</small>"
                    )
                ).add_to(m)

        if show_reseau:
            for _, med in med_reseau.iterrows():
                folium.CircleMarker(
                    location=[med.latitude, med.longitude],
                    radius=5, color="#10b981", fill=True,
                    fill_color="#10b981", fill_opacity=0.85, weight=1.5,
                    tooltip=folium.Tooltip(
                        f"🟢 {med.get('nom','—')}<br>"
                        f"{med.get('specialite_libelle','—')} — {med.get('commune','—')}"
                    )
                ).add_to(m)

        if show_delegues:
            for _, d in del_f.iterrows():
                folium.Marker(
                    location=[d.latitude, d.longitude],
                    tooltip=folium.Tooltip(f"🔴 {d['nom']} — {d['region']}"),
                    popup=folium.Popup(
                        f"<b>🔴 {d['nom']}</b><br>{d['region']} — {d['region_group']}<br>"
                        f"<hr style='margin:4px 0'>Visites : <b>{d['nb_visites']}</b>",
                        max_width=200
                    ),
                    icon=folium.Icon(color="red", icon="user", prefix="fa")
                ).add_to(m)

        plugins.Fullscreen().add_to(m)
        st_folium(m, height=580, use_container_width=True)

    with col_leg:
        st.markdown(f"""
        <div style='background:white;border:1px solid #e2e8f0;border-radius:12px;padding:1rem;font-size:.8rem;'>
            <p style='font-family:Syne,sans-serif;font-weight:700;font-size:.85rem;color:#0f172a;margin:0 0 1rem'>Légende</p>
            <div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.4rem'>
                <div style='width:13px;height:13px;border-radius:50%;background:#ef4444;flex-shrink:0'></div><span>Délégué</span>
            </div>
            <div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.9rem'>
                <div style='width:18px;height:18px;border-radius:50%;background:#fca5a5;border:1.5px dashed #ef4444;flex-shrink:0'></div>
                <span style='color:#64748b'>Zone 30 km</span>
            </div>
            <div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.4rem'>
                <div style='width:11px;height:11px;border-radius:50%;background:#3b82f6;flex-shrink:0'></div>
                <span style='color:#64748b'>Couvert</span>
            </div>
            <div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.4rem'>
                <div style='width:11px;height:11px;border-radius:50%;background:#ef4444;flex-shrink:0'></div>
                <span style='color:#64748b'>Non couvert</span>
            </div>
            <div style='display:flex;align-items:center;gap:.5rem;margin-bottom:.9rem'>
                <div style='width:11px;height:11px;border-radius:50%;background:#10b981;flex-shrink:0'></div>
                <span style='color:#64748b'>Réseau IDF</span>
            </div>
            <hr style='margin:.8rem 0;border-color:#e2e8f0'>
            <p style='color:#64748b;font-size:.68rem;margin:0 0 .2rem'><b style='color:#0f172a'>{len(med_couverts)}</b> couverts</p>
            <p style='color:#64748b;font-size:.68rem;margin:0 0 .2rem'><b style='color:#0f172a'>{len(med_non_couverts)}</b> non couverts</p>
            <p style='color:#64748b;font-size:.68rem;margin:0'><b style='color:#0f172a'>{taux_couv}%</b> taux couverture</p>
        </div>""", unsafe_allow_html=True)

# ══════════════════════════════════════════════
# TAB 2 — SAISIE TERRAIN
# ══════════════════════════════════════════════
with tab2:
    st.markdown("### 📝 Saisie d'une visite terrain")
    st.markdown(f"Connecté en tant que **{user['nom']}** — votre nom sera automatiquement enregistré.")

    with st.container():
        st.markdown("<div class='form-card'>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if is_admin:
                delegue_sel = st.selectbox("👤 Délégué", sorted(delegues["nom"].tolist()), key="saisie_del")
            else:
                delegue_sel = user["nom"]
                st.text_input("👤 Délégué", value=delegue_sel, disabled=True, key="saisie_del")
            date_visite = st.date_input("📅 Date de la visite", value=date.today(), key="saisie_date")
            type_visite = st.selectbox("🏷️ Type de visite", [
                "Visite standard","OPE spéciale","Formation","Remise documentation","Autre"
            ], key="saisie_type")

        with c2:
            medecin_nom    = st.text_input("👨‍⚕️ Nom du médecin", placeholder="Ex: Dr. Martin", key="saisie_med")
            specialite_sel = st.selectbox("🩺 Spécialité", [
                "Médecin généraliste","Cardiologue","Dermatologue","Pneumologue",
                "Rhumatologue","Neurologue","Gastro-entérologue","Autre"
            ], key="saisie_spec")
            ville_sel = st.text_input("🏙️ Ville", placeholder="Ex: Paris", key="saisie_ville")

        c3, c4 = st.columns(2)
        with c3:
            affiche    = st.selectbox("📌 Affiche déposée ?", ["✅ Oui","❌ Non"], key="saisie_aff")
        with c4:
            presentoir = st.selectbox("🗂️ Présentoir installé ?", ["✅ Oui","❌ Non"], key="saisie_pres")

        commentaire = st.text_area("💬 Commentaire", placeholder="Notes sur la visite...", key="saisie_com", height=80)
        st.markdown("</div>", unsafe_allow_html=True)

        col_btn, _ = st.columns([1, 3])
        with col_btn:
            if st.button("✅ Enregistrer la visite", use_container_width=True, type="primary"):
                if not medecin_nom:
                    st.warning("Veuillez renseigner le nom du médecin.")
                else:
                    del_info = delegues[delegues["nom"] == delegue_sel]
                    lat = float(del_info["latitude"].values[0]) if len(del_info) > 0 else 0
                    lon = float(del_info["longitude"].values[0]) if len(del_info) > 0 else 0
                    save_visite({
                        "date":            str(date_visite),
                        "delegue":         delegue_sel,
                        "medecin":         medecin_nom,
                        "specialite":      specialite_sel,
                        "ville":           ville_sel,
                        "type_visite":     type_visite,
                        "affiche_deposee": affiche,
                        "presentoir":      presentoir,
                        "commentaire":     commentaire,
                        "latitude":        lat,
                        "longitude":       lon,
                    })
                    st.success("✅ Visite enregistrée dans Google Sheets !")
                    st.rerun()

# ══════════════════════════════════════════════
# TAB 3 — HISTORIQUE VISITES
# ══════════════════════════════════════════════
with tab3:
    st.markdown("### 📋 Historique des visites terrain")
    visites_df = load_visites()

    if len(visites_df) == 0:
        st.info("Aucune visite saisie pour l'instant.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            if is_admin:
                del_hist = st.selectbox("Délégué", ["Tous"] + sorted(visites_df["delegue"].unique().tolist()), key="del_hist")
            else:
                del_hist = user["nom"]
                st.text_input("Délégué", value=del_hist, disabled=True, key="del_hist_disp")
        with c2:
            type_hist = st.selectbox("Type visite", ["Tous"] + sorted(visites_df["type_visite"].dropna().unique().tolist()), key="type_hist")
        with c3:
            aff_hist = st.selectbox("Affiche déposée", ["Tous","✅ Oui","❌ Non"], key="aff_hist")

        vis_filtre = visites_df.copy()
        if not is_admin:
            vis_filtre = vis_filtre[vis_filtre["delegue"] == user["nom"]]
        elif del_hist != "Tous":
            vis_filtre = vis_filtre[vis_filtre["delegue"] == del_hist]
        if type_hist != "Tous": vis_filtre = vis_filtre[vis_filtre["type_visite"]     == type_hist]
        if aff_hist  != "Tous": vis_filtre = vis_filtre[vis_filtre["affiche_deposee"] == aff_hist]

        st.markdown(f"**{len(vis_filtre)} visite(s) trouvée(s)**")
        st.divider()

        k1, k2, k3, k4 = st.columns(4)
        k1.metric("Total visites",     len(vis_filtre))
        k2.metric("Affiches déposées", len(vis_filtre[vis_filtre["affiche_deposee"]=="✅ Oui"]))
        k3.metric("Opé spéciales",     len(vis_filtre[vis_filtre["type_visite"].str.contains("OPE", na=False)]))
        k4.metric("Délégués actifs",   vis_filtre["delegue"].nunique())

        st.divider()
        cols_aff = ["date","delegue","medecin","specialite","ville","type_visite","affiche_deposee","presentoir","commentaire"]
        cols_ok  = [c for c in cols_aff if c in vis_filtre.columns]
        st.dataframe(
            vis_filtre[cols_ok].rename(columns={
                "date":"Date","delegue":"Délégué","medecin":"Médecin",
                "specialite":"Spécialité","ville":"Ville","type_visite":"Type",
                "affiche_deposee":"Affiche","presentoir":"Présentoir","commentaire":"Commentaire"
            }),
            use_container_width=True, hide_index=True, height=380
        )

        buf = io.BytesIO()
        vis_filtre.to_excel(buf, index=False, engine="openpyxl")
        buf.seek(0)
        st.download_button(
            label="⬇️ Exporter en Excel",
            data=buf,
            file_name=f"visites_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ══════════════════════════════════════════════
# TAB 4 — TABLEAU FILTRÉ
# ══════════════════════════════════════════════
with tab4:
    st.markdown("### 🔍 Tableau filtré — Médecins")
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        statut_f = st.selectbox("Statut", ["Tous","✅ Couverts","❌ Non couverts"])
    with c2:
        specs_col = med_avec_couverture["specialite_libelle"].dropna().unique().tolist() if "specialite_libelle" in med_avec_couverture.columns else []
        spec_f = st.selectbox("Spécialité", ["Toutes"] + sorted(specs_col))
    with c3:
        communes_col = med_avec_couverture["commune"].dropna().unique().tolist() if "commune" in med_avec_couverture.columns else []
        commune_f = st.selectbox("Commune", ["Toutes"] + sorted(communes_col))
    with c4:
        del_zones_col = med_couverts["delegue_zone"].dropna().unique().tolist() if "delegue_zone" in med_couverts.columns else []
        del_zone_f = st.selectbox("Zone délégué", ["Tous"] + sorted(del_zones_col))

    recherche = st.text_input("🔎 Rechercher un médecin", placeholder="Nom, prénom...")

    med_filtre = med_avec_couverture.copy()
    if statut_f == "✅ Couverts":       med_filtre = med_filtre[med_filtre["couvert"] == True]
    elif statut_f == "❌ Non couverts": med_filtre = med_filtre[med_filtre["couvert"] == False]
    if spec_f    != "Toutes"  and "specialite_libelle" in med_filtre.columns:
        med_filtre = med_filtre[med_filtre["specialite_libelle"] == spec_f]
    if commune_f != "Toutes"  and "commune" in med_filtre.columns:
        med_filtre = med_filtre[med_filtre["commune"] == commune_f]
    if del_zone_f != "Tous"   and "delegue_zone" in med_filtre.columns:
        med_filtre = med_filtre[med_filtre["delegue_zone"] == del_zone_f]
    if recherche:
        mask = med_filtre.apply(
            lambda r: recherche.lower() in str(r.get("nom","")).lower() or
                      recherche.lower() in str(r.get("prenom","")).lower(), axis=1)
        med_filtre = med_filtre[mask]

    st.markdown(f"**{len(med_filtre)} médecin(s) trouvé(s)**")
    st.divider()
    cols_aff = ["nom","prenom","specialite_libelle","commune","code_postal","couvert","delegue_zone"]
    cols_ok  = [c for c in cols_aff if c in med_filtre.columns]
    affich   = med_filtre[cols_ok].copy()
    if "couvert" in affich.columns:
        affich["couvert"] = affich["couvert"].map({True:"✅ Couvert", False:"❌ Non couvert"})
    affich = affich.rename(columns={
        "nom":"Nom","prenom":"Prénom","specialite_libelle":"Spécialité",
        "commune":"Commune","code_postal":"Code postal",
        "couvert":"Statut","delegue_zone":"Zone délégué"
    })
    st.dataframe(affich, use_container_width=True, hide_index=True, height=400)

    buf2 = io.BytesIO()
    med_filtre.to_excel(buf2, index=False, engine="openpyxl")
    buf2.seek(0)
    st.download_button(
        label=f"⬇️ Exporter ({len(med_filtre)} lignes) en Excel",
        data=buf2,
        file_name=f"medecins_filtre_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

# ══════════════════════════════════════════════
# TAB 5 — EXPORT
# ══════════════════════════════════════════════
with tab5:
    st.markdown("### 📥 Export complet")
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("""
        <div style='background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:1.2rem;'>
            <p style='font-weight:600;color:#166534;margin:0 0 .4rem'>📊 Excel complet</p>
            <p style='font-size:.82rem;color:#166534;margin:0'>Résumé · Délégués · Couverts · Non couverts · Visites terrain</p>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        excel_data = export_excel(del_f, med_couverts, med_non_couverts, med_reseau, visites_df)
        st.download_button(
            label="⬇️ Télécharger Excel",
            data=excel_data,
            file_name=f"covermap_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_b:
        st.markdown("""
        <div style='background:#eff6ff;border:1px solid #93c5fd;border-radius:12px;padding:1.2rem;'>
            <p style='font-weight:600;color:#1e40af;margin:0 0 .4rem'>🗺️ Carte HTML</p>
            <p style='font-size:.82rem;color:#1e40af;margin:0'>Carte interactive · Ctrl+P pour PDF</p>
        </div>""", unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)
        m_exp = folium.Map(location=[46.6, 2.3], zoom_start=6, tiles="CartoDB positron")
        for _, d in del_f.iterrows():
            folium.Circle(location=[d.latitude, d.longitude], radius=RAYON_KM*1000,
                color="#ef4444", fill=True, fill_color="#fca5a5", fill_opacity=0.1, weight=2).add_to(m_exp)
            folium.Marker(location=[d.latitude, d.longitude], tooltip=f"🔴 {d['nom']}",
                icon=folium.Icon(color="red", icon="user", prefix="fa")).add_to(m_exp)
        carte_html = f"""<!DOCTYPE html><html><head><meta charset='UTF-8'>
        <title>CoverMap</title></head><body style='margin:0'>{m_exp._repr_html_()}</body></html>"""
        st.download_button(
            label="⬇️ Télécharger la carte HTML",
            data=carte_html.encode("utf-8"),
            file_name=f"carte_{datetime.now().strftime('%Y%m%d_%H%M')}.html",
            mime="text/html",
            use_container_width=True
        )

    st.divider()
    st.markdown("### 📋 Résumé")
    st.dataframe(pd.DataFrame({
        "Indicateur": ["Délégués","Médecins France","Couverts","Non couverts","Réseau IDF","Visites saisies","Taux couverture"],
        "Valeur": [len(del_f), len(med_avec_couverture), len(med_couverts),
                   len(med_non_couverts), len(med_reseau), len(visites_df), f"{taux_couv}%"]
    }), use_container_width=True, hide_index=True)

# ══════════════════════════════════════════════
# TAB ADMIN — GESTION UTILISATEURS
# ══════════════════════════════════════════════
if is_admin and tab_admin:
    with tab_admin:
        page_admin_users()
