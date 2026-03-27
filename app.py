import streamlit as st
from docxtpl import DocxTemplate, RichText
import json
import io

# Configuration de la page pour que ça prenne bien l'écran de l'iPad
st.set_page_config(page_title="App PV Maulini", layout="wide")

st.title("⚡ Assistant PV de Séance - Maulini")

# --- 1. LE CERVEAU : IMPORT DU JSON (Mémoire de la semaine précédente) ---
with st.sidebar:
    st.header("📂 1. Préparer la séance")
    st.info("Glisse ici le petit fichier JSON de la semaine dernière.")
    fichier_json = st.file_uploader("Importer l'historique", type=["json"])
    
    # Si on charge un fichier, on passe les anciens points bleus en noir
    if fichier_json is not None and 'import_fait' not in st.session_state:
        historique = json.load(fichier_json)
        for p in historique:
            p['color'] = "#000000" # Force le noir pour les anciens points
        st.session_state.points = historique
        st.session_state.import_fait = True
        st.success("Ancien PV chargé. Les nouveautés de la semaine dernière sont passées en noir.")
        st.rerun()

# --- INITIALISATION DE BASE (Pour que tu aies tes points de mardi prêt à l'emploi) ---
if 'points' not in st.session_state:
    st.session_state.points = [
        {"qui": "Weibel SA", "quoi": "- Intervention pour pose du dernier étayage le 30.03 (5 jours)", "quand": "30.03.26", "color": "#000000"},
        {"qui": "Piasio SA", "quoi": "- Transmettre demande occupation domaine public\n- Déposer flyers", "quand": "OK", "color": "#000000"},
        {"qui": "Maulini SA", "quoi": "- Montage de la grue le 14.04", "quand": "14.04.26", "color": "#000000"}
    ]

# --- 2. L'ACTION : AJOUT EN SÉANCE ---
st.subheader("➕ 2. Ajouter un point (Séance en cours)")
with st.form("ajout"):
    col1, col2, col3 = st.columns([1, 2, 1])
    # La liste des entreprises habituelles pour gagner du temps
    qui = col1.selectbox("Entreprise", ["Maulini SA", "Piasio SA", "Weibel SA", "Cerutti sanitaires", "Kw services SA", "EDMS SA", "KARAKAS"])
    quoi = col2.text_area("Remarque (Utilise la dictée de ton mobile !)")
    quand = col3.text_input("Délai", placeholder="Ex: ASAP")
    
    if st.form_submit_button("Ajouter ce point (S'affichera en Bleu)"):
        st.session_state.points.append({"qui": qui, "quoi": quoi, "quand": quand, "color": "#0000FF"})
        st.rerun()

# --- 3. LE CONTRÔLE : GESTION DES POINTS ---
st.write("---")
st.subheader("📋 3. Ordre du jour actuel")
for i, p in enumerate(st.session_state.points):
    colA, colB = st.columns([11, 1])
    # Petit indicateur visuel sur l'app
    pastille = "🔵" if p['color'] == "#0000FF" else "⚫"
    colA.write(f"{pastille} **{p['qui']}** : {p['quoi']} *(Pour le: {p['quand']})*")
    
    # Le bouton pour rayer un point quand l'entreprise a fait le job
    if colB.button("❌", key=f"del_{i}", help="Supprimer définitivement ce point"):
        st.session_state.points.pop(i)
        st.rerun()

# --- 4. LA FINALISATION : EXPORTS (Fin de séance) ---
st.write("---")
st.subheader("🚀 4. Clôturer la séance")
st.write("Télécharge ces deux fichiers à la fin de la réunion : le Word pour M-Files, et le JSON pour la semaine prochaine.")

col_word, col_json = st.columns(2)

# --- BOUTON 1 : Génération du Word ---
try:
    doc = DocxTemplate("template_seance.docx")
    points_formates = []
    
    # On injecte les couleurs dans le document
    for p in st.session_state.points:
        rt = RichText()
        rt.add(p['quoi'], color=p['color'], bold=(p['color'] == "#000000"))
        points_formates.append({"qui": p['qui'], "quoi": rt, "quand": p['quand']})

    doc.render({"points": points_formates})
    bio_word = io.BytesIO()
    doc.save(bio_word)

    col_word.download_button(
        label="📥 1. Télécharger le PV (Word)", 
        data=bio_word.getvalue(), 
        file_name="PV_Seance_Maulini.docx", 
        type="primary"
    )
except Exception as e:
    st.error(f"⚠️ Erreur avec le Word. Vérifie que 'template_seance.docx' est bien là. ({e})")

# --- BOUTON 2 : Génération de la mémoire (JSON) ---
json_data = json.dumps(st.session_state.points, indent=4)
col_json.download_button(
    label="💾 2. Sauvegarder l'historique (JSON)", 
    data=json_data, 
    file_name="historique_pv.json"
)
