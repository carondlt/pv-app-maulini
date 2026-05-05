import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import json
import io
import datetime

# --- FONCTION UTILITAIRE POUR LE DESIGN DU WORD ---
def set_cell_background(cell, fill_color):
    """Permet de colorer le fond d'une cellule Word avec un code HEX"""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), fill_color)
    tcPr.append(shd)

# --- CONFIGURATION DE L'APPLICATION ---
st.set_page_config(page_title="App PV Universel", layout="wide")
st.title("⚡ Assistant PV de Séance")

# --- 1. PARAMÈTRES GÉNÉRAUX & MÉMOIRE ---
with st.sidebar:
    st.header("📂 1. Infos du projet & Historique")
    
    # Informations de la séance
    nom_projet = st.text_input("Nom du Projet", "Projet Alpha")
    numero_pv = st.number_input("Numéro du PV", min_value=1, value=1, step=1)
    date_seance = st.date_input("Date de la séance", datetime.date.today())
    lieu = st.text_input("Lieu", "Genève")
    
    # --- NOUVEAU : CHOIX DU RÉDACTEUR ---
    choix_redacteur = st.selectbox("Établi par", ["Caroline", "Autre..."])
    if choix_redacteur == "Autre...":
        redacteur = st.text_input("Précise le nom du rédacteur")
    else:
        redacteur = choix_redacteur
    # ------------------------------------
    
    prochaine_seance = st.text_input("Date prochaine séance", "La semaine prochaine")
    
    st.write("---")
    st.info("Glisse ici le fichier JSON de la séance précédente.")
    fichier_json = st.file_uploader("Importer l'historique", type=["json"])
    
    if fichier_json is not None and 'import_fait' not in st.session_state:
        historique = json.load(fichier_json)
        for p in historique:
            p['color'] = "#000000" # Les anciens points passent en noir
        st.session_state.points = historique
        st.session_state.import_fait = True
        st.success("Ancien PV chargé avec succès.")
        st.rerun()

if 'points' not in st.session_state:
    st.session_state.points = []

# --- 2. PRÉSENCES ---
st.subheader("👥 2. Liste des intervenants")
col_pres, col_exc, col_abs = st.columns(3)
presents = col_pres.text_area("Présents (un par ligne)", "Caroline\nReprésentant Weibel SA")
excuses = col_exc.text_area("Excusés (un par ligne)", "Architecte")
absents = col_abs.text_area("Absents (non excusés)", "")

# --- 3. L'ACTION : AJOUT DES POINTS ---
st.write("---")
st.subheader("➕ 3. Ajouter un point à l'ordre du jour")

intervenants_existants = list(set([p["qui"] for p in st.session_state.points]))
intervenants_existants = ["Nouveau..."] + sorted(intervenants_existants) if intervenants_existants else ["Nouveau...", "Maulini SA", "Weibel SA"]

with st.form("ajout"):
    col1, col2, col3 = st.columns([1, 2, 1])
    qui_existant = col1.selectbox("Responsable", intervenants_existants)
    qui_nouveau = col1.text_input("Ou nouveau responsable")
    
    qui = qui_nouveau if qui_existant == "Nouveau..." and qui_nouveau else (qui_existant if qui_existant != "Nouveau..." else "Général")
    quoi = col2.text_area("Action ou remarque")
    quand = col3.text_input("Délai", placeholder="Ex: 30.03.26")
    
    if st.form_submit_button("Ajouter (Nouveau point = Bleu)"):
        st.session_state.points.append({"qui": qui, "quoi": quoi, "quand": quand, "color": "#0000FF"})
        st.rerun()

# --- 4. LE CONTRÔLE ---
st.write("---")
st.subheader(f"📋 4. Liste des points en cours")
for i, p in enumerate(st.session_state.points):
    colA, colB = st.columns([11, 1])
    pastille = "🔵" if p['color'] == "#0000FF" else "⚫"
    colA.write(f"{pastille} **{p['qui']}** : {p['quoi']} *(Délai: {p['quand']})*")
    if colB.button("❌", key=f"del_{i}", help="Solder ce point"):
        st.session_state.points.pop(i)
        st.rerun()

# --- 5. EXPORTATION ---
st.write("---")
st.subheader("🚀 5. Clôturer et Exporter")
col_word, col_json = st.columns(2)

# ----- GÉNÉRATION DU DOCUMENT WORD ESTHÉTIQUE -----
doc = Document()

# 1. Configuration de la police globale (Gotham)
style_normal = doc.styles['Normal']
font = style_normal.font
font.name = 'Gotham'
font.size = Pt(11)

# 2. Titre Principal avec le numéro du PV
titre = doc.add_heading(level=0)
titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_titre = titre.add_run(f"PROCÈS-VERBAL DE SÉANCE N°{numero_pv}")
run_titre.font.name = 'Gotham'
run_titre.font.size = Pt(18)
run_titre.font.bold = True
run_titre.font.color.rgb = RGBColor(44, 62, 80)

doc.add_paragraph() # Espace

# 3. En-tête informatif
table_info = doc.add_table(rows=3, cols=2)
table_info.autofit = True
# Ligne 1
table_info.cell(0, 0).text = f"Projet : {nom_projet}"
table_info.cell(0, 1).text = f"Date : {date_seance.strftime('%d.%m.%Y')}"
# Ligne 2
table_info.cell(1, 0).text = f"Lieu : {lieu}"
# Utilisation de la variable 'redacteur' qui vient du menu déroulant ou du champ texte
table_info.cell(1, 1).text = f"Établi par : {redacteur}" 
# Ligne 3
table_info.cell(2, 0).text = f"Prochaine séance : {prochaine_seance}"

for row in table_info.rows:
    for cell in row.cells:
        for paragraph in cell.paragraphs:
            if ":" in paragraph.text:
                texte_split = paragraph.text.split(":")
                paragraph.text = ""
                run_bold = paragraph.add_run(texte_split[0] + " :")
                run_bold.bold = True
                paragraph.add_run(texte_split[1])

doc.add_paragraph() # Espace
doc.add_paragraph("---") # Ligne de séparation visuelle

# 4. Liste des présences, excusés et absents
def ajouter_liste_intervenants(document, titre_liste, contenu):
    p = document.add_paragraph()
    p.add_run(f"{titre_liste} : ").bold = True
    texte_propre = contenu.strip()
    if texte_propre:
        p.add_run(texte_propre.replace('\n', ', '))
    else:
        p.add_run("Aucun")

ajouter_liste_intervenants(doc, "Présents", presents)
ajouter_liste_intervenants(doc, "Excusés", excuses)
ajouter_liste_intervenants(doc, "Absents", absents)

doc.add_paragraph() # Espace

# 5. Le grand tableau des points
doc.add_heading("Suivi des Actions", level=1)
table_actions = doc.add_table(rows=1, cols=3)
table_actions.style = 'Table Grid'

# En-têtes
hdr_cells = table_actions.rows[0].cells
en_tetes = ['Responsable', 'Action / Remarque', 'Délai']

for i, text in enumerate(en_tetes):
    hdr_cells[i].text = text
    set_cell_background(hdr_cells[i], 'EAEAEA') 
    paragraph = hdr_cells[i].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.runs[0]
    run.font.bold = True
    run.font.name = 'Gotham'

# Remplissage des tâches
for p in st.session_state.points:
    row_cells = table_actions.add_row().cells
    
    # Responsable
    row_cells[0].text = p['qui']
    
    # Action avec gestion de la couleur
    p_quoi = row_cells[1].paragraphs[0]
    run_action = p_quoi.add_run(p['quoi'])
    run_action.font.name = 'Gotham'
    
    if p['color'] == "#0000FF":
        run_action.font.color.rgb = RGBColor(0, 102, 204) # Bleu nouveau point
    else:
        run_action.font.color.rgb = RGBColor(0, 0, 0)
        run_action.font.bold = True # Ancien point en noir et gras
        
    # Délai
    p_quand = row_cells[2].paragraphs[0]
    p_quand.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_quand.add_run(p['quand']).font.name = 'Gotham'

# Sauvegarde Word
bio_word = io.BytesIO()
doc.save(bio_word)
nom_fichier_word = f"PV_n{numero_pv}_{nom_projet.replace(' ', '_')}_{date_seance.strftime('%Y%m%d')}.docx"

col_word.download_button(
    label="📥 1. Télécharger le PV (Word)", 
    data=bio_word.getvalue(), 
    file_name=nom_fichier_word, 
    type="primary"
)

# ----- SAUVEGARDE JSON -----
json_data = json.dumps(st.session_state.points, indent=4)
col_json.download_button(
    label="💾 2. Sauvegarder l'historique (JSON)", 
    data=json_data, 
    file_name=f"historique_{nom_projet.replace(' ', '_')}.json"
)
