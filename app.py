import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
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
st.set_page_config(page_title="App PV Maulini", layout="wide")
st.title("⚡ Assistant PV de Séance")

# --- 1. PARAMÈTRES GÉNÉRAUX & MÉMOIRE ---
with st.sidebar:
    st.header("📂 1. Infos du projet & Historique")
    
    # Informations de la séance
    nom_projet = st.text_input("Nom du Projet", "Construction Bas-Carbone")
    date_seance = st.date_input("Date de la séance", datetime.date.today())
    lieu = st.text_input("Lieu", "Genève")
    redacteur = st.text_input("Établi par", "Caroline")
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
st.subheader("👥 2. Présences")
col_pres, col_exc = st.columns(2)
presents = col_pres.text_area("Présents (un par ligne)", "Caroline (Maulini)\nReprésentant Weibel SA")
excuses = col_exc.text_area("Excusés / Absents", "Architecte (en congé)")

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

# 2. Titre Principal (Centré, grand, couleur gris foncé/bleuté)
titre = doc.add_heading(level=0)
titre.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_titre = titre.add_run("PROCÈS-VERBAL DE SÉANCE")
run_titre.font.name = 'Gotham'
run_titre.font.size = Pt(18)
run_titre.font.bold = True
run_titre.font.color.rgb = RGBColor(44, 62, 80) # Bleu-Gris élégant

doc.add_paragraph() # Espace

# 3. En-tête informatif (Tableau invisible pour aligner les textes proprement)
table_info = doc.add_table(rows=3, cols=2)
table_info.autofit = True
# Ligne 1
table_info.cell(0, 0).text = f"Projet : {nom_projet}"
table_info.cell(0, 1).text = f"Date : {date_seance.strftime('%d.%m.%Y')}"
# Ligne 2
table_info.cell(1, 0).text = f"Lieu : {lieu}"
table_info.cell(1, 1).text = f"Établi par : {redacteur}"
# Ligne 3
table_info.cell(2, 0).text = f"Prochaine séance : {prochaine_seance}"

# Mettre en gras les libellés dans le tableau info (un peu de style)
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

# 4. Liste des présences
p_pres = doc.add_paragraph()
p_pres.add_run("Présents : \n").bold = True
p_pres.add_run(presents.replace('\n', ', ') if presents else "Aucun")

p_exc = doc.add_paragraph()
p_exc.add_run("Excusés : \n").bold = True
p_exc.add_run(excuses.replace('\n', ', ') if excuses else "Aucun")

doc.add_paragraph() # Espace

# 5. Le grand tableau des points (Style épuré mais clair)
doc.add_heading("Suivi des Actions", level=1)
table_actions = doc.add_table(rows=1, cols=3)
table_actions.style = 'Table Grid'

# Styliser la ligne d'en-tête (Gris clair, texte centré et en gras)
hdr_cells = table_actions.rows[0].cells
en_tetes = ['Responsable', 'Action / Remarque', 'Délai']

for i, text in enumerate(en_tetes):
    hdr_cells[i].text = text
    set_cell_background(hdr_cells[i], 'EAEAEA') # Fond gris clair (#EAEAEA)
    paragraph = hdr_cells[i].paragraphs[0]
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.runs[0]
    run.font.bold = True
    run.font.name = 'Gotham'

# Remplissage dynamique des tâches
for p in st.session_state.points:
    row_cells = table_actions.add_row().cells
    
    # Responsable
    row_cells[0].text = p['qui']
    
    # Action avec gestion de la couleur
    p_quoi = row_cells[1].paragraphs[0]
    run_action = p_quoi.add_run(p['quoi'])
    run_action.font.name = 'Gotham'
    
    if p['color'] == "#0000FF":
        run_action.font.color.rgb = RGBColor(0, 102, 204) # Un beau bleu pro
    else:
        run_action.font.color.rgb = RGBColor(0, 0, 0)
        run_action.font.bold = True
        
    # Délai
    p_quand = row_cells[2].paragraphs[0]
    p_quand.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_quand.add_run(p['quand']).font.name = 'Gotham'

# Sauvegarde Word
bio_word = io.BytesIO()
doc.save(bio_word)
nom_fichier_word = f"PV_{nom_projet.replace(' ', '_')}_{date_seance.strftime('%Y%m%d')}.docx"

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
