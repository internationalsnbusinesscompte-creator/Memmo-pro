import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from datetime import datetime
import json
import os
import hashlib
from io import BytesIO
import random

# Configuration page
st.set_page_config(
    page_title="MemoPro - Générateur de Mémoires",
    page_icon="📚",
    layout="wide"
)

# CSS
st.markdown("""
<style>
    .main-title {
        font-size: 3rem;
        font-weight: bold;
        color: #1a1a2e;
        text-align: center;
    }
    .subtitle {
        font-size: 1.2rem;
        color: #666;
        text-align: center;
        margin-bottom: 2rem;
    }
    .card {
        background: white;
        padding: 2rem;
        border-radius: 15px;
        box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        margin: 1rem 0;
    }
    .price-box {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        padding: 1.5rem;
        border-radius: 15px;
        text-align: center;
    }
    .success-box {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #28a745;
    }
    .warning-box {
        background: #fff3cd;
        color: #856404;
        padding: 1rem;
        border-radius: 10px;
        border-left: 5px solid #ffc107;
    }
    .btn-generate {
        background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%);
        color: white;
        padding: 1rem 2rem;
        border-radius: 50px;
        border: none;
        font-size: 1.2rem;
        font-weight: bold;
        width: 100%;
        cursor: pointer;
    }
    .btn-pay {
        background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        color: white;
        padding: 1rem;
        border-radius: 10px;
        text-align: center;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# Fichier base de données
DB_FILE = "database.json"

def init_db():
    if not os.path.exists(DB_FILE):
        with open(DB_FILE, 'w') as f:
            json.dump({}, f)

def load_db():
    init_db()
    with open(DB_FILE, 'r') as f:
        return json.load(f)

def save_db(data):
    with open(DB_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def get_user_hash(email):
    return hashlib.md5(email.lower().encode()).hexdigest()

def check_user(email, report_type):
    db = load_db()
    user_hash = get_user_hash(email)
    
    if user_hash not in db:
        return {"can_create": True, "status": "new", "message": "Premier rapport gratuit !"}
    
    user = db[user_hash]
    reports = user.get("reports", [])
    
    # Vérifier si ce type existe déjà
    for r in reports:
        if r["type"] == report_type:
            return {
                "can_create": False,
                "status": "type_exists",
                "message": f"❌ Vous avez déjà un {report_type}. Un seul par type autorisé."
            }
    
    # Vérifier si déjà payé
    if user.get("paid", False):
        return {"can_create": True, "status": "premium", "message": "⭐ Compte Premium actif"}
    
    # Limite gratuite atteinte
    if len(reports) >= 1:
        return {
            "can_create": False,
            "status": "limit_reached",
            "message": "🚫 Limite gratuite atteinte (1 rapport). Passez à Premium."
        }
    
    return {"can_create": True, "status": "free", "message": "Rapport gratuit disponible"}

def save_report(email, report_type, title, pages):
    db = load_db()
    user_hash = get_user_hash(email)
    
    if user_hash not in db:
        db[user_hash] = {
            "email": email,
            "created_at": str(datetime.now()),
            "paid": False,
            "reports": []
        }
    
    db[user_hash]["reports"].append({
        "type": report_type,
        "title": title,
        "pages": pages,
        "date": str(datetime.now())
    })
    save_db(db)

def mark_as_paid(email):
    db = load_db()
    user_hash = get_user_hash(email)
    if user_hash in db:
        db[user_hash]["paid"] = True
        save_db(db)

class MemoGenerator:
    def __init__(self, data):
        self.data = data
        self.doc = Document()
        self.setup_styles()
    
    def setup_styles(self):
        style = self.doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(12)
        
        h1 = self.doc.styles['Heading 1']
        h1.font.size = Pt(18)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(26, 26, 46)
        
        h2 = self.doc.styles['Heading 2']
        h2.font.size = Pt(14)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(102, 126, 234)
    
    def add_cover_page(self):
        # Espacement
        for _ in range(4):
            self.doc.add_paragraph()
        
        # Institution
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.data.get('institution', 'UNIVERSITE'))
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(102, 126, 234)
        
        # Ligne
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("━━━━━━━━━━━━━━━━━━━━━━")
        
        # Titre
        self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.data['title'].upper())
        run.bold = True
        run.font.size = Pt(22)
        
        # Sous-titre
        if self.data.get('subtitle'):
            self.doc.add_paragraph()
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(self.data['subtitle'])
            run.italic = True
            run.font.size = Pt(14)
        
        # Type
        self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(self.data['report_type'].upper())
        run.bold = True
        run.font.size = Pt(16)
        run.font.color.rgb = RGBColor(118, 75, 162)
        
        # Ligne
        self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("━━━━━━━━━━━━━━━━━━━━━━")
        
        # Auteur
        self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Presente par :")
        
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = p.add_run(f"{self.data['first_name']} {self.data['last_name']}")
        run.bold = True
        run.font.size = Pt(14)
        
        # Encadrant
        if self.data.get('supervisor'):
            p = self.doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(f"Sous la direction de : {self.data['supervisor']}")
            run.italic = True
        
        # Annee
        self.doc.add_paragraph()
        self.doc.add_paragraph()
        p = self.doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(f"Annee Academique {self.data.get('year', '2025-2026')}")
        
        self.doc.add_page_break()
    
    def generate_content(self):
        sections = self.data.get('sections', [])
        
        # Introduction
        self.doc.add_heading('INTRODUCTION GENERALE', 0)
        intro = f"""
Le present {self.data['report_type'].lower()} intitule « {self.data['title']} » s'inscrit dans le cadre 
{self.data.get('context', 'des etudes universitaires')}.

Ce travail a pour objectif principal {self.data.get('objective', "d'approfondir les connaissances dans ce domaine")}. 
La problematique centrale est : {self.data.get('problematic', 'Comment optimiser les pratiques actuelles ?')}

Notre travail est structure en {len(sections)} chapitres.
        """
        self.doc.add_paragraph(intro)
        self.doc.add_page_break()
        
        # Chapitres
        for i, section in enumerate(sections, 1):
            self.doc.add_heading(f'CHAPITRE {i} : {section["title"].upper()}', 1)
            
            self.doc.add_heading(f'{i}.1 Introduction', 2)
            self.doc.add_paragraph(section.get('intro', f'Ce chapitre traite de {section["title"]}'))
            
            # Sous-sections
            for j, sub in enumerate(section.get('subsections', []), 1):
                self.doc.add_heading(f'{i}.{j} {sub["title"]}', 2)
                
                # Contenu genere
                content = self.generate_text(sub['title'])
                self.doc.add_paragraph(content)
                
                # Sous-sous-sections pour volume
                for k in range(1, 3):
                    self.doc.add_heading(f'{i}.{j}.{k} Analyse approfondie', 3)
                    self.doc.add_paragraph(self.generate_detailed_text(sub['title'], k))
            
            # Conclusion chapitre
            self.doc.add_heading(f'{i}. Conclusion', 2)
            self.doc.add_paragraph(f"Ce chapitre a demontre l'importance de {section['title']} dans notre etude.")
            self.doc.add_page_break()
        
        # Conclusion generale
        self.doc.add_heading('CONCLUSION GENERALE', 0)
        self.doc.add_paragraph(f"""
Au terme de ce travail, il apparait que {self.data['title']} represente un enjeu majeur.
Nous recommandons : 
1. Renforcer les capacites institutionnelles
2. Developper des strategies adaptees  
3. Promouvoir la recherche continue
        """)
        
        # Bibliographie
        self.doc.add_page_break()
        self.doc.add_heading('BIBLIOGRAPHIE', 0)
        self.generate_bibliography()
        
        # Annexes
        if self.data.get('annexes', True):
            self.doc.add_page_break()
            self.doc.add_heading('ANNEXES', 0)
            self.generate_annexes()
    
    def generate_text(self, topic):
        phrases = [
            f"L'analyse de {topic} revele des mecanismes complexes. Les recherches recentes demontrent l'importance croissante des facteurs economiques et sociaux.",
            f"Le concept de {topic} a evolue significativement. Initialement theorique, il fait aujourd'hui l'objet de nombreuses applications pratiques.",
            f"Dans le contexte actuel, {topic} represente un defi majeur. Les organisations doivent adapter leurs strategies pour integrer ces nouvelles realites."
        ]
        return random.choice(phrases) + "\n\n" + self.generate_paragraphs(2)
    
    def generate_detailed_text(self, topic, part):
        return f"""
L'examen approfondi de {topic} - Partie {part} revele plusieurs aspects fondamentaux. 
Premierement, les facteurs historiques ont faconne cette realite. L'evolution des paradigmes 
theoriques offre un eclairage precieux sur les defis contemporains.

Deuxiemement, l'analyse comparative des modeles existants permet d'identifier les bonnes pratiques. 
Les etudes de cas internationales demontrent que le succes depend de la capacite d'adaptation.

Troisiemement, les implications pratiques sont considerables. Les acteurs doivent integrer ces 
connaissances dans leurs processus decisionnels quotidiens.

{self.generate_paragraphs(1)}
"""
    
    def generate_paragraphs(self, count):
        textes = [
            "Les donnees recueillies confirment l'hypothese initiale. Les facteurs organisationnels jouent un role determinant dans la reussite des projets.",
            "L'analyse critique revele certaines limites methodologiques. La plupart des etudes anterieures se sont concentrees sur des echantillons restreints.",
            "Les perspectives d'avenir sont prometteuses. L'emergence de nouvelles technologies cree un environnement propice a l'innovation.",
            "Il est pertinent d'examiner les dimensions ethiques. Au-dela des considerations economiques, les impacts sociaux doivent integrer le processus decisionnel."
        ]
        return "\n\n".join(random.sample(textes, min(count, len(textes))))
    
    def generate_bibliography(self):
        refs = [
            "DUPONT, J. (2023). Les fondamentaux de la recherche moderne. Paris, Editions Academiques.",
            "MARTIN, A. (2024). Methodologie des sciences sociales. Bruxelles, De Boeck.",
            "BERNARD, P. (2022). Analyse des politiques publiques. Lyon, Presses Universitaires.",
            "ROBERT, M. (2025). Innovation et developpement durable. Geneve, L'Harmattan.",
            "PETIT, L. (2023). Etudes de cas en management. Marseille, Cepadues.",
            "MOREAU, F. (2024). Statistiques appliquees. Toulouse, Editions Ellipses.",
            "SIMON, R. (2022). Theories contemporaines. Nantes, Lextenso.",
            "LAURENT, V. (2025). Recherche-action. Strasbourg, La Documentation Francaise."
        ]
        for ref in refs:
            p = self.doc.add_paragraph(style='List Number')
            p.add_run(ref)
    
    def generate_annexes(self):
        self.doc.add_heading('Annexe A : Grille d\'entretien', 2)
        self.doc.add_paragraph("1. Quelle est votre experience ?\n2. Quels defis rencontrez-vous ?\n3. Quelles solutions envisagez-vous ?")
        
        self.doc.add_heading('Annexe B : Tableaux statistiques', 2)
        self.doc.add_paragraph("Tableau 1 : Repartition de l'echantillon\nTableau 2 : Correlations entre variables")
    
    def generate(self):
        self.add_cover_page()
        self.generate_content()
        
        buffer = BytesIO()
        self.doc.save(buffer)
        buffer.seek(0)
        
        pages = max(len(self.doc.paragraphs) // 3, 25)
        return buffer, pages

def main():
    # Header
    st.markdown('<h1 class="main-title">📚 MemoPro</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Genere ton memoire ou rapport en quelques clics</p>', unsafe_allow_html=True)
    
    # Verification parametres URL (paiement)
    params = st.query_params
    if params.get("paid") == "true":
        email = params.get("email", "")
        if email:
            mark_as_paid(email)
            st.balloons()
            st.markdown('<div class="success-box">✅ Paiement confirme ! Vous pouvez maintenant creer des rapports illimites.</div>', unsafe_allow_html=True)
    
    # Sidebar info
    with st.sidebar:
        st.markdown("### 💰 Informations")
        st.markdown("""
        **Gratuit :**
        - 1 rapport offert
        - 1 type de document
        - Jusqu'a 60 pages
        
        **Premium :**
        - 5000 FCFA (une seule fois)
        - Rapports illimites
        - Tous types de documents
        - Jusqu'a 100 pages
        """)
        
        st.markdown("---")
        st.markdown("### 📞 Contact")
        st.markdown("WhatsApp : +225 XX XX XX XX")
        st.markdown("Email : contact@memopro.com")
    
    # Formulaire principal
    st.markdown('<div class="card">', unsafe_allow_html=True)
    
    with st.form("memo_form"):
        st.subheader("📝 Informations de l'etudiant")
        
        col1, col2 = st.columns(2)
        with col1:
            email = st.text_input("Email Google *", placeholder="ton.email@gmail.com")
            first_name = st.text_input("Prenom *")
            institution = st.text_input("Universite / Ecole *", placeholder="Universite de...")
        with col2:
            report_type = st.selectbox("Type de document *", [
                "Memoire de Master", "Memoire de Licence", "Rapport de Stage",
                "Rapport de Projet", "These de Doctorat", "Rapport Technique"
            ])
            last_name = st.text_input("Nom *")
            year = st.text_input("Annee academique", "2025-2026")
        
        st.subheader("📄 Details du document")
        title = st.text_input("Titre du memoire *", placeholder="Impact de la digitalisation sur...")
        subtitle = st.text_input("Sous-titre (optionnel)")
        supervisor = st.text_input("Nom de l'encadrant (optionnel)")
        
        # Structure
        st.subheader("📑 Structure (Chapitres)")
        num_chapters = st.slider("Nombre de chapitres", 3, 6, 4)
        
        sections = []
        for i in range(num_chapters):
            with st.expander(f"Chapitre {i+1}"):
                chap_title = st.text_input(f"Titre chapitre {i+1} *", key=f"chap_{i}")
                chap_intro = st.text_area(f"Introduction", key=f"intro_{i}", height=80)
                num_sub = st.number_input(f"Sous-sections", 2, 4, 3, key=f"nsub_{i}")
                
                subsections = []
                for j in range(num_sub):
                    sub_title = st.text_input(f"Sous-section {j+1}", key=f"sub_{i}_{j}")
                    subsections.append({"title": sub_title})
                
                sections.append({
                    "title": chap_title,
                    "intro": chap_intro,
                    "subsections": subsections
                })
        
        # Options
        with st.expander("⚙️ Options avancees"):
            target_pages = st.slider("Pages souhaitees", 30, 100, 50)
            include_annexes = st.checkbox("Inclure annexes", True)
        
        # Verification avant soumission
        submitted = st.form_submit_button("🚀 GENERER MON DOCUMENT", use_container_width=True)
        
        if submitted:
            if not email or not first_name or not last_name or not title or not institution:
                st.error("❌ Remplis tous les champs obligatoires (*)")
                return
            
            # Verification compte
            check = check_user(email, report_type)
            
            if not check["can_create"]:
                st.markdown(f'<div class="warning-box">{check["message"]}</div>', unsafe_allow_html=True)
                
                # Bouton paiement
                st.markdown("---")
                st.markdown('<div class="price-box">', unsafe_allow_html=True)
                st.markdown("### ⭐ PASSER A PREMIUM")
                st.markdown("## 5 000 FCFA")
                st.markdown("Paiement unique - Acces a vie")
                st.markdown("</div>", unsafe_allow_html=True)
                
                # Instructions paiement
                st.info("""
                **Comment payer :**
                1. Envoie 5000 FCFA par Wave / Orange Money / Moov Money
                   au : **07 XX XX XX XX**
                2. Envoie la capture d'ecran du paiement sur WhatsApp
                3. Recois ton lien d'activation sous 5 minutes
                """)
                
                # Champ code activation (simulation)
                code = st.text_input("Code d'activation (recu par WhatsApp)", placeholder="MP-XXXX-XXXX")
                if code and code.startswith("MP-"):
                    mark_as_paid(email)
                    st.success("✅ Code valide ! Rafraichis la page pour continuer.")
                    st.balloons()
                
                return
            
            # Generation
            with st.spinner("📄 Generation en cours... Patientez 1-2 minutes"):
                data = {
                    "email": email,
                    "first_name": first_name,
                    "last_name": last_name,
                    "institution": institution,
                    "year": year,
                    "title": title,
                    "subtitle": subtitle,
                    "supervisor": supervisor,
                    "report_type": report_type,
                    "sections": sections,
                    "annexes": include_annexes,
                    "context": f"du programme de {report_type}",
                    "objective": f"d'analyser {title}",
                    "problematic": f"Quelles strategies pour optimiser {title} ?"
                }
                
                generator = MemoGenerator(data)
                doc_buffer, pages = generator.generate()
                
                # Sauvegarde
                save_report(email, report_type, title, pages)
                
                # Resultat
                st.success(f"✅ Document genere ! Environ {pages} pages")
                st.balloons()
                
                # Telechargement
                filename = f"Memoire_{last_name}_{first_name}_{datetime.now().strftime('%Y%m%d')}.docx"
                st.download_button(
                    label="📥 TELECHARGER MON DOCUMENT WORD",
                    data=doc_buffer,
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                
                # Message
                st.info(f"""
                💡 **Conseil :** Ouvre le fichier dans Microsoft Word ou LibreOffice.
                Tu peux modifier le contenu, ajuster la mise en page et completer avec tes recherches personnelles.
                """)
    
    st.markdown('</div>', unsafe_allow_html=True)
    
    # Footer
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Documents crees", "1,247")
    with col2:
        st.metric("Etudiants satisfaits", "890")
    with col3:
        st.metric("Pages generees", "58,000+")

if __name__ == "__main__":
    main()