import streamlit as st
import fitz  # PyMuPDF
import pandas as pd
import io
import re

st.set_page_config(page_title="Extracteur PDF Pro", page_icon="🛡️")

# --- FONCTION DE NETTOYAGE (Anti-Erreur IllegalCharacterError) ---
def clean_string(text):
    """Supprime les caractères non-imprimables qui font planter Excel."""
    if not isinstance(text, str):
        return text
    # On garde les caractères standards et on retire les codes de contrôle ASCII invisibles
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]', '', text)

st.title("🛡️ Extracteur PDF Haute Compatibilité")
st.write("Cette version nettoie automatiquement les caractères spéciaux pour éviter les erreurs Excel.")

fichiers_uploades = st.file_uploader("Choisir les PDF", type="pdf", accept_multiple_files=True)

if fichiers_uploades:
    toutes_les_donnees = []
    
    for fichier in fichiers_uploades:
        try:
            doc = fitz.open(stream=fichier.read(), filetype="pdf")
            
            # Extraction du texte
            texte_brut = ""
            for page in doc:
                texte_brut += page.get_text() + "\n"
            
            # Nettoyage systématique de chaque champ extrait
            texte_nettoye = clean_string(texte_brut.strip())
            nom_fichier_nettoye = clean_string(fichier.name)
            
            if texte_nettoye:
                toutes_les_donnees.append({
                    "Nom_Fichier": nom_fichier_nettoye,
                    "Statut": "Succès",
                    "Contenu": texte_nettoye
                })
            else:
                toutes_les_donnees.append({
                    "Nom_Fichier": nom_fichier_nettoye,
                    "Statut": "Avertissement",
                    "Contenu": "[Image ou Scan détecté - Aucun texte extractible]"
                })
                
        except Exception as e:
            st.error(f"Erreur sur {fichier.name} : {e}")

    if toutes_les_donnees:
        df = pd.DataFrame(toutes_les_donnees)
        
        st.success(f"Traitement terminé : {len(toutes_les_donnees)} fichiers.")
        st.dataframe(df.head(10)) # Aperçu des 10 premiers
        
        # Préparation du fichier Excel
        try:
            tampon = io.BytesIO()
            with pd.ExcelWriter(tampon, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 Télécharger l'Excel nettoyé",
                data=tampon.getvalue(),
                file_name="extraction_securisee.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Erreur lors de la génération Excel : {e}")
            st.info("Astuce : Si l'erreur persiste, un caractère très rare bloque encore le fichier.")

else:
    st.info("Sélectionnez vos fichiers PDF pour lancer l'aspirateur de texte.")