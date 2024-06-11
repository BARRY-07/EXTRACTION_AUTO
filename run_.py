import streamlit as st
from utilities_ import extraire_titres_numerotes, convert_df_to_excel

def main():
    st.set_page_config(page_title="FAST DPGF", layout="wide")

    # Couleurs basées sur le logo
    primaryColor = "#8B4513"  # brun foncé
    secondaryColor = "#D4A76A"  # couleur dorée
    backgroundColor = "#FFF8DC"  # crème légère
    secondaryBackgroundColor = "#F5F5F5"  # gris très clair
    textColor = "#363636"
    font = "sans serif"

    st.markdown(
        f"""
        <style>
            @keyframes gradient {{
                0% {{ background-position: 0% 50%; }}
                50% {{ background-position: 100% 50%; }}
                100% {{ background-position: 0% 50%; }}
            }}
            .reportview-container {{
                background: linear-gradient(-45deg, {primaryColor}, {secondaryColor}, {backgroundColor}, {secondaryBackgroundColor});
                background-size: 400% 400%;
                animation: gradient 15s ease infinite;
                color: {textColor};
                font-family: {font};
            }}
            .sidebar .sidebar-content {{
                background: {secondaryBackgroundColor};
                padding: 20px;
            }}
            header .decoration {{
                background: {primaryColor};
            }}
            .stButton>button {{
                color: white;
                background-color: {primaryColor};
                border: none;
                padding: 10px 24px;
                border-radius: 8px;
                transition: background-color 0.3s ease;
            }}
            .stButton>button:hover {{
                background-color: {secondaryColor};
            }}
            .css-18e3th9 {{
                padding: 10px;
            }}
            .css-1d391kg {{
                padding-top: 3.5rem;
                padding-left: 1rem;
                padding-right: 1rem;
            }}
            .dataframe {{
                width: 100% !important;
                height: auto;
            }}
            .marquee {{
                font-size: 24px;
                color: white;
                background-color: {primaryColor};
                padding: 10px;
                border-radius: 5px;
                margin: 20px 0;
                width: 100%;
                overflow: hidden;
                position: relative;
            }}
            .marquee div {{
                display: inline-block;
                width: 100%;
                height: 100%;
                white-space: nowrap;
                animation: marquee 10s linear infinite;
            }}
            @keyframes marquee {{
                0%   {{ transform: translateX(100%); }}
                100% {{ transform: translateX(-100%); }}
            }}
        </style>
        """,
        unsafe_allow_html=True
    )

    with st.sidebar:
        st.image("logo-cetab.jpg", width=200)
        st.markdown("""
            <h2>Groupe CETAB</h2>
            <p>
Le Groupe CETAB (Centre Etude Technique Aquitain du Bâtiment) est un Bureau d’études pluridisciplinaire spécialisé dans l’ingénierie du bâtiment, de l’infrastructure et de l’environnement.</p>
        """, unsafe_allow_html=True)

        st.write("## Paramètres")
        uploaded_file = st.file_uploader("Choisissez le contrat (.docx)", type=['docx'])

        # Champs de saisie pour les contenus des cellules A2 et C2
        cell_A2_content = st.text_input("Batiment", "Extension bâtiment ONERA – Bâtiments H2 et O à PALAISEAU")
        cell_C2_content = st.text_input("N° du Lot", "xx - xxxxxxxxxxxxxxxxx")
        feuille = 'LOT ' + cell_C2_content[:2]

    # Texte défilant
    st.markdown("""
    <div class="marquee">
        <div>Ceci est un outil interne au groupe CETAB permettant une extraction rapide des ouvrages dans les contrats.</div>
    </div>
    """, unsafe_allow_html=True)

    st.title("DECOMPOSEUR DE PRIX GLOBAL ET FORFAITAIRE")

    if uploaded_file is not None:
        with st.spinner('Extraction des ouvrages...'):
            df_titres = extraire_titres_numerotes(uploaded_file)
            df_titres_ = df_titres.set_index('N°')
        st.dataframe(df_titres_, height=800, width=1200)

        if st.button('Télécharger le DPGF au format Excel'):
            excel_data = convert_df_to_excel(df_titres, cell_A2_content, cell_C2_content, feuille)
            st.download_button(label="📥 Télécharger",
                               data=excel_data,
                               file_name=f'DPGF_Lot_{cell_C2_content}.xlsx',
                               mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == "__main__":
    main()
