import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Simulation Excel multi-feuilles", layout="wide")
st.title("📊 Simulation Excel multi-feuilles")

# 1️⃣ Upload Excel
uploaded_file = st.file_uploader("Importer un fichier Excel (.xlsx)", type=["xlsx"])

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    st.write("Feuilles disponibles :", xls.sheet_names)

    # 2️⃣ Sélection de la feuille à modifier
    sheet_selected = st.selectbox("Choisir une feuille à modifier", xls.sheet_names)
    df = pd.read_excel(uploaded_file, sheet_name=sheet_selected)

    st.subheader("Aperçu de la feuille sélectionnée")
    st.dataframe(df)

    # 3️⃣ Masquer / afficher certaines colonnes
    cols_selected = st.multiselect(
        "Colonnes à afficher",
        df.columns.tolist(),
        default=df.columns.tolist()
    )
    df_display = df[cols_selected]
    st.subheader("Colonnes affichées")
    st.dataframe(df_display)

    # 4️⃣ Modifier les colonnes “calculées” simples
    # Exemple: détecter une colonne contenant "Valeur ajustée"
    calc_cols = [col for col in df.columns if "Valeur ajustée" in col]

    for col in calc_cols:
        st.subheader(f"Modifier la colonne calculée : {col}")
        # Ancienne formule (exemple simplifié)
        old_formula = f"{col} = col_base * facteur"
        st.write("Ancienne formule :", old_formula)
        # Nouvelle formule entrée par le client
        new_formula = st.text_input(f"Nouvelle formule pour {col}", value=old_formula)

        # Appliquer la formule avec eval
        # Attention: ce code est pour une formule simple et illustrative
        # Ici col_base et facteur doivent être présents dans df
        try:
            df[col] = df.eval(new_formula)
        except Exception as e:
            st.error(f"Erreur dans la formule pour {col}: {e}")

    # 5️⃣ Générer le fichier final avec toutes les feuilles
    sheets_dict = {}
    for sheet in xls.sheet_names:
        if sheet == sheet_selected:
            sheets_dict[sheet] = df
        else:
            sheets_dict[sheet] = pd.read_excel(uploaded_file, sheet_name=sheet)

    # Créer Excel en mémoire
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        for sheet_name, sheet_df in sheets_dict.items():
            sheet_df.to_excel(writer, sheet_name=sheet_name, index=False)
    output.seek(0)

    # 6️⃣ Télécharger le fichier final
    st.download_button(
        "📥 Télécharger le fichier Excel final",
        data=output,
        file_name="fichier_final.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
