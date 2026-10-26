import streamlit as st
import pandas as pd
import plotly.express as px

# ==========================================
# calculBA - Analyse de données en génie civil
# ==========================================

st.set_page_config(
    page_title="calculBA | Génie civil",
    page_icon="🏗️",
    layout="wide",
)

st.title("🏗️ calculBA")
st.subheader("Application d'analyse de données en génie civil")
st.caption("Collecte • Calculs • Statistiques • Visualisation")

COLONNES = ["Projet", "Materiau", "Resistance_MPa", "Cout_FCFA", "Quantite"]


def dataframe_vide():
    return pd.DataFrame({
        "Projet": pd.Series(dtype="str"),
        "Materiau": pd.Series(dtype="str"),
        "Resistance_MPa": pd.Series(dtype="float"),
        "Cout_FCFA": pd.Series(dtype="float"),
        "Quantite": pd.Series(dtype="float"),
    })


if "donnees" not in st.session_state:
    st.session_state.donnees = dataframe_vide()

menu = st.sidebar.radio(
    "Navigation",
    [
        "Accueil",
        "Saisie des données",
        "Importer un fichier",
        "Analyse statistique",
        "Calculs de génie civil",
        "Exporter les données",
    ],
)

if menu == "Accueil":
    st.header("Bienvenue sur calculBA")
    st.write(
        "Une plateforme pour organiser, analyser et visualiser "
        "des données de projets de génie civil."
    )

    df = st.session_state.donnees
    col1, col2, col3 = st.columns(3)
    col1.metric("Nombre d'enregistrements", len(df))

    resistances = pd.to_numeric(df["Resistance_MPa"], errors="coerce").dropna()
    couts = pd.to_numeric(df["Cout_FCFA"], errors="coerce").dropna()
    col2.metric(
        "Résistance moyenne",
        f"{resistances.mean():.2f} MPa" if not resistances.empty else "—",
    )
    col3.metric(
        "Coût total",
        f"{couts.sum():,.0f} FCFA" if not couts.empty else "0 FCFA",
    )

    st.info(
        "Commence par saisir des données ou importer un fichier CSV "
        "pour obtenir tes premières analyses."
    )
    st.markdown(
        """
        ### Fonctionnalités
        - Saisie manuelle des données de projets.
        - Importation de fichiers CSV.
        - Statistiques descriptives et graphiques.
        - Calculs indicatifs de génie civil.
        - Exportation des données analysées.
        """
    )

elif menu == "Saisie des données":
    st.header("📝 Saisie des données")

    with st.form("formulaire_donnees"):
        projet = st.text_input("Nom du projet")
        materiau = st.selectbox(
            "Matériau", ["Béton", "Acier", "Sol", "Granulat", "Autre"]
        )
        resistance = st.number_input(
            "Résistance (MPa)", min_value=0.0, value=25.0, step=0.5
        )
        cout = st.number_input(
            "Coût (FCFA)", min_value=0.0, value=0.0, step=1000.0
        )
        quantite = st.number_input(
            "Quantité", min_value=0.0, value=1.0, step=1.0
        )
        ajouter = st.form_submit_button("Ajouter les données")

    if ajouter:
        if not projet.strip():
            st.error("Saisis le nom du projet.")
        else:
            nouvelle_ligne = pd.DataFrame([{
                "Projet": projet.strip(),
                "Materiau": materiau,
                "Resistance_MPa": resistance,
                "Cout_FCFA": cout,
                "Quantite": quantite,
            }])
            st.session_state.donnees = pd.concat(
                [st.session_state.donnees, nouvelle_ligne],
                ignore_index=True,
            )
            st.success("Données ajoutées avec succès.")

    st.dataframe(st.session_state.donnees, use_container_width=True)

    if st.button("Effacer toutes les données saisies"):
        st.session_state.donnees = dataframe_vide()
        st.rerun()

elif menu == "Importer un fichier":
    st.header("📂 Importer des données")
    st.write(
        "Importe un fichier CSV contenant les colonnes : "
        "Projet, Materiau, Resistance_MPa, Cout_FCFA, Quantite."
    )
    fichier = st.file_uploader("Choisir un fichier CSV", type=["csv"])

    if fichier is not None:
        try:
            donnees_importees = pd.read_csv(fichier)
            st.write("Aperçu du fichier")
            st.dataframe(donnees_importees, use_container_width=True)

            if st.button("Ajouter les données importées"):
                manquantes = [
                    col for col in COLONNES if col not in donnees_importees.columns
                ]
                if manquantes:
                    st.error("Colonnes manquantes : " + ", ".join(manquantes))
                else:
                    donnees_importees = donnees_importees[COLONNES].copy()
                    for col in ["Resistance_MPa", "Cout_FCFA", "Quantite"]:
                        donnees_importees[col] = pd.to_numeric(
                            donnees_importees[col], errors="coerce"
                        )
                    donnees_importees["Projet"] = donnees_importees["Projet"].fillna("")
                    donnees_importees["Materiau"] = donnees_importees["Materiau"].fillna("Autre")
                    st.session_state.donnees = pd.concat(
                        [st.session_state.donnees, donnees_importees],
                        ignore_index=True,
                    )
                    st.success("Importation terminée.")
        except Exception as erreur:
            st.error(f"Impossible de lire le fichier : {erreur}")

elif menu == "Analyse statistique":
    st.header("📊 Analyse statistique")
    df = st.session_state.donnees.copy()

    if df.empty:
        st.warning("Ajoute ou importe des données pour commencer.")
    else:
        st.subheader("Données disponibles")
        st.dataframe(df, use_container_width=True)

        numeriques = []
        for col in ["Resistance_MPa", "Cout_FCFA", "Quantite"]:
            if pd.to_numeric(df[col], errors="coerce").notna().any():
                numeriques.append(col)

        if numeriques:
            variable = st.selectbox("Variable à analyser", numeriques)
            valeurs = pd.to_numeric(df[variable], errors="coerce").dropna()

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Moyenne", f"{valeurs.mean():,.2f}")
            c2.metric("Médiane", f"{valeurs.median():,.2f}")
            c3.metric("Minimum", f"{valeurs.min():,.2f}")
            c4.metric("Maximum", f"{valeurs.max():,.2f}")

            st.plotly_chart(
                px.histogram(
                    df, x=variable, title=f"Distribution de {variable}", nbins=15
                ),
                use_container_width=True,
            )
            st.plotly_chart(
                px.box(df, y=variable, title=f"Dispersion de {variable}"),
                use_container_width=True,
            )

        if df["Materiau"].notna().any():
            st.plotly_chart(
                px.histogram(
                    df, x="Materiau", color="Materiau",
                    title="Répartition par matériau"
                ),
                use_container_width=True,
            )

        df_couts = df.copy()
        df_couts["Cout_FCFA"] = pd.to_numeric(
            df_couts["Cout_FCFA"], errors="coerce"
        )
        if df_couts["Cout_FCFA"].notna().any():
            st.plotly_chart(
                px.bar(
                    df_couts, x="Projet", y="Cout_FCFA", color="Materiau",
                    title="Coût par enregistrement et projet"
                ),
                use_container_width=True,
            )

        st.subheader("Résumé statistique")
        st.dataframe(
            df.select_dtypes(include="number").describe().T,
            use_container_width=True,
        )

elif menu == "Calculs de génie civil":
    st.header("🧮 Calculs de génie civil")
    calcul = st.selectbox(
        "Choisir un calcul",
        ["Volume de béton", "Dosage en ciment", "Contrainte", "Poids volumique"],
    )

    if calcul == "Volume de béton":
        st.subheader("Volume d'un élément")
        longueur = st.number_input("Longueur (m)", min_value=0.0, value=5.0)
        largeur = st.number_input("Largeur (m)", min_value=0.0, value=3.0)
        hauteur = st.number_input(
            "Épaisseur/hauteur (m)", min_value=0.0, value=0.2
        )
        if st.button("Calculer le volume"):
            st.success(f"Volume géométrique : {longueur * largeur * hauteur:.3f} m³")

    elif calcul == "Dosage en ciment":
        st.subheader("Estimation de la masse de ciment")
        volume = st.number_input(
            "Volume de béton (m³)", min_value=0.0, value=1.0
        )
        dosage = st.number_input("Dosage (kg/m³)", min_value=0.0, value=350.0)
        if st.button("Calculer la quantité de ciment"):
            masse = volume * dosage
            st.success(f"Ciment nécessaire : {masse:.1f} kg")
            st.write(f"Soit environ {masse / 50:.2f} sacs de 50 kg.")

    elif calcul == "Contrainte":
        st.subheader("Contrainte moyenne")
        force = st.number_input("Force appliquée (kN)", min_value=0.0, value=100.0)
        surface = st.number_input(
            "Surface sollicitée (m²)", min_value=0.000001, value=0.1
        )
        if st.button("Calculer la contrainte"):
            contrainte = force / surface
            st.success(f"Contrainte moyenne : {contrainte:.2f} kN/m²")
            st.caption("1 kN/m² = 0,001 MPa.")

    elif calcul == "Poids volumique":
        st.subheader("Poids volumique")
        poids = st.number_input("Poids (kN)", min_value=0.0, value=20.0)
        volume = st.number_input("Volume (m³)", min_value=0.000001, value=1.0)
        if st.button("Calculer le poids volumique"):
            st.success(f"Poids volumique : {poids / volume:.2f} kN/m³")

    st.warning(
        "Ces calculs sont indicatifs. Ils ne remplacent pas les vérifications "
        "réglementaires, les études de sol ni la validation d'un ingénieur."
    )

elif menu == "Exporter les données":
    st.header("📥 Exporter les données")
    df = st.session_state.donnees

    if df.empty:
        st.info("Aucune donnée à exporter.")
    else:
        st.dataframe(df, use_container_width=True)
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "Télécharger les données en CSV",
            data=csv,
            file_name="calculBA_donnees.csv",
            mime="text/csv",
        )
        st.caption("Le fichier téléchargé peut être réimporté ultérieurement.")

st.sidebar.divider()
st.sidebar.caption("calculBA • Prototype de démonstration")
