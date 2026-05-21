import streamlit as st
import pandas as pd
from thefuzz import process
import io

# ─────────────────────────────────────────────
# Config page
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="TaxRef Match",
    page_icon="🌿",
    layout="wide",
)

st.title("🌿 TaxRef Match — v18")
st.markdown(
    "Outil de réconciliation de noms d'espèces avec le référentiel **TaxRef v18**. "
    "Uploadez vos deux fichiers, lancez le matching, puis téléchargez le résultat."
)
st.divider()


# ─────────────────────────────────────────────
# Fonctions (identiques au script original)
# ─────────────────────────────────────────────

def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).strip()
    parts = name.split()
    if len(parts) >= 2:
        name = parts[0] + " " + parts[1]
    return " ".join(name.split()).lower()


@st.cache_data(show_spinner="Chargement du TaxRef en mémoire… (une seule fois par session)")
def load_taxref(file_bytes):
    taxref = pd.read_csv(
        io.BytesIO(file_bytes),
        sep="\t",
        quotechar='"',
        dtype=str,
        low_memory=False,
    )

    if "LB_NOM" not in taxref.columns:
        raise ValueError("Le fichier TaxRef ne contient pas la colonne 'LB_NOM'.")

    taxref["nom_clean"] = taxref["LB_NOM"].fillna("").apply(clean_name)
    taxref["CD_NOM"]    = taxref["CD_NOM"].fillna("")
    taxref["CD_REF"]    = taxref["CD_REF"].fillna("")
    taxref["RANG"]      = taxref["RANG"].fillna("")

    for col in ["FAMILLE", "NOM_COMPLET", "NOM_VALIDE"]:
        if col not in taxref.columns:
            taxref[col] = ""
        else:
            taxref[col] = taxref[col].fillna("")

    return taxref


def match_taxref(df, taxref, col_name="Nom_cite"):
    df = df.copy()
    df["nom_clean"] = df[col_name].apply(clean_name)

    merged = df.merge(
        taxref[["nom_clean", "CD_NOM", "CD_REF", "RANG", "FAMILLE", "NOM_COMPLET", "NOM_VALIDE"]],
        on="nom_clean",
        how="left",
    )

    for col in ["Type", "Similarité", "NOM_VALIDE"]:
        if col not in merged.columns:
            merged[col] = pd.NA

    missing = merged[merged["CD_NOM"].isna() | (merged["CD_NOM"] == "")].copy()

    if len(missing) > 0:
        tax_list = taxref["nom_clean"].tolist()
        fuzzy_info = []

        progress_bar = st.progress(0, text="Fuzzy matching en cours…")
        total = len(missing)

        for i, name in enumerate(missing["nom_clean"]):
            if name and name.strip():
                res = process.extractOne(name, tax_list)
                match, score = (res[0], res[1]) if res else (None, 0)
            else:
                match, score = None, 0
            fuzzy_info.append((name, match, score))
            progress_bar.progress((i + 1) / total, text=f"Fuzzy matching… {i+1}/{total}")

        progress_bar.empty()

        fuzzy_df = pd.DataFrame(fuzzy_info, columns=["nom_clean", "match_clean", "similarity"])
        fuzzy_join = fuzzy_df.merge(
            taxref,
            left_on="match_clean",
            right_on="nom_clean",
            suffixes=("", "_tax"),
            how="left",
        )

        for _, row in fuzzy_join.iterrows():
            idxs = merged[merged["nom_clean"] == row["nom_clean"]].index
            for idx in idxs:
                merged.at[idx, "CD_NOM"]      = row.get("CD_NOM", "")
                merged.at[idx, "CD_REF"]      = row.get("CD_REF", "")
                merged.at[idx, "RANG"]        = row.get("RANG", "")
                merged.at[idx, "FAMILLE"]     = row.get("FAMILLE", "")
                merged.at[idx, "NOM_COMPLET"] = row.get("NOM_COMPLET", "")
                merged.at[idx, "NOM_VALIDE"]  = row.get("NOM_VALIDE", "")
                merged.at[idx, "Similarité"]  = row.get("similarity", 0)
                merged.at[idx, "Type"]        = "Flou"

    merged.loc[merged["Type"].isna(), "Similarité"] = 100
    merged.loc[merged["Type"].isna(), "Type"]        = "Exact"
    merged["Similarité"] = merged["Similarité"].fillna(0).astype(float)
    merged["Type"]       = merged["Type"].fillna("Inconnu")

    return merged


def create_output_table(merged):
    return pd.DataFrame({
        "Nom cité":                          merged["Nom_cite"],
        "Classification":                    merged["FAMILLE"],
        "CD_NOM":                            merged["CD_NOM"],
        "CD_REF":                            merged["CD_REF"],
        "Nom complet":                       merged["NOM_COMPLET"],
        "Nom valide":                        merged["NOM_VALIDE"],
        "Similarité":                        merged["Similarité"],
        "Niveau taxonomique le plus précis": merged["RANG"],
        "Type de réconciliation":            merged["Type"],
    })


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Interface
# ─────────────────────────────────────────────

col1, col2 = st.columns(2)

with col1:
    st.subheader("1 · Fichier TaxRef v18")
    taxref_file = st.file_uploader(
        "Uploader TAXREFv18.txt",
        type=["txt"],
        help="Fichier texte tabulé issu de TaxRef v18 (~300 Mo).",
    )

with col2:
    st.subheader("2 · Fichier occurrences")
    input_file = st.file_uploader(
        "Uploader votre fichier Excel (.xlsx)",
        type=["xlsx"],
        help="Doit contenir une colonne nommée exactement 'Nom_cite'.",
    )

st.divider()

if taxref_file and input_file:
    # Chargement TaxRef (mis en cache)
    try:
        taxref = load_taxref(taxref_file.read())
    except ValueError as e:
        st.error(f"❌ Erreur TaxRef : {e}")
        st.stop()

    # Chargement occurrences
    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Excel : {e}")
        st.stop()

    if "Nom_cite" not in df.columns:
        st.error("❌ Le fichier Excel doit contenir une colonne **'Nom_cite'** (respect de la casse).")
        st.stop()

    st.success(f"✅ {len(df)} occurrences chargées — {len(taxref)} entrées TaxRef en mémoire.")

    # Aperçu
    with st.expander("Aperçu du fichier occurrences"):
        st.dataframe(df.head(10), use_container_width=True)

    # Lancement
    if st.button("🚀 Lancer le matching", type="primary", use_container_width=True):
        with st.spinner("Matching en cours…"):
            try:
                merged = match_taxref(df, taxref)
                final  = create_output_table(merged)
            except Exception as e:
                st.error(f"❌ Erreur pendant le matching : {e}")
                st.stop()

        st.success("✅ Matching terminé !")

        # Statistiques
        exact_count = (final["Type de réconciliation"] == "Exact").sum()
        fuzzy_count = (final["Type de réconciliation"] == "Flou").sum()
        other_count = len(final) - exact_count - fuzzy_count

        m1, m2, m3 = st.columns(3)
        m1.metric("Correspondances exactes", exact_count)
        m2.metric("Correspondances floues",  fuzzy_count)
        m3.metric("Non réconciliés",         other_count)

        # Résultats
        st.subheader("Résultats")
        st.dataframe(final, use_container_width=True, height=400)

        # Téléchargement
        st.download_button(
            label="⬇️ Télécharger le résultat (.xlsx)",
            data=to_excel_bytes(final),
            file_name="output_taxref_match.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )

else:
    st.info("⬆️ Uploadez les deux fichiers ci-dessus pour commencer.")
