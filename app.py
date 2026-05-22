import streamlit as st
import pandas as pd
from thefuzz import process
from collections import defaultdict
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
# Contenu du README embarqué
# ─────────────────────────────────────────────
README_MD = """# 🌿 TaxRef Match — Guide utilisateur

## I. Comment utiliser l'outil

### Prérequis
- Un navigateur web (Chrome, Firefox, Edge…)
- Le fichier **TAXREFv18.txt** (référentiel MNHN, ~300 Mo)
- Un fichier Excel **(.xlsx)** rempli selon le format attendu

---

### Étape 1 — Télécharger le template

Depuis la page de l'application, cliquez sur **"Télécharger le template (.xlsx)"**.  
Ce fichier contient les trois colonnes attendues :

| Colonne | Obligatoire | Description |
|---|---|---|
| `fk` | Non | Clé primaire libre — identifiant que vous choisissez (numéro, code, etc.) |
| `nom_cite` | **Oui** | Nom d'espèce à réconcilier avec TaxRef |
| `Classification` | Non | Rang taxonomique supérieur pour valider le match (voir §4) |

---

### Étape 2 — Remplir le template

- **`fk`** : laissez vide ou remplissez avec vos propres identifiants. Cette colonne est conservée telle quelle dans le résultat final.
- **`nom_cite`** : saisissez les noms d'espèces tels que vous les avez (avec ou sans auteur, l'outil ne conserve que Genre + épithète).
- **`Classification`** : optionnel. Vous pouvez renseigner n'importe quel rang supérieur : règne, phylum, classe, ordre, famille, sous-famille. Exemples : `Animalia`, `Actinopterygii`, `Syngnathidae`. L'outil vérifiera si le match TaxRef est cohérent avec cette valeur.

---

### Étape 3 — Lancer le matching

1. Uploadez **TAXREFv18.txt** via le premier bouton
2. Uploadez votre **fichier Excel** via le second bouton
3. Cliquez sur **"Lancer le matching"**
4. Une barre de progression s'affiche pendant le fuzzy matching
5. Une fois terminé, des métriques résument les résultats

---

### Étape 4 — Télécharger le résultat

Cliquez sur **"Télécharger le résultat (.xlsx)"**. Le fichier contient les colonnes suivantes :

| Colonne | Description |
|---|---|
| `fk` | Votre clé primaire, inchangée |
| `Nom cité` | Le nom tel que saisi dans `nom_cite` |
| `CD_NOM` | Identifiant TaxRef du taxon trouvé |
| `CD_REF` | Identifiant TaxRef du taxon valide de référence |
| `Nom complet` | Nom complet TaxRef avec auteur |
| `Nom valide` | Nom valide selon TaxRef |
| `Similarité` | Score de correspondance (100 = exact, 80–99 = flou) |
| `Niveau taxonomique` | Rang TaxRef (ES, SSES, GN…) |
| `Type de réconciliation` | Exact, Flou ou Non réconcilié |
| `Validation classification` | ✅ Cohérent / ⚠️ Incohérent / — (si non renseigné) |

---

## II. Comment fonctionne l'outil

### 1. Nettoyage des noms (`clean_name`)

Avant tout matching, chaque nom est normalisé :
- Suppression de l'auteur et de l'année (ex. `Syngnathus taenionotus Cantor, 1850` → `syngnathus taenionotus`)
- Conservation uniquement du **Genre + épithète spécifique** (les deux premiers mots)
- Mise en minuscules et suppression des espaces superflus

Cette normalisation est appliquée aussi bien aux noms cités qu'aux noms TaxRef (`LB_NOM`).

---

### 2. Match exact

L'outil tente d'abord une **jointure exacte** entre le nom nettoyé et la colonne `nom_clean` de TaxRef.  
Si le nom est trouvé : score de similarité = 100, type = `Exact`.

---

### 3. Match flou (fuzzy matching)

Pour les noms non trouvés en exact, l'outil applique un **matching flou** avec deux contraintes :

**Contrainte A — Genre identique**  
L'outil ne compare le nom qu'avec les entrées TaxRef ayant **le même genre** (premier mot identique).  
Exemple : `syngnathus taenionotus` n'est comparé qu'aux autres `syngnathus *` de TaxRef.  
Cela évite les faux positifs comme `Syngnathus taenionotus` → `Notus Fieber` qui partageaient des lettres sans lien taxonomique.

**Contrainte B — Seuil de similarité à 80**  
Même au sein du même genre, si le meilleur score est inférieur à 80/100, le nom est classé `Non réconcilié` plutôt que de proposer un match douteux.

L'algorithme utilisé est **RapidFuzz / Levenshtein** via la bibliothèque `thefuzz`.

---

### 4. Validation par la Classification

Si la colonne `Classification` est renseignée, l'outil compare la valeur saisie avec les colonnes hiérarchiques du taxon trouvé dans TaxRef :  
`REGNE`, `PHYLUM`, `CLASSE`, `ORDRE`, `FAMILLE`, `SOUS_FAMILLE`.

- **✅ Cohérent** : la valeur correspond à l'un de ces rangs dans TaxRef
- **⚠️ Incohérent** : la valeur ne correspond à aucun rang → le match mérite vérification manuelle
- **—** : colonne `Classification` non renseignée pour cette ligne

---

### 5. Colonnes TaxRef utilisées

| Colonne TaxRef | Utilisation |
|---|---|
| `LB_NOM` | Nom de base pour le matching |
| `CD_NOM` | Identifiant du taxon |
| `CD_REF` | Identifiant du taxon valide |
| `RANG` | Niveau taxonomique |
| `FAMILLE` | Famille |
| `NOM_COMPLET` | Nom complet avec auteur |
| `NOM_VALIDE` | Nom valide |
| `REGNE` `PHYLUM` `CLASSE` `ORDRE` `SOUS_FAMILLE` | Validation de la classification |
"""


# ─────────────────────────────────────────────
# Fonctions utilitaires
# ─────────────────────────────────────────────

def clean_name(name):
    if pd.isna(name):
        return ""
    name = str(name).strip()
    parts = name.split()
    if len(parts) >= 2:
        name = parts[0] + " " + parts[1]
    return " ".join(name.split()).lower()


def get_template_bytes():
    df = pd.DataFrame(columns=["fk", "nom_cite", "Classification"])
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Chargement TaxRef
# ─────────────────────────────────────────────

RANK_COLS = ["REGNE", "PHYLUM", "CLASSE", "ORDRE", "FAMILLE", "SOUS_FAMILLE"]

@st.cache_data(show_spinner="Chargement du TaxRef en mémoire… (une seule fois par session)")
def load_taxref(file_bytes):
    cols_needed = [
        "LB_NOM", "CD_NOM", "CD_REF", "RANG",
        "FAMILLE", "NOM_COMPLET", "NOM_VALIDE",
        "REGNE", "PHYLUM", "CLASSE", "ORDRE", "SOUS_FAMILLE",
    ]
    taxref = pd.read_csv(
        io.BytesIO(file_bytes),
        sep="\t",
        quotechar='"',
        dtype=str,
        low_memory=False,
        usecols=lambda c: c in cols_needed,
    )

    if "LB_NOM" not in taxref.columns:
        raise ValueError("Le fichier TaxRef ne contient pas la colonne 'LB_NOM'.")

    taxref["nom_clean"] = taxref["LB_NOM"].fillna("").apply(clean_name)

    for col in ["CD_NOM", "CD_REF", "RANG"] + RANK_COLS + ["NOM_COMPLET", "NOM_VALIDE"]:
        if col not in taxref.columns:
            taxref[col] = ""
        else:
            taxref[col] = taxref[col].fillna("")

    return taxref


# ─────────────────────────────────────────────
# Matching
# ─────────────────────────────────────────────

def build_genus_index(taxref):
    """Construit un index {genre: [liste de noms_clean]} pour le fuzzy contraint."""
    genus_index = defaultdict(list)
    for name in taxref["nom_clean"]:
        parts = name.split()
        if parts:
            genus_index[parts[0]].append(name)
    return genus_index


def fuzzy_with_constraint(name, genus_index, threshold=80):
    """Fuzzy matching contraint au genre + seuil de rejet."""
    parts = name.split()
    if not parts:
        return None, 0
    genus = parts[0]
    candidates = genus_index.get(genus, [])
    if not candidates:
        return None, 0
    res = process.extractOne(name, candidates)
    if res is None:
        return None, 0
    match, score = res[0], res[1]
    if score < threshold:
        return None, 0
    return match, score


def validate_classification(user_val, taxref_row):
    """Compare la Classification saisie avec les rangs hiérarchiques TaxRef."""
    if pd.isna(user_val) or str(user_val).strip() == "":
        return "—"
    user_clean = str(user_val).strip().lower()
    for col in RANK_COLS:
        taxref_val = taxref_row.get(col, "")
        if str(taxref_val).strip().lower() == user_clean:
            return "✅ Cohérent"
    return "⚠️ Incohérent"


def match_taxref(df, taxref):
    df = df.copy()

    # Normalisation nom
    col_nom = "nom_cite" if "nom_cite" in df.columns else "Nom_cite"
    df["nom_clean"] = df[col_nom].apply(clean_name)

    # Gestion colonne fk
    has_fk = "fk" in df.columns

    # Gestion colonne Classification
    has_classif = "Classification" in df.columns

    # ── Match exact ──────────────────────────────
    merge_cols = ["nom_clean", "CD_NOM", "CD_REF", "RANG",
                  "FAMILLE", "NOM_COMPLET", "NOM_VALIDE"] + RANK_COLS
    merged = df.merge(taxref[merge_cols], on="nom_clean", how="left")

    for col in ["Type", "Similarité"]:
        if col not in merged.columns:
            merged[col] = pd.NA

    # ── Fuzzy matching pour les non matchés ──────
    missing_mask = merged["CD_NOM"].isna() | (merged["CD_NOM"] == "")
    missing = merged[missing_mask].copy()

    if len(missing) > 0:
        genus_index = build_genus_index(taxref)

        progress_bar = st.progress(0, text="Fuzzy matching en cours…")
        total = len(missing)
        fuzzy_info = []

        for i, name in enumerate(missing["nom_clean"]):
            if name and name.strip():
                match, score = fuzzy_with_constraint(name, genus_index, threshold=80)
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
                if row.get("match_clean") is None or pd.isna(row.get("match_clean")):
                    # Aucun match acceptable → Non réconcilié
                    merged.at[idx, "CD_NOM"]      = ""
                    merged.at[idx, "CD_REF"]       = ""
                    merged.at[idx, "RANG"]         = ""
                    merged.at[idx, "FAMILLE"]      = ""
                    merged.at[idx, "NOM_COMPLET"]  = ""
                    merged.at[idx, "NOM_VALIDE"]   = ""
                    merged.at[idx, "Similarité"]   = 0
                    merged.at[idx, "Type"]         = "Non réconcilié"
                    for rc in RANK_COLS:
                        merged.at[idx, rc] = ""
                else:
                    merged.at[idx, "CD_NOM"]      = row.get("CD_NOM", "")
                    merged.at[idx, "CD_REF"]       = row.get("CD_REF", "")
                    merged.at[idx, "RANG"]         = row.get("RANG", "")
                    merged.at[idx, "FAMILLE"]      = row.get("FAMILLE", "")
                    merged.at[idx, "NOM_COMPLET"]  = row.get("NOM_COMPLET", "")
                    merged.at[idx, "NOM_VALIDE"]   = row.get("NOM_VALIDE", "")
                    merged.at[idx, "Similarité"]   = row.get("similarity", 0)
                    merged.at[idx, "Type"]         = "Flou"
                    for rc in RANK_COLS:
                        merged.at[idx, rc] = row.get(rc, "")

    # Matches exacts
    merged.loc[merged["Type"].isna(), "Similarité"] = 100
    merged.loc[merged["Type"].isna(), "Type"]        = "Exact"
    merged["Similarité"] = merged["Similarité"].fillna(0).astype(float)
    merged["Type"]       = merged["Type"].fillna("Non réconcilié")

    # ── Validation Classification ─────────────────
    if has_classif:
        merged["Validation classification"] = merged.apply(
            lambda row: validate_classification(row.get("Classification", ""), row),
            axis=1,
        )
    else:
        merged["Validation classification"] = "—"

    return merged, col_nom, has_fk


def create_output_table(merged, col_nom, has_fk):
    data = {}
    if has_fk:
        data["fk"] = merged["fk"]
    data.update({
        "Nom cité":                          merged[col_nom],
        "Classification":                    merged["FAMILLE"],
        "CD_NOM":                            merged["CD_NOM"],
        "CD_REF":                            merged["CD_REF"],
        "Nom complet":                       merged["NOM_COMPLET"],
        "Nom valide":                        merged["NOM_VALIDE"],
        "Similarité":                        merged["Similarité"],
        "Niveau taxonomique le plus précis": merged["RANG"],
        "Type de réconciliation":            merged["Type"],
        "Validation classification":         merged["Validation classification"],
    })
    return pd.DataFrame(data)


# ─────────────────────────────────────────────
# Interface
# ─────────────────────────────────────────────

# Boutons de téléchargement ressources
with st.sidebar:
    st.header("📥 Ressources")
    st.download_button(
        label="📄 Télécharger le template (.xlsx)",
        data=get_template_bytes(),
        file_name="template_taxref_match.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.download_button(
        label="📖 Télécharger le guide (.md)",
        data=README_MD.encode("utf-8"),
        file_name="Guide_TaxRef_Match.md",
        mime="text/markdown",
        use_container_width=True,
    )
    st.divider()
    st.caption(
        "**Format attendu du fichier Excel :**\n\n"
        "- `fk` : clé primaire libre *(optionnel)*\n"
        "- `nom_cite` : nom d'espèce *(obligatoire)*\n"
        "- `Classification` : rang supérieur *(optionnel)*"
    )

# Upload
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
        help="Doit contenir une colonne 'nom_cite'. Colonnes optionnelles : 'fk', 'Classification'.",
    )

st.divider()

if taxref_file and input_file:
    try:
        taxref = load_taxref(taxref_file.read())
    except ValueError as e:
        st.error(f"❌ Erreur TaxRef : {e}")
        st.stop()

    try:
        df = pd.read_excel(input_file)
    except Exception as e:
        st.error(f"❌ Impossible de lire le fichier Excel : {e}")
        st.stop()

    # Validation colonne nom
    col_nom = "nom_cite" if "nom_cite" in df.columns else ("Nom_cite" if "Nom_cite" in df.columns else None)
    if col_nom is None:
        st.error("❌ Le fichier Excel doit contenir une colonne **`nom_cite`** (respect de la casse).")
        st.stop()

    st.success(f"✅ {len(df)} occurrences chargées — {len(taxref)} entrées TaxRef en mémoire.")

    # Info colonnes détectées
    detected = []
    if "fk" in df.columns:
        detected.append("`fk` ✅")
    if "Classification" in df.columns:
        detected.append("`Classification` ✅")
    if detected:
        st.info(f"Colonnes optionnelles détectées : {' · '.join(detected)}")

    with st.expander("Aperçu du fichier occurrences"):
        st.dataframe(df.head(10), use_container_width=True)

    if st.button("🚀 Lancer le matching", type="primary", use_container_width=True):
        with st.spinner("Matching en cours…"):
            try:
                merged, col_nom_used, has_fk = match_taxref(df, taxref)
                final = create_output_table(merged, col_nom_used, has_fk)
            except Exception as e:
                st.error(f"❌ Erreur pendant le matching : {e}")
                st.stop()

        st.success("✅ Matching terminé !")

        # Métriques
        exact_count  = (final["Type de réconciliation"] == "Exact").sum()
        fuzzy_count  = (final["Type de réconciliation"] == "Flou").sum()
        unrec_count  = (final["Type de réconciliation"] == "Non réconcilié").sum()
        coher_count  = (final["Validation classification"] == "✅ Cohérent").sum()
        incoher_count = (final["Validation classification"] == "⚠️ Incohérent").sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Correspondances exactes", exact_count)
        m2.metric("Correspondances floues",  fuzzy_count)
        m3.metric("Non réconciliés",         unrec_count)

        if "Classification" in df.columns:
            m4, m5, _ = st.columns(3)
            m4.metric("Classification cohérente",    coher_count)
            m5.metric("Classification incohérente",  incoher_count)

        st.subheader("Résultats")
        st.dataframe(final, use_container_width=True, height=400)

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
