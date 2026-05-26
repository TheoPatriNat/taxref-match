import streamlit as st
import pandas as pd
from thefuzz import process
from collections import defaultdict
import io

st.set_page_config(page_title="TaxRef Match", page_icon="🌿", layout="wide")

st.title("🌿 TaxRef Match — v18")
st.markdown(
    "Outil de réconciliation de noms d'espèces avec le référentiel **TaxRef v18**. "
    "Uploadez vos deux fichiers, lancez le matching, puis téléchargez le résultat."
)
st.divider()

# ─────────────────────────────────────────────
# README embarqué
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
| `fk` | Non | Clé primaire libre — identifiant que vous choisissez |
| `nom_cite` | **Oui** | Nom d'espèce à réconcilier avec TaxRef |
| `Classification` | Non | Rang taxonomique supérieur pour valider le match |

---

### Étape 2 — Remplir le template

- **`fk`** : laissez vide ou remplissez avec vos propres identifiants. Conservée telle quelle dans le résultat.
- **`nom_cite`** : saisissez les noms d'espèces tels que vous les avez (avec ou sans auteur).
- **`Classification`** : optionnel. N'importe quel rang supérieur : règne, phylum, classe, ordre, famille, sous-famille. Ex : `Animalia`, `Actinopterygii`, `Syngnathidae`.

---

### Étape 3 — Lancer le matching

1. Uploadez **TAXREFv18.txt** via le premier bouton
2. Uploadez votre **fichier Excel** via le second bouton
3. Cliquez sur **"Lancer le matching"**
4. Une barre de progression s'affiche pendant le fuzzy matching
5. Des métriques résument les résultats

---

### Étape 4 — Télécharger le résultat

Cliquez sur **"Télécharger le résultat (.xlsx)"**.

| Colonne | Description |
|---|---|
| `fk` | Votre clé primaire, inchangée |
| `Nom cité` | Le nom tel que saisi |
| `CD_NOM` | Identifiant TaxRef du taxon trouvé |
| `CD_REF` | Identifiant TaxRef du taxon valide |
| `Nom complet` | Nom complet TaxRef avec auteur |
| `Nom valide` | Nom valide selon TaxRef |
| `Similarité` | Score de correspondance (100 = exact, 80–99 = flou) |
| `Niveau taxonomique` | Rang TaxRef (ES, SSES, GN…) |
| `Type de réconciliation` | Exact, Flou ou Non réconcilié |
| `Validation classification` | ✅ Cohérent / ⚠️ Incohérent / — |

---

## II. Comment fonctionne l'outil

### 1. Nettoyage des noms

Avant tout matching, chaque nom est normalisé :
- Suppression de l'auteur et de l'année
- Conservation uniquement du **Genre + épithète spécifique**
- Mise en minuscules

### 2. Match exact

Jointure directe entre le nom nettoyé et `LB_NOM` de TaxRef.
Si trouvé : score = 100, type = `Exact`.

### 3. Match flou

Pour les noms non trouvés en exact, deux contraintes :

**Contrainte A — Genre identique**
Seules les entrées TaxRef du même genre sont candidates.
Évite les faux positifs comme `Syngnathus taenionotus` → `Notus Fieber`.

**Contrainte B — Seuil à 80**
Si le meilleur score est < 80 → `Non réconcilié`.

### 4. Validation par la Classification

La valeur saisie est comparée aux colonnes hiérarchiques TaxRef :
`REGNE`, `PHYLUM`, `CLASSE`, `ORDRE`, `FAMILLE`, `SOUS_FAMILLE`.

- **✅ Cohérent** : correspondance trouvée
- **⚠️ Incohérent** : aucune correspondance → vérification manuelle recommandée
- **—** : colonne non renseignée

### 5. Colonnes TaxRef utilisées

`LB_NOM`, `CD_NOM`, `CD_REF`, `RANG`, `FAMILLE`, `NOM_COMPLET`, `NOM_VALIDE`,
`REGNE`, `PHYLUM`, `CLASSE`, `ORDRE`, `SOUS_FAMILLE`
"""

# ─────────────────────────────────────────────
# Constantes
# ─────────────────────────────────────────────
RANK_COLS = ["REGNE", "PHYLUM", "CLASSE", "ORDRE", "FAMILLE", "SOUS_FAMILLE"]
TAXREF_COLS_NEEDED = [
    "LB_NOM", "CD_NOM", "CD_REF", "RANG",
    "FAMILLE", "NOM_COMPLET", "NOM_VALIDE",
] + RANK_COLS


# ─────────────────────────────────────────────
# Fonctions utilitaires
# ─────────────────────────────────────────────
def clean_name(name):
    if pd.isna(name):
        return ""
    parts = str(name).strip().split()
    if len(parts) >= 2:
        return (parts[0] + " " + parts[1]).lower()
    return " ".join(parts).lower()


def get_template_bytes():
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame(columns=["fk", "nom_cite", "Classification"]).to_excel(w, index=False)
    return buf.getvalue()


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    return buf.getvalue()


# ─────────────────────────────────────────────
# Chargement TaxRef
# ─────────────────────────────────────────────
@st.cache_data(show_spinner="Chargement du TaxRef en mémoire… (une seule fois par session)")
def load_taxref(file_bytes):
    taxref = pd.read_csv(
        io.BytesIO(file_bytes),
        sep="\t",
        quotechar='"',
        dtype=str,
        low_memory=False,
        usecols=lambda c: c in TAXREF_COLS_NEEDED,
    )
    if "LB_NOM" not in taxref.columns:
        raise ValueError("Le fichier TaxRef ne contient pas la colonne 'LB_NOM'.")

    # Remplir les colonnes manquantes
    for col in TAXREF_COLS_NEEDED:
        if col not in taxref.columns:
            taxref[col] = ""
        else:
            taxref[col] = taxref[col].fillna("")

    # Nettoyage des noms
    taxref["nom_clean"] = taxref["LB_NOM"].apply(clean_name)

    # ← CORRECTION CLEF : dédupliquer sur nom_clean (garde le premier CD_REF = taxon valide)
    taxref = taxref.sort_values("CD_REF").drop_duplicates(subset=["nom_clean"], keep="first")

    return taxref.reset_index(drop=True)


# ─────────────────────────────────────────────
# Fuzzy matching
# ─────────────────────────────────────────────
def build_genus_index(taxref):
    idx = defaultdict(list)
    for name in taxref["nom_clean"]:
        parts = name.split()
        if parts:
            idx[parts[0]].append(name)
    return idx


def fuzzy_with_constraint(name, genus_index, threshold=80):
    parts = name.split()
    if not parts:
        return None, 0
    candidates = genus_index.get(parts[0], [])
    if not candidates:
        return None, 0
    res = process.extractOne(name, candidates)
    if res is None or res[1] < threshold:
        return None, 0
    return res[0], res[1]


# ─────────────────────────────────────────────
# Validation classification
# ─────────────────────────────────────────────
def validate_classification(user_val, taxref_row):
    if pd.isna(user_val) or str(user_val).strip() == "":
        return "—"
    user_clean = str(user_val).strip().lower()
    for col in RANK_COLS:
        if str(taxref_row.get(col, "")).strip().lower() == user_clean:
            return "✅ Cohérent"
    return "⚠️ Incohérent"


# ─────────────────────────────────────────────
# Matching principal
# ─────────────────────────────────────────────
def match_taxref(df, taxref):
    df = df.copy().reset_index(drop=True)

    col_nom = "nom_cite" if "nom_cite" in df.columns else "Nom_cite"
    has_fk      = "fk" in df.columns
    has_classif = "Classification" in df.columns

    df["nom_clean"] = df[col_nom].apply(clean_name)

    # Colonnes de sortie TaxRef
    out_cols = ["nom_clean", "CD_NOM", "CD_REF", "RANG",
            "NOM_COMPLET", "NOM_VALIDE"] + RANK_COLS

    # ── 1. Match exact ────────────────────────────────────────────────
    taxref_lookup = taxref[out_cols].copy()
    taxref_lookup = taxref_lookup.loc[:, ~taxref_lookup.columns.duplicated()]

    # Séparer les lignes qui matchent et celles qui ne matchent pas
    df["_idx"] = df.index
    exact = df.merge(taxref_lookup, on="nom_clean", how="inner")
    exact["Type"]      = "Exact"
    exact["Similarité"] = 100.0

    matched_names = set(exact["nom_clean"].unique())
    unmatched_df  = df[~df["nom_clean"].isin(matched_names)].copy()

    # ── 2. Fuzzy matching ─────────────────────────────────────────────
    fuzzy_rows = []

    if not unmatched_df.empty:
        unique_unmatched = [n for n in unmatched_df["nom_clean"].unique() if n.strip()]
        genus_index      = build_genus_index(taxref)

        progress_bar = st.progress(0, text="Fuzzy matching en cours…")
        fuzzy_name_map = {}  # nom_clean → (match_clean, score)

        for i, name in enumerate(unique_unmatched):
            match, score = fuzzy_with_constraint(name, genus_index, threshold=80)
            fuzzy_name_map[name] = (match, score)
            progress_bar.progress((i + 1) / len(unique_unmatched),
                                   text=f"Fuzzy matching… {i+1}/{len(unique_unmatched)}")
        progress_bar.empty()

        taxref_by_name = taxref.set_index("nom_clean")

        for _, row in unmatched_df.iterrows():
            name  = row["nom_clean"]
            match, score = fuzzy_name_map.get(name, (None, 0))
            new_row = row.to_dict()
            if match and match in taxref_by_name.index:
                tax = taxref_by_name.loc[match]
                # .loc peut retourner un DataFrame si doublons résiduels → on prend la 1re ligne
                if isinstance(tax, pd.DataFrame):
                    tax = tax.iloc[0]
                new_row["CD_NOM"]      = str(tax.get("CD_NOM", ""))
                new_row["CD_REF"]      = str(tax.get("CD_REF", ""))
                new_row["RANG"]        = str(tax.get("RANG", ""))
                new_row["FAMILLE"]     = str(tax.get("FAMILLE", ""))
                new_row["NOM_COMPLET"] = str(tax.get("NOM_COMPLET", ""))
                new_row["NOM_VALIDE"]  = str(tax.get("NOM_VALIDE", ""))
                for rc in RANK_COLS:
                    new_row[rc] = str(tax.get(rc, ""))
                new_row["Type"]       = "Flou"
                new_row["Similarité"] = float(score)
            else:
                for c in ["CD_NOM", "CD_REF", "RANG", "FAMILLE",
                          "NOM_COMPLET", "NOM_VALIDE"] + RANK_COLS:
                    new_row[c] = ""
                new_row["Type"]       = "Non réconcilié"
                new_row["Similarité"] = 0.0
            fuzzy_rows.append(new_row)

    # ── 3. Assemblage ─────────────────────────────────────────────────
    frames = [exact]
    if fuzzy_rows:
        frames.append(pd.DataFrame(fuzzy_rows))

    merged = pd.concat(frames, ignore_index=True)

    # Remettre dans l'ordre original
    merged = merged.sort_values("_idx").drop(columns=["_idx"]).reset_index(drop=True)

    # Nettoyage types
    merged["Similarité"] = pd.to_numeric(merged["Similarité"], errors="coerce").fillna(0.0)
    merged["Type"]       = merged["Type"].fillna("Non réconcilié")
    for col in ["CD_NOM", "CD_REF", "RANG", "FAMILLE", "NOM_COMPLET", "NOM_VALIDE"] + RANK_COLS:
        if col in merged.columns:
            merged[col] = merged[col].fillna("")

    # ── 4. Validation classification ──────────────────────────────────
    if has_classif:
        merged["Validation classification"] = merged.apply(
            lambda r: validate_classification(r.get("Classification", ""), r), axis=1
        )
    else:
        merged["Validation classification"] = "—"

    return merged, col_nom, has_fk


# ─────────────────────────────────────────────
# Table de sortie
# ─────────────────────────────────────────────
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
        "**Format attendu :**\n\n"
        "- `fk` : clé primaire *(optionnel)*\n"
        "- `nom_cite` : nom d'espèce *(obligatoire)*\n"
        "- `Classification` : rang supérieur *(optionnel)*"
    )

col1, col2 = st.columns(2)
with col1:
    st.subheader("1 · Fichier TaxRef v18")
    taxref_file = st.file_uploader("Uploader TAXREFv18.txt", type=["txt"])
with col2:
    st.subheader("2 · Fichier occurrences")
    input_file = st.file_uploader("Uploader votre fichier Excel (.xlsx)", type=["xlsx"])

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

    col_nom = "nom_cite" if "nom_cite" in df.columns else ("Nom_cite" if "Nom_cite" in df.columns else None)
    if col_nom is None:
        st.error("❌ Le fichier Excel doit contenir une colonne **`nom_cite`**.")
        st.stop()

    st.success(f"✅ {len(df)} occurrences chargées — {len(taxref)} entrées TaxRef en mémoire.")

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
        try:
            merged, col_nom_used, has_fk = match_taxref(df, taxref)
            final = create_output_table(merged, col_nom_used, has_fk)
        except Exception as e:
            st.error(f"❌ Erreur pendant le matching : {e}")
            st.stop()

        st.success("✅ Matching terminé !")

        exact_count   = (final["Type de réconciliation"] == "Exact").sum()
        fuzzy_count   = (final["Type de réconciliation"] == "Flou").sum()
        unrec_count   = (final["Type de réconciliation"] == "Non réconcilié").sum()

        m1, m2, m3 = st.columns(3)
        m1.metric("Correspondances exactes", exact_count)
        m2.metric("Correspondances floues",  fuzzy_count)
        m3.metric("Non réconciliés",         unrec_count)

        if "Classification" in df.columns:
            coher_count   = (final["Validation classification"] == "✅ Cohérent").sum()
            incoher_count = (final["Validation classification"] == "⚠️ Incohérent").sum()
            m4, m5, _ = st.columns(3)
            m4.metric("Classification cohérente",   coher_count)
            m5.metric("Classification incohérente", incoher_count)

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
