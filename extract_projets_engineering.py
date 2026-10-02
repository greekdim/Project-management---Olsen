"""
Extraction Odoo des projets "PRO (LIG)" + "Engineering" pour les réunions
engineering.

- Garde les projets portant LES DEUX étiquettes (PRO (LIG) ET Engineering).
- Exclut les étapes Annulé / Cloturé / Autres / Template / Canceled
  (comparaison exacte, insensible à la casse et aux accents).
- Trie par étape (STAGE_ORDER) -> chef de projet -> numéro de projet.
- Génère un Excel : "Projets actifs", "Synthèse", "Projets archivés".

Identifiants : variables d'environnement ODOO_URL, ODOO_DB, ODOO_USER,
ODOO_PASSWORD, ou fichier .env à côté du script (voir .env.example).
"""
import datetime
import getpass
import os
import re
import sys
import traceback
import unicodedata
import xmlrpc.client

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- Répertoire de travail robuste (fonctionne en double-clic) ---
if getattr(sys, 'frozen', False):
    script_dir = os.path.dirname(sys.executable)
else:
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        script_dir = os.path.abspath(".")
os.chdir(script_dir)


# ==================================================================
#  PARAMÈTRES
# ==================================================================
def load_dotenv(path=".env"):
    """Charge un fichier .env minimal (KEY=VALUE) sans écraser l'environnement."""
    if not os.path.exists(path):
        return
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or "=" not in line:
                continue
            key, _, value = line.partition("=")
            os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


load_dotenv()
URL = os.environ.get("ODOO_URL", "https://olsen-engineering.odoo.com")
DB = os.environ.get("ODOO_DB", "mynalios-olsen-main-7388485")
USERNAME = os.environ.get("ODOO_USER", "")
PASSWORD = os.environ.get("ODOO_PASSWORD", "")

EXCLUSIONS = ["Annulé", "Cloturé", "Autres", "Template", "Canceled"]
REQUIRED_TAGS = ["PRO (LIG)", "Engineering"]

# Ordre de tri voulu pour les étapes (stage_id)
STAGE_ORDER = [
    "Facture finale",
    "Réception et CE",
    "Livraison et montage",
    "Atelier",
    "Approvisionnement",
    "Technique / Étude",
    "Kick-off",
    "Nouveau",
]

OUTPUT_FILE = "project_review_odoo.xlsx"
SHEET_MAIN = "Projets actifs"
SHEET_SUMMARY = "Synthèse"
SHEET_ARCHIVE = "Projets archivés"

COLUMNS = [
    "id_odoo", "numero_projet", "description", "partner_id", "stage_id", "chef_de_projet",
    "date_debut", "date_fin",
    "sujets_a_discuter", "actions_a_entreprendre", "responsable_action",
    "date_limite", "statut", "derniere_review", "commentaires",
]
DATE_COLUMNS = ["date_debut", "date_fin", "date_limite", "derniere_review"]

HEADER_MAP = {
    "id_odoo": "ID Odoo",
    "numero_projet": "N° Projet",
    "description": "Projet",
    "partner_id": "Client",
    "stage_id": "Étape",
    "chef_de_projet": "Chef de projet",
    "date_debut": "Date début",
    "date_fin": "Date fin",
    "sujets_a_discuter": "Sujets à discuter",
    "actions_a_entreprendre": "Actions à entreprendre",
    "responsable_action": "Responsable action",
    "date_limite": "Date limite",
    "statut": "Statut",
    "derniere_review": "Dernière review",
    "commentaires": "Commentaires",
}
ARCHIVE_COLUMNS = COLUMNS + ["date_archivage"]
ARCHIVE_HEADER_MAP = dict(HEADER_MAP, date_archivage="Archivé le")
REVERSE_HEADER_MAP = {v: k for k, v in ARCHIVE_HEADER_MAP.items()}

# Numéro de projet attendu en tête du nom : lettre + 2 chiffres + "-" + 5 chiffres
PROJECT_NUMBER_RE = re.compile(r'^([A-Za-z]\d{2}-\d{5})\s*(.*)$')


def normalize(text):
    """Minuscule, sans accents, espaces nettoyés : pour comparer des libellés."""
    text = unicodedata.normalize("NFKD", text or "")
    text = "".join(c for c in text if not unicodedata.combining(c))
    return " ".join(text.lower().split())


STAGE_PRIORITY = {normalize(name): i for i, name in enumerate(STAGE_ORDER)}
EXCLUSIONS_NORMALIZED = {normalize(e) for e in EXCLUSIONS}


def is_excluded_stage(stage_name):
    return normalize(stage_name) in EXCLUSIONS_NORMALIZED


def stage_priority(stage_name):
    """Position de tri d'une étape ; les étapes inconnues vont à la fin."""
    return STAGE_PRIORITY.get(normalize(stage_name), len(STAGE_ORDER))


def split_project_name(display_name):
    """Sépare le nom Odoo en (numéro, description)."""
    if not display_name:
        return "", ""
    match = PROJECT_NUMBER_RE.match(display_name.strip())
    if match:
        return match.group(1), match.group(2).strip(" -:–\t")
    return "", display_name.strip()


def to_date(value):
    """str Odoo / Timestamp / date / None / NaN -> datetime.date ou None."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, str):
        v = value.strip()
        if not v or v.lower() == "nan":
            return None
        try:
            return datetime.datetime.strptime(v[:10], "%Y-%m-%d").date()
        except ValueError:
            return None
    if isinstance(value, datetime.datetime):  # inclut pd.Timestamp
        return value.date()
    if isinstance(value, datetime.date):
        return value
    return None


def to_text(value):
    """Valeur texte d'un ancien fichier Excel (NaN/NaT -> '')."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value)


# ==================================================================
#  LECTURE DE L'ANCIEN FICHIER (pour alimenter l'archive)
# ==================================================================
def load_previous_notes(path):
    """Retourne (old_main, old_archive) : dicts {id_odoo: {colonnes...}}."""
    if not os.path.exists(path):
        return {}, {}

    def _load_sheet(sheet_name):
        try:
            df = pd.read_excel(path, sheet_name=sheet_name)
        except Exception:
            return {}
        df = df.rename(columns=REVERSE_HEADER_MAP)
        if "id_odoo" not in df.columns:
            return {}
        result = {}
        for _, row in df.iterrows():
            try:
                pid = int(row["id_odoo"])
            except (ValueError, TypeError):
                continue
            result[pid] = {col: row[col] for col in df.columns if col != "id_odoo"}
        return result

    return _load_sheet(SHEET_MAIN), _load_sheet(SHEET_ARCHIVE)


# ==================================================================
#  CONSTRUCTION DES LIGNES
# ==================================================================
def build_rows(projects, old_main, old_archive):
    """Retourne (main_rows, archive_rows).

    La feuille principale repart de zéro à chaque extraction. L'archive garde
    le dernier état connu (avec notes) des projets qui sortent de la liste."""
    today = datetime.date.today()
    current_ids = {p["id"] for p in projects}

    main_rows = []
    for p in projects:
        main_rows.append({
            "id_odoo": p["id"],
            "numero_projet": p["numero_projet"],
            "description": p["description"],
            "partner_id": p["partner_id"],
            "stage_id": p["stage_id"],
            "chef_de_projet": p["manager"] or "",
            "date_debut": to_date(p["date_start"]),
            "date_fin": to_date(p["date_end"]),
            "sujets_a_discuter": "",
            "actions_a_entreprendre": "",
            "responsable_action": p["manager"] or "",
            "date_limite": None,
            "statut": "",
            "derniere_review": None,
            "commentaires": "",
        })

    # Tri : étape (ordre imposé, inconnues à la fin par ordre alphabétique)
    # -> chef de projet -> numéro de projet
    main_rows.sort(key=lambda r: (
        stage_priority(r["stage_id"]),
        normalize(r["stage_id"]) if stage_priority(r["stage_id"]) == len(STAGE_ORDER) else "",
        normalize(r["chef_de_projet"]),
        r["numero_projet"] or "",
    ))

    archive_rows = []
    for pid in (set(old_main) | set(old_archive)) - current_ids:
        if pid in old_main:
            source, date_archivage = old_main[pid], today
        else:
            source = old_archive[pid]
            date_archivage = to_date(source.get("date_archivage")) or today
        row = {col: (to_date(source.get(col)) if col in DATE_COLUMNS else to_text(source.get(col)))
               for col in COLUMNS if col != "id_odoo"}
        row["id_odoo"] = pid
        row["date_archivage"] = date_archivage
        archive_rows.append(row)
    archive_rows.sort(key=lambda r: r["date_archivage"] or today, reverse=True)

    return main_rows, archive_rows


def build_summary(main_rows):
    """Tableau (étape x chef de projet) : nombre de projets, dans l'ordre des étapes."""
    if not main_rows:
        return [], []
    managers = sorted({r["chef_de_projet"] or "(non assigné)" for r in main_rows}, key=normalize)
    stages = []
    for r in main_rows:  # main_rows est déjà trié par étape
        if r["stage_id"] not in stages:
            stages.append(r["stage_id"])
    table = []
    for stage in stages:
        counts = [sum(1 for r in main_rows
                      if r["stage_id"] == stage and (r["chef_de_projet"] or "(non assigné)") == m)
                  for m in managers]
        table.append([stage or "(sans étape)"] + counts + [sum(counts)])
    totals = [sum(row[i] for row in table) for i in range(1, len(managers) + 2)]
    table.append(["Total"] + totals)
    return ["Étape"] + managers + ["Total"], table


# ==================================================================
#  ÉCRITURE DU FICHIER EXCEL
# ==================================================================
THIN_BORDER = Border(*([Side(style="thin", color="D9D9D9")] * 4))
STAGE_BREAK_BORDER = Border(left=Side(style="thin", color="D9D9D9"),
                            right=Side(style="thin", color="D9D9D9"),
                            bottom=Side(style="thin", color="D9D9D9"),
                            top=Side(style="medium", color="1F4E78"))
BAND_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)


def write_workbook(main_rows, archive_rows, output_file):
    wb = Workbook()
    _write_sheet(wb.active, SHEET_MAIN, COLUMNS, HEADER_MAP, main_rows, is_main=True)
    _write_summary(wb.create_sheet(SHEET_SUMMARY), main_rows)
    _write_sheet(wb.create_sheet(SHEET_ARCHIVE), SHEET_ARCHIVE, ARCHIVE_COLUMNS,
                 ARCHIVE_HEADER_MAP, archive_rows, is_main=False)
    wb.save(output_file)


def _write_summary(ws, main_rows):
    ws.title = SHEET_SUMMARY
    header, table = build_summary(main_rows)
    for c, label in enumerate(header, start=1):
        cell = ws.cell(row=1, column=c, value=label)
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r, line in enumerate(table, start=2):
        for c, value in enumerate(line, start=1):
            cell = ws.cell(row=r, column=c, value=value)
            cell.border = THIN_BORDER
            if c > 1:
                cell.alignment = Alignment(horizontal="center")
            if r == len(table) + 1:
                cell.font = Font(bold=True)
    ws.column_dimensions["A"].width = 24
    for c in range(2, len(header) + 1):
        ws.column_dimensions[get_column_letter(c)].width = 18
    ws.freeze_panes = "B2"


def _write_sheet(ws, title, columns, header_map, rows, is_main):
    ws.title = title

    for col_idx, col_key in enumerate(columns, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header_map[col_key])
        cell.font, cell.fill = HEADER_FONT, HEADER_FILL
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 45

    previous_stage = None
    for row_idx, row in enumerate(rows, start=2):
        band = (row_idx % 2 == 0)
        ws.row_dimensions[row_idx].height = 45
        # Trait épais à chaque changement d'étape (feuille principale)
        stage_changed = is_main and row.get("stage_id") != previous_stage
        previous_stage = row.get("stage_id")
        for col_idx, col_key in enumerate(columns, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row.get(col_key, ""))
            cell.border = STAGE_BREAK_BORDER if stage_changed else THIN_BORDER
            if band:
                cell.fill = BAND_FILL
            if col_key in DATE_COLUMNS or col_key == "date_archivage":
                cell.number_format = "DD/MM/YYYY"
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if col_key == "stage_id":
                cell.font = Font(bold=True)

    last_row = max(len(rows) + 1, 2)
    last_col_letter = get_column_letter(len(columns))

    width_map = {
        "id_odoo": 9, "numero_projet": 12, "description": 30, "partner_id": 22,
        "stage_id": 18, "chef_de_projet": 18, "date_debut": 12, "date_fin": 12,
        "sujets_a_discuter": 32, "actions_a_entreprendre": 32,
        "responsable_action": 16, "date_limite": 12, "statut": 14,
        "derniere_review": 14, "commentaires": 28, "date_archivage": 12,
    }
    for col_idx, col_key in enumerate(columns, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width_map.get(col_key, 15)
    ws.column_dimensions[get_column_letter(columns.index("id_odoo") + 1)].hidden = True

    ws.auto_filter.ref = f"A1:{last_col_letter}{last_row}"
    ws.freeze_panes = f"{get_column_letter(columns.index('description') + 2)}2"

    if not is_main or not rows:
        return

    # Action en retard et pas terminée -> ligne en rouge
    statut = get_column_letter(columns.index("statut") + 1)
    limite = get_column_letter(columns.index("date_limite") + 1)
    formula = (f'AND(${limite}2<>"",${limite}2<TODAY(),'
               f'LOWER(TRIM(${statut}2))<>"fait")')
    ws.conditional_formatting.add(
        f"A2:{last_col_letter}{last_row}",
        FormulaRule(formula=[formula],
                    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")),
    )


# ==================================================================
#  ODOO
# ==================================================================
def fetch_projects():
    global USERNAME, PASSWORD
    if not USERNAME or not PASSWORD:
        if not (sys.stdin and sys.stdin.isatty()):
            raise Exception("Identifiants manquants : définissez ODOO_USER et ODOO_PASSWORD "
                            "(variables d'environnement ou fichier .env, voir .env.example).")
        print("ℹ️  Aucun identifiant trouvé (.env / variables d'environnement) : saisie manuelle.")
        USERNAME = USERNAME or input("Identifiant Odoo (email) : ").strip()
        PASSWORD = PASSWORD or getpass.getpass("Mot de passe (invisible à la saisie) : ")

    print("🔌 Connexion à Odoo...")
    common = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/common", allow_none=True)
    uid = common.authenticate(DB, USERNAME, PASSWORD, {})
    if not uid:
        raise Exception("Authentification échouée — vérifiez vos identifiants.")
    models = xmlrpc.client.ServerProxy(f"{URL}/xmlrpc/2/object", allow_none=True)

    def call(model, method, args, kwargs=None):
        return models.execute_kw(DB, uid, PASSWORD, model, method, args, kwargs or {})

    print("✅ Connecté.\n⏳ Résolution des tags requis...")
    tag_records = call('project.tags', 'search_read',
                       [[['name', 'in', REQUIRED_TAGS]]], {'fields': ['id', 'name']})
    tag_ids_by_name = {t['name']: t['id'] for t in tag_records}
    missing = set(REQUIRED_TAGS) - set(tag_ids_by_name)
    if missing:
        raise Exception(f"Tags introuvables dans Odoo : {missing}")

    print("⏳ Récupération des étapes à exclure...")
    stages = call('project.project.stage', 'search_read', [[]], {
        'fields': ['id', 'name'], 'context': {'active_test': False}})
    excluded_stage_ids = [s['id'] for s in stages if is_excluded_stage(s['name'])]

    print("⏳ Récupération des projets...")
    # Une condition par tag => le projet doit porter TOUS les tags requis (ET).
    domain = [['tag_ids', 'in', [tid]] for tid in tag_ids_by_name.values()]
    if excluded_stage_ids:
        domain.append(['stage_id', 'not in', excluded_stage_ids])
    raw_projects = call('project.project', 'search_read', [domain], {
        'fields': ['stage_id', 'display_name', 'partner_id', 'user_id', 'date_start', 'date']})
    print(f"✅ {len(raw_projects)} projets récupérés.")

    projects = []
    for p in raw_projects:
        stage = p['stage_id'][1] if p['stage_id'] else ''
        if is_excluded_stage(stage):  # filet de sécurité côté client
            continue
        numero, description = split_project_name(p["display_name"])
        projects.append({
            "id": p["id"],
            "numero_projet": numero,
            "description": description,
            "partner_id": p["partner_id"][1] if p["partner_id"] else "",
            "stage_id": stage,
            "manager": p["user_id"][1] if p["user_id"] else "",
            "date_start": p["date_start"],
            "date_end": p["date"],
        })
    return projects


def main():
    projects = fetch_projects()

    print("📂 Lecture de l'ancien fichier (si présent)...")
    old_main, old_archive = load_previous_notes(OUTPUT_FILE)

    unknown = {p["stage_id"] for p in projects if stage_priority(p["stage_id"]) == len(STAGE_ORDER)}
    if unknown:
        print(f"⚠️  Étape(s) absente(s) de STAGE_ORDER (placées en fin de liste) : {sorted(unknown)}")
        print("   -> Corrigez STAGE_ORDER en haut du script si l'orthographe diffère.")

    main_rows, archive_rows = build_rows(projects, old_main, old_archive)

    print("💾 Écriture du fichier Excel...")
    try:
        write_workbook(main_rows, archive_rows, OUTPUT_FILE)
        print(f"\n✅ Export terminé : {OUTPUT_FILE}")
        print(f"   - {len(main_rows)} projets actifs")
        print(f"   - {len(archive_rows)} projets archivés (clôturés / sortis des critères)")
    except PermissionError:
        print(f"\n❌ Fichier '{OUTPUT_FILE}' déjà ouvert. Fermez-le et réessayez.")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\n❌ Erreur : {e}")
        traceback.print_exc()
    finally:
        if sys.stdin and sys.stdin.isatty():
            input("\nAppuyez sur Entrée pour fermer...")
