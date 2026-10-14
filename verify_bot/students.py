import unicodedata

import openpyxl
from dataclasses import dataclass

# Column layout of the "Liste Affichage" exports (header row, Matricule = column E).
EXPECTED_COLUMNS = {
    "numero": 0,       # A  N°
    "palier": 1,       # B  Palier
    "specialite": 2,   # C  Spécialité
    "section": 3,      # D  Section
    "matricule": 4,    # E  Matricule
    "nom": 5,          # F  Nom
    "prenom": 6,       # G  Prénom
    "etat": 7,         # H  Etat
    "groupe_td": 8,    # I  Groupe TD
    "groupe_tp": 9,    # J  Groupe TP
}

FIELD_ALIASES = {
    "numero": {"n°", "n", "no", "num", "numero"},
    "palier": {"palier"},
    "specialite": {"specialite", "specialites", "filiere"},
    "section": {"section"},
    "matricule": {"matricule", "matricules"},
    "nom": {"nom", "noms", "nom de famille"},
    "prenom": {"prenom", "prenoms", "prenom(s)"},
    "etat": {"etat", "statut"},
    "groupe_td": {"groupe td", "td", "groupe_td", "grp td"},
    "groupe_tp": {"groupe tp", "tp", "groupe_tp", "grp tp"},
}


@dataclass
class Student:
    numero: str
    palier: str
    specialite: str
    section: str
    matricule: str
    nom: str
    prenom: str
    etat: str
    groupe_td: str
    groupe_tp: str

    @property
    def full_name(self) -> str:
        return f"{self.nom} {self.prenom}".strip()

    @property
    def is_sii(self) -> bool:
        return (self.specialite or "").strip().upper() == "SII"

    @property
    def is_admis(self) -> bool:
        return (self.etat or "").strip().upper() == "ADM"


def _norm(value) -> str:
    if value is None:
        return ""
    text = unicodedata.normalize("NFKD", str(value))
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.replace("_", " ").replace(".", " ").replace(":", " ")
    return " ".join(text.split()).strip().lower()


def _columns_by_name(row) -> dict[str, int]:
    columns: dict[str, int] = {}
    for idx, cell in enumerate(row):
        name = _norm(cell)
        if not name:
            continue
        for field, aliases in FIELD_ALIASES.items():
            if name in aliases and field not in columns:
                columns[field] = idx
    return columns


def _find_header(rows) -> tuple[int, dict[str, int]]:
    expected_idx = EXPECTED_COLUMNS["matricule"]
    for i, row in enumerate(rows):
        if len(row) > expected_idx and _norm(row[expected_idx]) == "matricule":
            return i, dict(EXPECTED_COLUMNS)

    for i, row in enumerate(rows):
        columns = _columns_by_name(row)
        if "matricule" in columns:
            return i, columns

    raise ValueError(
        "Could not find the header row: expected 'Matricule' in column E "
        "(row with headers N°, Palier, Spécialité, Section, Matricule, ...)."
    )


def load_students(path: str) -> tuple[dict[str, Student], int]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header_row, columns = _find_header(rows)

    def cell(row, field) -> str:
        idx = columns.get(field)
        if idx is None or idx >= len(row):
            return ""
        value = row[idx]
        return "" if value is None else str(value).strip()

    students: dict[str, Student] = {}
    skipped = 0
    for row in rows[header_row + 1:]:
        matricule = cell(row, "matricule")
        if not matricule:
            continue
        matricule = matricule.upper()
        if matricule in students:
            skipped += 1
            continue

        numero_idx = columns.get("numero")
        numero_raw = row[numero_idx] if numero_idx is not None and numero_idx < len(row) else None
        if isinstance(numero_raw, (int, float)):
            numero = str(int(numero_raw))
        else:
            numero = cell(row, "numero")

        students[matricule] = Student(
            numero=numero,
            palier=cell(row, "palier"),
            specialite=cell(row, "specialite"),
            section=cell(row, "section"),
            matricule=matricule,
            nom=cell(row, "nom"),
            prenom=cell(row, "prenom"),
            etat=cell(row, "etat"),
            groupe_td=_num(cell(row, "groupe_td")),
            groupe_tp=_num(cell(row, "groupe_tp")),
        )
    if not students:
        raise ValueError(f"No students found under the header row in {path}.")
    return students, skipped


def _num(value) -> str:
    if value is None:
        return ""
    try:
        number = float(str(value).strip())
    except (TypeError, ValueError):
        return str(value).strip()
    if number.is_integer():
        return str(int(number))
    return str(value).strip()
