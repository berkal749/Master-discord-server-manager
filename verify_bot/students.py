import csv
import re
import unicodedata
from decimal import Decimal

import openpyxl
from dataclasses import dataclass
from pathlib import Path

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
    "etat": {"etat", "statut", "etat civil"},
    "groupe_td": {"groupe td", "td", "groupe_td", "grp td", "groupe"},
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
        return (self.specialite or "").strip().upper() == "BIOINFO"

    @property
    def is_bioinfo(self) -> bool:
        return (self.specialite or "").strip().upper() == "BIOINFO"

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


def _norm_cell(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def _load_from_csv(path: Path) -> tuple[dict[str, Student], int]:
    students: dict[str, Student] = {}
    skipped = 0
    rows = []
    for enc in ("utf-8-sig", "utf-8", "latin-1", "cp1252"):
        try:
            with open(path, encoding=enc) as f:
                reader = csv.reader(f)
                rows = list(reader)
            break
        except Exception:
            rows = []
            continue
    if not rows:
        return students, skipped

    # Find header row
    header_idx = -1
    columns: dict[str, int] = {}
    for i, row in enumerate(rows):
        row_norm = [_norm(row[c]) if c < len(row) else "" for c in range(len(row))]
        # look for matricule
        if any("matricule" in r for r in row_norm):
            header_idx = i
            # map
            for idx, cellv in enumerate(row):
                n = _norm(cellv)
                for field, aliases in FIELD_ALIASES.items():
                    if n in aliases and field not in columns:
                        columns[field] = idx
            break
    if header_idx < 0:
        return students, skipped
    if "matricule" not in columns:
        # try fixed positions for this format
        columns = {"matricule": 0, "nom": 1, "prenom": 2, "etat": 3, "groupe_td": 4}

    for row in rows[header_idx + 1:]:
        if not row or len(row) == 0:
            continue
        def c(field):
            idx = columns.get(field)
            if idx is None or idx >= len(row):
                return ""
            return _norm_cell(row[idx])
        matricule_raw = c("matricule")
        if not matricule_raw or matricule_raw.lower().startswith('liste'):
            continue
        matricule = ""
        try:
            s = str(matricule_raw).strip().replace(" ", "").replace("'", "")
            s_up = s.upper().replace(",", ".")
            if "E" in s_up:
                # Parse scientific notation manually to preserve full precision
                # e.g., 2.22231E+11
                parts = s_up.split('E')
                mantissa = parts[0]
                exp = int(parts[1])
                # Remove decimal point from mantissa
                if '.' in mantissa:
                    int_part, frac_part = mantissa.split('.')
                    frac_digits = len(frac_part)
                    num = int_part + frac_part
                    num = num.lstrip('+').lstrip('-')
                    if exp >= frac_digits:
                        matricule = num + '0' * (exp - frac_digits)
                    else:
                        matricule = num[:exp - frac_digits] if (exp - frac_digits) > 0 else num
                else:
                    num = mantissa.lstrip('+').lstrip('-')
                    matricule = num + '0' * exp
                # Clean up sign if any
                if mantissa.startswith('-'):
                    matricule = '-' + matricule
            elif re.match(r'^[0-9]+\.[0-9]+$', s_up):
                # Could be float representation of integer
                d = Decimal(s_up)
                # If it's effectively an integer
                if d == d.to_integral_value():
                    matricule = format(d.quantize(Decimal('1')), 'f')
                else:
                    matricule = re.sub(r'[^0-9]', '', s_up)
            else:
                matricule = re.sub(r'[^0-9]', '', s)
        except Exception:
            matricule = re.sub(r'[^0-9]', '', str(matricule_raw))
        if not matricule.isdigit() or len(matricule) < 8:
            # One more try - if it's still short, might need different parsing
            continue
        if matricule in students:
            skipped += 1
            continue

        nom = c("nom")
        prenom = c("prenom")
        etat = c("etat")
        groupe_td = c("groupe_td") or c("groupe")
        students[matricule] = Student(
            numero="",
            palier="M1",
            specialite="BIOINFO",
            section="",
            matricule=matricule,
            nom=nom,
            prenom=prenom,
            etat=etat,
            groupe_td=_num(groupe_td),
            groupe_tp="",
        )
    return students, skipped


def _load_from_xlsx_groups(path: Path) -> tuple[dict[str, Student], int]:
    students: dict[str, Student] = {}
    skipped = 0
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception:
        return students, skipped

    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        rows = list(ws.iter_rows(values_only=True))
        # find header
        header_idx = -1
        for i, row in enumerate(rows):
            if not row:
                continue
            # look for Matricule in row
            for c in row:
                if c is not None and str(c).strip().lower() == 'matricule':
                    header_idx = i
                    break
            if header_idx >= 0:
                break
        if header_idx < 0 or header_idx >= len(rows):
            continue
        header = rows[header_idx]
        # map columns
        col_map = {}
        for idx, c in enumerate(header):
            if c is None:
                continue
            n = _norm(c)
            if 'matricule' in n and 'matricule' not in col_map:
                col_map['matricule'] = idx
            elif ('nom' in n and 'prenom' not in n) and 'nom' not in col_map:
                col_map['nom'] = idx
            elif 'prenom' in n and 'prenom' not in col_map:
                col_map['prenom'] = idx
            elif ('etat' == n or 'etat' in n) and 'etat' not in col_map:
                col_map['etat'] = idx
            elif ('groupe' in n and 'td' in n) and 'groupe_td' not in col_map:
                col_map['groupe_td'] = idx
            elif n == 'groupe' and 'groupe_td' not in col_map:
                col_map['groupe_td'] = idx

        for row in rows[header_idx + 1:]:
            if not row:
                continue
            def gc(field):
                idx = col_map.get(field)
                if idx is None or idx >= len(row):
                    return ""
                v = row[idx]
                return "" if v is None else str(v).strip()

            mraw = gc('matricule')
            if not mraw:
                continue
            # handle numeric or string
            try:
                # if it's float-like from excel
                if isinstance(row[col_map['matricule']] if 'matricule' in col_map and col_map['matricule'] < len(row) else None, float):
                    matricule = str(int(row[col_map['matricule']]))
                else:
                    s = str(mraw).strip().replace(" ", "")
                    # try int if it's a number
                    # remove .0
                    if re.match(r'^\d+\.0+$', s):
                        matricule = s.split('.')[0]
                    elif re.match(r'^\d+$', s):
                        matricule = s
                    else:
                        # try float
                        try:
                            f = float(s)
                            if f.is_integer():
                                matricule = str(int(f))
                            else:
                                matricule = re.sub(r'[^0-9]', '', s)
                        except Exception:
                            matricule = re.sub(r'[^0-9]', '', s)
            except Exception:
                matricule = re.sub(r'[^0-9]', '', str(mraw))

            if not matricule.isdigit() or len(matricule) < 8:
                continue
            if matricule in students:
                skipped += 1
                continue

            nom = gc('nom')
            prenom = gc('prenom')
            etat = gc('etat')
            groupe_td = gc('groupe_td')
            # if groupe_td empty, try to infer from sheet name
            if not groupe_td and ws_name:
                m = re.search(r'(\d+)', ws_name)
                if m:
                    groupe_td = m.group(1)

            students[matricule] = Student(
                numero="",
                palier="M1",
                specialite="BIOINFO",
                section="",
                matricule=matricule,
                nom=nom,
                prenom=prenom,
                etat=etat,
                groupe_td=_num(groupe_td),
                groupe_tp="",
            )
    return students, skipped


def load_students(path: str) -> tuple[dict[str, Student], int]:
    p = Path(path)
    students: dict[str, Student] = {}
    skipped = 0

    # If it's a CSV file, load directly (merge both groups if it's one of them)
    if p.suffix.lower() == ".csv":
        name = p.name.lower()
        if "group" in name:
            for c in [p.parent / "Group 1.csv", p.parent / "Group 2.csv", Path.cwd() / "Group 1.csv", Path.cwd() / "Group 2.csv", Path(__file__).parent.parent / "Group 1.csv", Path(__file__).parent.parent / "Group 2.csv"]:
                if c.exists():
                    s2, sk2 = _load_from_csv(c)
                    for k, v in s2.items():
                        if k not in students:
                            students[k] = v
                        else:
                            skipped += 1
                    skipped += sk2
            if students:
                return students, skipped
        return _load_from_csv(p)

    # If path points to a specific file that doesn't exist, look for files
    if not p.exists():
        candidates = []
        for base in [p.parent, Path.cwd(), Path(__file__).parent, Path(__file__).parent.parent]:
            candidates.extend([
                base / "ListBioInfo-M1-26-27-Groups.xlsx",
                base / "Group 1.csv",
                base / "Group 2.csv",
            ])
        for c in candidates:
            if c.exists():
                p = c
                break
        if not p.exists():
            raise FileNotFoundError(f"Student file not found: {path}")

    # Try as XLSX (handle multi-sheet group files)
    if p.suffix.lower() in (".xlsx", ".xlsm", ".xltx", ".xltm"):
        try:
            # Try multi-sheet groups first
            s_multi, sk_multi = _load_from_xlsx_groups(p)
            if s_multi:
                return s_multi, sk_multi
        except Exception:
            pass
        try:
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
            if students:
                return students, skipped
        except Exception:
            pass

    # Fallback to Group CSVs
    for c in [p.parent / "Group 1.csv", p.parent / "Group 2.csv", Path.cwd() / "Group 1.csv", Path.cwd() / "Group 2.csv", Path(__file__).parent.parent / "Group 1.csv", Path(__file__).parent.parent / "Group 2.csv"]:
        if c.exists():
            s2, sk2 = _load_from_csv(c)
            for k, v in s2.items():
                if k not in students:
                    students[k] = v
                else:
                    skipped += 1
            skipped += sk2

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
