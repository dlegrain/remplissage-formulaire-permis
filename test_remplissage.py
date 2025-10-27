#!/usr/bin/env python3
"""Remplissage contextuel du formulaire de demande de permis.

Le script détecte les sections (Cadre 1, 2, 3, ...) du document Word et
remplace uniquement les placeholders correspondant aux champs connus en
fonction du contexte (personne physique, personne morale, localisation,
objet de la demande, ...).

Modifiez les dictionnaires `FORM_DATA` et `FIELD_RULES` pour adapter les
valeurs ou ajouter de nouveaux champs.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Sequence, Set, Tuple

from docx import Document
from docx.document import Document as _Document
from docx.table import Table, _Cell
from docx.text.paragraph import Paragraph

# Regex pour détecter les zones à remplir (séquences de points, tirets, ellipses, etc.).
PLACEHOLDER_PATTERN = re.compile(r"[._·•\u2022\u2024\u2025\u2026\u00b7\-–—]{3,}")

# Marqueurs de sections et sous-sections pour suivre le contexte courant.
SECTION_MARKERS = {
    "demandeur": re.compile(r"Cadre\s*1\s*[-–]\s*Demandeur", re.IGNORECASE),
    "localisation": re.compile(r"Cadre\s*2", re.IGNORECASE),
    "objet": re.compile(r"Cadre\s*3", re.IGNORECASE),
}

SUBSECTION_MARKERS = {
    "demandeur": {
        "personne_physique": re.compile(r"Personne\s+physique", re.IGNORECASE),
        "personne_morale": re.compile(r"Personne\s+morale", re.IGNORECASE),
    },
}


@dataclass
class FormContext:
    """Représente le contexte de remplissage courant."""

    section: Optional[str] = None
    subsection: Optional[str] = None


@dataclass(frozen=True)
class FieldRule:
    """Règle de remplacement contextuelle."""

    id: str
    pattern: re.Pattern[str]
    keys: Sequence[str]
    action: str = "replace"  # "replace" (par défaut) ou "append"


FORM_DATA: Dict[Tuple[str, Optional[str]], Dict[str, str]] = {
    ("demandeur", "personne_physique"): {
        "nom": "MARTIN",
        "prenom": "Jean",
        "numero_national": "85.10.25-123.45",
        "rue": "Rue du Moulin",
        "numero": "12",
        "boite": "",
        "code_postal": "4420",
        "commune": "Saint-Nicolas",
        "pays": "Belgique",
        "telephone": "+32 498 12 34 56",
        "courriel": "jean.martin@example.com",
    },
    ("demandeur", "personne_morale"): {
        "denomination": "Martin Construction SRL",
        "forme_juridique": "SRL",
        "numero_bce": "0753.123.456",
        "rue": "Chaussée Verte",
        "numero": "101",
        "boite": "B",
        "code_postal": "4000",
        "commune": "Liège",
        "pays": "Belgique",
        "telephone": "+32 4 234 56 78",
        "courriel": "contact@martinconstruction.be",
    },
    ("localisation", None): {
        "rue": "Rue des Tilleuls",
        "numero": "58",
        "boite": "",
        "commune_affichage": "5000 Namur",
    },
    ("objet", None): {
        "description_generale": (
            "Le projet prévoit la rénovation complète de l'annexe arrière "
            "afin d'y aménager un atelier de menuiserie légère. Les travaux "
            "comprennent la démolition de la toiture existante, la pose d'une "
            "toiture plate végétalisée, le remplacement des menuiseries par "
            "des châssis en aluminium thermolaqué ainsi que l'isolation de "
            "l'enveloppe par l'extérieur. Un nouvel accès PMR et une zone de "
            "stationnement perméable sont également prévus."
        ),
    },
}

FIELD_RULES: Dict[Tuple[str, Optional[str]], List[FieldRule]] = {
    ("demandeur", "personne_physique"): [
        FieldRule(
            id="demandeur_nom_prenom",
            pattern=re.compile(r"^Nom\s*:", re.IGNORECASE),
            keys=("nom", "prenom"),
        ),
        FieldRule(
            id="demandeur_numero_national",
            pattern=re.compile(r"^N°\s*national", re.IGNORECASE),
            keys=("numero_national",),
        ),
        FieldRule(
            id="demandeur_adresse_rue",
            pattern=re.compile(r"^Rue", re.IGNORECASE),
            keys=("rue", "numero", "boite"),
        ),
        FieldRule(
            id="demandeur_adresse_commune",
            pattern=re.compile(r"^Code\s*postal", re.IGNORECASE),
            keys=("code_postal", "commune", "pays"),
        ),
        FieldRule(
            id="demandeur_telephone",
            pattern=re.compile(r"^Téléphone", re.IGNORECASE),
            keys=("telephone",),
        ),
        FieldRule(
            id="demandeur_courriel",
            pattern=re.compile(r"^Courriel", re.IGNORECASE),
            keys=("courriel",),
        ),
    ],
    ("demandeur", "personne_morale"): [
        FieldRule(
            id="morale_denomination",
            pattern=re.compile(r"^Dénomination", re.IGNORECASE),
            keys=("denomination",),
        ),
        FieldRule(
            id="morale_forme",
            pattern=re.compile(r"^Forme\s+juridique", re.IGNORECASE),
            keys=("forme_juridique",),
        ),
        FieldRule(
            id="morale_bce",
            pattern=re.compile(r"^Numéro\s+BCE", re.IGNORECASE),
            keys=("numero_bce",),
        ),
        FieldRule(
            id="morale_adresse_rue",
            pattern=re.compile(r"^Rue", re.IGNORECASE),
            keys=("rue", "numero", "boite"),
        ),
        FieldRule(
            id="morale_adresse_commune",
            pattern=re.compile(r"^Code\s*postal", re.IGNORECASE),
            keys=("code_postal", "commune", "pays"),
        ),
        FieldRule(
            id="morale_telephone",
            pattern=re.compile(r"^Téléphone", re.IGNORECASE),
            keys=("telephone",),
        ),
        FieldRule(
            id="morale_courriel",
            pattern=re.compile(r"^Courriel", re.IGNORECASE),
            keys=("courriel",),
        ),
    ],
    ("localisation", None): [
        FieldRule(
            id="localisation_rue",
            pattern=re.compile(r"^Rue", re.IGNORECASE),
            keys=("rue", "numero", "boite"),
        ),
        FieldRule(
            id="localisation_commune",
            pattern=re.compile(r"^Commune", re.IGNORECASE),
            keys=("commune_affichage",),
        ),
    ],
    ("objet", None): [
        FieldRule(
            id="objet_description",
            pattern=re.compile(r"Décrivez\s+l[’']entièreté\s+du\s+projet", re.IGNORECASE),
            keys=("description_generale",),
            action="append",
        ),
    ],
}


def normalize_text(value: str) -> str:
    """Remplace les espaces insécables et supprime les bords."""

    return value.replace("\xa0", " ").strip()


def update_context_from_text(text: str, context: FormContext) -> None:
    """Met à jour le contexte courant en fonction du texte analysé."""

    for section, marker in SECTION_MARKERS.items():
        if marker.search(text):
            context.section = section
            context.subsection = None
    if context.section in SUBSECTION_MARKERS:
        for subsection, marker in SUBSECTION_MARKERS[context.section].items():
            if marker.search(text):
                context.subsection = subsection


def collect_values(data: Dict[str, str], keys: Sequence[str]) -> Optional[List[str]]:
    """Récupère les valeurs ordonnées associées aux clés données."""

    values: List[str] = []
    for key in keys:
        if key not in data:
            return None
        values.append(str(data.get(key, "")))
    return values


def replace_placeholders(text: str, values: Sequence[str]) -> str:
    """Remplace séquentiellement les placeholders trouvés dans le texte."""

    if not values:
        return text

    consumed = 0
    values = list(values)
    iterator = iter(values)

    def _substitute(match: re.Match[str]) -> str:
        nonlocal consumed
        try:
            value = next(iterator)
        except StopIteration:
            return match.group(0)
        consumed += 1
        return value

    new_text, _ = PLACEHOLDER_PATTERN.subn(_substitute, text, count=len(values))

    if consumed < len(values):
        remaining = " ".join(values[consumed:])
        if remaining:
            new_text = f"{new_text.strip()} {remaining}".strip()
    return new_text


def log_replacement(rule_id: str, values: Sequence[str]) -> None:
    """Affiche en console une trace du remplacement effectué."""

    joined = ", ".join(v for v in values if v)
    print(f"  - {rule_id}: {joined}")


def process_paragraph(paragraph: Paragraph, context: FormContext, filled_rules: Set[str]) -> None:
    """Analyse un paragraphe, met à jour le contexte et applique les règles."""

    raw_text = paragraph.text
    text = normalize_text(raw_text)
    if not text:
        return

    update_context_from_text(text, context)

    data = FORM_DATA.get((context.section, context.subsection))
    if not data:
        return

    rules = FIELD_RULES.get((context.section, context.subsection), [])
    for rule in rules:
        if rule.id in filled_rules:
            continue
        if not rule.pattern.search(text):
            continue

        values = collect_values(data, rule.keys)
        if values is None:
            continue

        if rule.action == "append":
            appended = values[0]
            if appended:
                paragraph.add_run(f"\n{appended}")
                filled_rules.add(rule.id)
                log_replacement(rule.id, (appended,))
            continue

        new_text = replace_placeholders(raw_text, values)
        if new_text != raw_text:
            paragraph.text = new_text
            filled_rules.add(rule.id)
            log_replacement(rule.id, values)


def process_table(table: Table, context: FormContext, filled_rules: Set[str]) -> None:
    """Parcourt récursivement les tableaux du document."""

    for row in table.rows:
        for cell in row.cells:
            for block in iter_block_items(cell):
                if isinstance(block, Paragraph):
                    process_paragraph(block, context, filled_rules)
                elif isinstance(block, Table):
                    process_table(block, context, filled_rules)


def iter_block_items(parent: _Document | _Cell) -> Iterator[Paragraph | Table]:
    """Itère sur les paragraphes et tableaux dans l'ordre du document."""

    if isinstance(parent, _Document):
        parent_element = parent.element.body
    else:
        parent_element = parent._tc

    for child in parent_element.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def remplir_formulaire_intelligent() -> Tuple[Path, Set[str]]:
    """Remplit le formulaire DOCX avec les données d'exemple."""

    base_dir = Path(__file__).resolve().parent
    template_path = base_dir / "annexe-6-demande-de-permis-sans-architecte.docx"
    if not template_path.exists():
        raise FileNotFoundError(f"Impossible de trouver le modèle : {template_path}")

    output_path = base_dir / "formulaire_rempli.docx"

    doc = Document(template_path)
    context = FormContext()
    filled_rules: Set[str] = set()

    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            process_paragraph(block, context, filled_rules)
        elif isinstance(block, Table):
            process_table(block, context, filled_rules)

    doc.save(output_path)
    return output_path, filled_rules


def main() -> None:
    print("🚀 Remplissage intelligent du formulaire...")
    output_path, filled_rules = remplir_formulaire_intelligent()
    expected = {rule.id for rules in FIELD_RULES.values() for rule in rules}
    print("Champs renseignés :")
    for rule_id in sorted(expected):
        status = "OK" if rule_id in filled_rules else "—"
        print(f"  {status} {rule_id}")
    missing = sorted(expected - filled_rules)
    if missing:
        print("Champs restants à compléter :", ", ".join(missing))
    print(f"✅ Fichier généré : {output_path}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:  # pylint: disable=broad-except
        print(f"❌ Erreur : {exc}")
