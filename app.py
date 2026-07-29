from __future__ import annotations

from datetime import date, datetime, time
from html import escape
from io import BytesIO
from pathlib import Path
import re
from typing import Any

import pandas as pd
import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas


# =========================================================
# PAGE SETTINGS
# =========================================================

st.set_page_config(
    page_title="EntoLabel",
    page_icon="🪲",
    layout="wide",
)

st.title("🪲 EntoLabel")
st.caption(
    "Create collection and determination labels from Excel or CSV files."
)


# =========================================================
# CONSTANTS
# =========================================================

NOT_USED = "— not used —"

ROMAN_MONTHS = {
    1: "I",
    2: "II",
    3: "III",
    4: "IV",
    5: "V",
    6: "VI",
    7: "VII",
    8: "VIII",
    9: "IX",
    10: "X",
    11: "XI",
    12: "XII",
}


# =========================================================
# FONT REGISTRATION
# =========================================================

def register_font(
    internal_name: str,
    possible_paths: list[str],
    fallback_name: str,
) -> str:
    """Register the first available TrueType font."""

    for font_path in possible_paths:
        path = Path(font_path)

        if not path.exists():
            continue

        try:
            pdfmetrics.registerFont(
                TTFont(internal_name, str(path))
            )
            return internal_name
        except Exception:
            continue

    return fallback_name


PDF_FONT_REGULAR = register_font(
    internal_name="EntoLabelRegular",
    possible_paths=[
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman.ttf",
    ],
    fallback_name="Helvetica",
)

PDF_FONT_ITALIC = register_font(
    internal_name="EntoLabelItalic",
    possible_paths=[
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Italic.ttf",
        "/System/Library/Fonts/Supplemental/Arial Italic.ttf",
        "/Library/Fonts/Arial Italic.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica Oblique.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Italic.ttf",
    ],
    fallback_name="Helvetica-Oblique",
)

PDF_FONT_BOLD = register_font(
    internal_name="EntoLabelBold",
    possible_paths=[
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/usr/share/fonts/truetype/liberation2/LiberationSans-Bold.ttf",
        "/System/Library/Fonts/Supplemental/Arial Bold.ttf",
        "/Library/Fonts/Arial Bold.ttf",
        "/System/Library/Fonts/Supplemental/Helvetica Bold.ttf",
        "/System/Library/Fonts/Supplemental/Times New Roman Bold.ttf",
    ],
    fallback_name="Helvetica-Bold",
)


# =========================================================
# GENERAL DATA HELPERS
# =========================================================

def clean_value(value: Any) -> str:
    """Convert an Excel value to readable text."""

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        return value.strftime("%d-%m-%Y")

    if isinstance(value, datetime):
        return value.strftime("%d-%m-%Y")

    if isinstance(value, date):
        return value.strftime("%d-%m-%Y")

    if isinstance(value, time):
        return value.strftime("%H:%M:%S")

    if isinstance(value, float) and value.is_integer():
        return str(int(value))

    return str(value).strip()


def get_value(
    row: pd.Series,
    column_name: str,
) -> str:
    """Get a value from a selected Excel column."""

    if column_name == NOT_USED:
        return ""

    if column_name not in row.index:
        return ""

    return clean_value(row[column_name])


def combine_columns(
    row: pd.Series,
    selected_columns: list[str],
    separator: str = " ",
) -> str:
    """Combine several Excel columns."""

    values = []

    for column_name in selected_columns:
        value = get_value(row, column_name)

        if value:
            values.append(value)

    return separator.join(values)


# =========================================================
# EXCEL ROW RANGE PARSING
# =========================================================

def parse_excel_row_ranges(
    range_text: str,
    minimum_row: int,
    maximum_row: int,
) -> tuple[list[int], list[str]]:
    """Parse entries such as ``10-20, 34, 41-56``.

    The returned row numbers are unique and keep the order in which
    the ranges were entered. Both ends of every range are included.
    """

    cleaned_text = (
        range_text
        .replace("–", "-")
        .replace("—", "-")
        .replace(";", ",")
        .replace("\n", ",")
    )

    parts = [
        part.strip()
        for part in cleaned_text.split(",")
        if part.strip()
    ]

    if not parts:
        return [], ["Enter at least one Excel row or row range."]

    selected_rows: list[int] = []
    seen_rows: set[int] = set()
    errors: list[str] = []

    for part in parts:
        single_match = re.fullmatch(r"\d+", part)
        range_match = re.fullmatch(r"(\d+)\s*-\s*(\d+)", part)

        if single_match:
            start_row = end_row = int(part)
        elif range_match:
            start_row = int(range_match.group(1))
            end_row = int(range_match.group(2))

            if start_row > end_row:
                errors.append(
                    f'Range "{part}" runs backwards. '
                    "Put the smaller row number first."
                )
                continue
        else:
            errors.append(
                f'Could not understand "{part}". '
                "Use formats such as 10-20 or 34."
            )
            continue

        if start_row < minimum_row or end_row > maximum_row:
            errors.append(
                f'"{part}" is outside the available data rows '
                f"{minimum_row}-{maximum_row}."
            )
            continue

        for excel_row in range(start_row, end_row + 1):
            if excel_row not in seen_rows:
                selected_rows.append(excel_row)
                seen_rows.add(excel_row)

    return selected_rows, errors


# =========================================================
# DATE FORMATTING
# =========================================================

def parse_date_value(value: Any) -> pd.Timestamp | None:
    """Try to interpret a value as a date."""

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        return value

    if isinstance(value, datetime):
        return pd.Timestamp(value)

    if isinstance(value, date):
        return pd.Timestamp(value)

    text = str(value).strip()

    if not text:
        return None

    for parsing_options in (
        {"dayfirst": True},
        {"dayfirst": False},
    ):
        try:
            parsed = pd.to_datetime(
                text,
                errors="raise",
                **parsing_options,
            )
            return pd.Timestamp(parsed)
        except Exception:
            continue

    return None


def format_specimen_date(
    value: Any,
    date_format: str,
) -> str:
    """Format collection dates in the selected style."""

    parsed_date = parse_date_value(value)

    if parsed_date is None:
        return clean_value(value)

    day = parsed_date.day
    month = parsed_date.month
    year = parsed_date.year

    if date_format == "15.VII.2026":
        return f"{day}.{ROMAN_MONTHS[month]}.{year}"

    if date_format == "15/07/2026":
        return f"{day:02d}/{month:02d}/{year}"

    if date_format == "15-07-2026":
        return f"{day:02d}-{month:02d}-{year}"

    if date_format == "2026-07-15":
        return f"{year}-{month:02d}-{day:02d}"

    return f"{day}.{ROMAN_MONTHS[month]}.{year}"


# =========================================================
# PERSON NAME FORMATTING
# =========================================================

def shorten_person_name(name: str) -> str:
    """Saniya Sagutdinova -> S. Sagutdinova."""

    cleaned_name = name.strip()

    if not cleaned_name:
        return ""

    parts = cleaned_name.split()

    if len(parts) < 2:
        return cleaned_name

    first_name = parts[0]
    surname = " ".join(parts[1:])

    return f"{first_name[0]}. {surname}"


def split_people(value: str) -> list[str]:
    """Split people separated by commas or semicolons."""

    if not value:
        return []

    normalized = value.replace(";", ",")

    return [
        person.strip()
        for person in normalized.split(",")
        if person.strip()
    ]


def format_people(
    value: str,
    shorten_first_names: bool,
) -> str:
    """Format a collector or identifier list."""

    people = split_people(value)

    if shorten_first_names:
        people = [
            shorten_person_name(person)
            for person in people
        ]

    return ", ".join(people)


def normalize_sex(value: Any) -> str:
    """Convert common sex values to compact biological symbols."""

    text = clean_value(value)

    if not text:
        return ""

    normalized = re.sub(r"[._-]+", " ", text.lower()).strip()

    male_values = {
        "m",
        "male",
        "mannlich",
        "männlich",
        "masculine",
        "♂",
    }
    female_values = {
        "f",
        "female",
        "weiblich",
        "feminine",
        "♀",
    }

    if normalized in male_values:
        return "♂"

    if normalized in female_values:
        return "♀"

    return text


def build_additional_specimen_lines(
    row: pd.Series,
    collecting_method_column: str,
    habitat_column: str,
    host_column: str,
    sex_column: str,
    life_stage_column: str,
    layout: str,
    separator: str,
) -> list[dict[str, str]]:
    """Build optional method, ecology, sex, and life-stage text."""

    collecting_method = get_value(row, collecting_method_column)
    habitat = get_value(row, habitat_column)
    host = get_value(row, host_column)
    sex = normalize_sex(get_value(row, sex_column))
    life_stage = get_value(row, life_stage_column)

    labelled_parts: list[str] = []

    if collecting_method:
        labelled_parts.append(f"method: {collecting_method}")

    if habitat:
        labelled_parts.append(f"hab.: {habitat}")

    if host:
        labelled_parts.append(f"host: {host}")

    specimen_state = " ".join(
        part
        for part in (sex, life_stage)
        if part
    )

    if specimen_state:
        labelled_parts.append(specimen_state)

    if not labelled_parts:
        return []

    if layout == "Separate line for each field":
        return [
            {
                "text": part,
                "style": "regular",
            }
            for part in labelled_parts
        ]

    return [
        {
            "text": separator.join(labelled_parts),
            "style": "regular",
        }
    ]


# =========================================================
# COORDINATE FORMATTING
# =========================================================

def parse_coordinate(value: Any) -> float | None:
    """Try to convert a coordinate to a float."""

    if value is None:
        return None

    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass

    text = str(value).strip().replace(",", ".")

    if not text:
        return None

    try:
        return float(text)
    except ValueError:
        return None


def format_coordinate(
    value: Any,
    coordinate_type: str,
    decimal_places: int,
) -> str:
    """
    Format latitude or longitude.

    Negative latitude becomes S.
    Negative longitude becomes W.
    """

    number = parse_coordinate(value)

    if number is None:
        return clean_value(value)

    absolute_value = abs(number)

    if coordinate_type == "latitude":
        direction = "N" if number >= 0 else "S"
    else:
        direction = "E" if number >= 0 else "W"

    return f"{absolute_value:.{decimal_places}f}°{direction}"


def format_altitude(value: Any) -> str:
    """Add m to altitude unless it already contains a unit."""

    text = clean_value(value)

    if not text:
        return ""

    lowered = text.lower()

    if " m" in lowered or lowered.endswith("m"):
        return text

    return f"{text} m"


def build_coordinates_line(
    row: pd.Series,
    latitude_column: str,
    longitude_column: str,
    altitude_column: str,
    decimal_places: int,
    coordinate_separator: str,
) -> str:
    """Build one optional coordinates/altitude line."""

    latitude = ""

    if latitude_column != NOT_USED:
        latitude = format_coordinate(
            row[latitude_column],
            coordinate_type="latitude",
            decimal_places=decimal_places,
        )

    longitude = ""

    if longitude_column != NOT_USED:
        longitude = format_coordinate(
            row[longitude_column],
            coordinate_type="longitude",
            decimal_places=decimal_places,
        )

    altitude = ""

    if altitude_column != NOT_USED:
        altitude = format_altitude(
            row[altitude_column]
        )

    parts = [
        value
        for value in (latitude, longitude, altitude)
        if value
    ]

    return coordinate_separator.join(parts)


# =========================================================
# LABEL DATA BUILDERS
# =========================================================

def build_collection_lines(
    row: pd.Series,
    specimen_id_column: str,
    locality_columns: list[str],
    locality_separator: str,
    latitude_column: str,
    longitude_column: str,
    altitude_column: str,
    print_coordinates: bool,
    coordinate_decimal_places: int,
    coordinate_separator: str,
    date_column: str,
    date_format: str,
    collector_column: str,
    shorten_collector_names: bool,
    collecting_method_column: str,
    habitat_column: str,
    host_column: str,
    sex_column: str,
    life_stage_column: str,
    additional_details_layout: str,
    additional_details_separator: str,
) -> list[dict[str, str]]:
    """Build the main collection label."""

    specimen_id = get_value(row, specimen_id_column)

    locality = combine_columns(
        row,
        locality_columns,
        separator=locality_separator,
    )

    coordinates_line = ""

    if print_coordinates:
        coordinates_line = build_coordinates_line(
            row=row,
            latitude_column=latitude_column,
            longitude_column=longitude_column,
            altitude_column=altitude_column,
            decimal_places=coordinate_decimal_places,
            coordinate_separator=coordinate_separator,
        )

    raw_date = (
        row[date_column]
        if date_column != NOT_USED
        and date_column in row.index
        else ""
    )

    formatted_date = format_specimen_date(
        raw_date,
        date_format,
    )

    collectors = format_people(
        get_value(row, collector_column),
        shorten_first_names=shorten_collector_names,
    )

    additional_specimen_lines = build_additional_specimen_lines(
        row=row,
        collecting_method_column=collecting_method_column,
        habitat_column=habitat_column,
        host_column=host_column,
        sex_column=sex_column,
        life_stage_column=life_stage_column,
        layout=additional_details_layout,
        separator=additional_details_separator,
    )

    lines: list[dict[str, str]] = []

    if specimen_id:
        lines.append(
            {
                "text": specimen_id,
                "style": "bold",
            }
        )

    if locality:
        lines.append(
            {
                "text": locality,
                "style": "regular",
            }
        )

    if coordinates_line:
        lines.append(
            {
                "text": coordinates_line,
                "style": "regular",
            }
        )

    if formatted_date:
        lines.append(
            {
                "text": formatted_date,
                "style": "regular",
            }
        )

    if collectors:
        lines.append(
            {
                "text": f"leg. {collectors}",
                "style": "regular",
            }
        )

    lines.extend(additional_specimen_lines)

    return lines


def build_determination_lines(
    row: pd.Series,
    specimen_id_column: str,
    print_specimen_id: bool,
    scientific_name_columns: list[str],
    identifier_column: str,
    shorten_identifier_names: bool,
    identification_year_mode: str,
    identification_year_column: str,
    fixed_identification_year: str,
) -> list[dict[str, str]]:
    """Build a separate determination label."""

    specimen_id = get_value(
        row,
        specimen_id_column,
    )

    scientific_name = combine_columns(
        row,
        scientific_name_columns,
        separator=" ",
    )

    identifier = format_people(
        get_value(row, identifier_column),
        shorten_first_names=shorten_identifier_names,
    )

    if identification_year_mode == "Column from Excel":
        identification_year = get_value(
            row,
            identification_year_column,
        )
    elif identification_year_mode == "One year for all labels":
        identification_year = fixed_identification_year.strip()
    else:
        identification_year = ""

    lines: list[dict[str, str]] = []

    if print_specimen_id and specimen_id:
        lines.append(
            {
                "text": specimen_id,
                "style": "bold",
            }
        )

    if scientific_name:
        lines.append(
            {
                "text": scientific_name,
                "style": "italic",
            }
        )

    determination_parts = []

    if identifier:
        determination_parts.append(identifier)

    if identification_year:
        determination_parts.append(identification_year)

    if determination_parts:
        lines.append(
            {
                "text": f"det. {' '.join(determination_parts)}",
                "style": "regular",
            }
        )

    return lines


# =========================================================
# DARWIN CORE EXPORT HELPERS
# =========================================================

def option_index(
    options: list[str],
    preferred_value: str,
) -> int:
    """Return a safe selectbox index for a preferred value."""

    try:
        return options.index(preferred_value)
    except ValueError:
        return 0


def format_dwc_date(value: Any) -> str:
    """Format a date as an ISO 8601 value when possible."""

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    if isinstance(value, pd.Timestamp):
        return value.date().isoformat()

    if isinstance(value, datetime):
        return value.date().isoformat()

    if isinstance(value, date):
        return value.isoformat()

    text = str(value).strip()

    if not text:
        return ""

    # Years, year-month values, full ISO dates, and ISO intervals can
    # already be valid Darwin Core eventDate/dateIdentified values.
    if re.fullmatch(
        r"\d{4}(?:-\d{2}(?:-\d{2})?)?"
        r"(?:[T ][^/]+)?(?:/.*)?",
        text,
    ):
        return text.replace(" ", "T", 1) if " " in text else text

    parsed_date = parse_date_value(value)

    if parsed_date is None:
        return text

    return parsed_date.date().isoformat()


def format_dwc_people(value: Any) -> str:
    """Format a list of people using the Darwin Core list separator."""

    text = clean_value(value)

    if not text:
        return ""

    people = [
        person.strip()
        for person in re.split(r"\s*[;,|]\s*", text)
        if person.strip()
    ]

    return " | ".join(people)


def normalize_dwc_sex(value: Any) -> str:
    """Normalize common sex values to readable Darwin Core values."""

    text = clean_value(value)

    if not text:
        return ""

    normalized = re.sub(r"[._-]+", " ", text.lower()).strip()

    if normalized in {
        "m",
        "male",
        "mannlich",
        "männlich",
        "masculine",
        "♂",
    }:
        return "male"

    if normalized in {
        "f",
        "female",
        "weiblich",
        "feminine",
        "♀",
    }:
        return "female"

    if normalized in {
        "hermaphrodite",
        "hermaphroditic",
        "zwitter",
    }:
        return "hermaphrodite"

    return text


def format_decimal_number(value: Any) -> str:
    """Return a compact decimal number or an empty string."""

    number = parse_coordinate(value)

    if number is None:
        return ""

    return f"{number:.10f}".rstrip("0").rstrip(".")


def extract_numeric_value(value: Any) -> str:
    """Extract the first number from values such as '320 m'."""

    text = clean_value(value).replace(",", ".")

    if not text:
        return ""

    match = re.search(r"[-+]?\d+(?:\.\d+)?", text)

    if match is None:
        return ""

    try:
        number = float(match.group(0))
    except ValueError:
        return ""

    return f"{number:.10f}".rstrip("0").rstrip(".")


def format_associated_host(value: Any) -> str:
    """Represent a host relationship in dwc:associatedTaxa."""

    host = clean_value(value)

    if not host:
        return ""

    safe_host = host.replace('"', "'")
    return f'"host":"{safe_host}"'


def get_raw_value(
    row: pd.Series,
    column_name: str,
) -> Any:
    """Return the original cell value from an optional column."""

    if column_name == NOT_USED:
        return ""

    if column_name not in row.index:
        return ""

    return row[column_name]


def build_dwc_occurrence_id(
    row: pd.Series,
    settings: dict[str, Any],
    catalog_number: str,
) -> str:
    """Build occurrenceID according to the selected strategy."""

    mode = settings["occurrence_id_mode"]

    if mode == "Use an Excel column":
        return get_value(
            row,
            settings["occurrence_id_column"],
        )

    if mode == "Prefix + catalogNumber":
        prefix = settings["occurrence_id_prefix"].strip()

        if prefix and catalog_number:
            return f"{prefix}{catalog_number}"

        return ""

    if mode == "Use catalogNumber directly":
        return catalog_number

    return ""


def build_dwc_identification_date(
    row: pd.Series,
    settings: dict[str, Any],
) -> str:
    """Build dateIdentified from the selected Darwin Core setting."""

    mode = settings["identification_date_mode"]

    if mode == "Use an Excel column":
        return format_dwc_date(
            get_raw_value(
                row,
                settings["identification_date_column"],
            )
        )

    if mode == "Use current determination settings":
        current_mode = settings["current_identification_year_mode"]

        if current_mode == "Column from Excel":
            return format_dwc_date(
                get_raw_value(
                    row,
                    settings["current_identification_year_column"],
                )
            )

        if current_mode == "One year for all labels":
            return format_dwc_date(
                settings["current_fixed_identification_year"]
            )

    return ""


def create_darwin_core_dataframe(
    data: pd.DataFrame,
    settings: dict[str, Any],
) -> pd.DataFrame:
    """Create a flat Simple Darwin Core table from selected records."""

    records: list[dict[str, str]] = []

    for _, row in data.iterrows():
        catalog_number = get_value(
            row,
            settings["catalog_number_column"],
        )

        raw_event_date = get_raw_value(
            row,
            settings["event_date_column"],
        )

        locality = combine_columns(
            row,
            settings["locality_columns"],
            separator=settings["locality_separator"],
        )

        scientific_name = combine_columns(
            row,
            settings["scientific_name_columns"],
            separator=" ",
        )

        raw_elevation = get_raw_value(
            row,
            settings["elevation_column"],
        )

        record = {
            "basisOfRecord": settings["basis_of_record"],
            "datasetName": settings["dataset_name"].strip(),
            "institutionCode": settings["institution_code"].strip(),
            "collectionCode": settings["collection_code"].strip(),
            "license": settings["license"].strip(),
            "rightsHolder": settings["rights_holder"].strip(),
            "catalogNumber": catalog_number,
            "occurrenceID": build_dwc_occurrence_id(
                row=row,
                settings=settings,
                catalog_number=catalog_number,
            ),
            "recordNumber": get_value(
                row,
                settings["record_number_column"],
            ),
            "occurrenceStatus": settings["occurrence_status"],
            "recordedBy": format_dwc_people(
                get_raw_value(
                    row,
                    settings["recorded_by_column"],
                )
            ),
            "individualCount": get_value(
                row,
                settings["individual_count_column"],
            ),
            "sex": normalize_dwc_sex(
                get_raw_value(
                    row,
                    settings["sex_column"],
                )
            ),
            "lifeStage": get_value(
                row,
                settings["life_stage_column"],
            ),
            "preparations": get_value(
                row,
                settings["preparations_column"],
            ),
            "associatedTaxa": format_associated_host(
                get_raw_value(
                    row,
                    settings["host_column"],
                )
            ),
            "eventDate": format_dwc_date(raw_event_date),
            "verbatimEventDate": clean_value(raw_event_date),
            "samplingProtocol": get_value(
                row,
                settings["sampling_protocol_column"],
            ),
            "habitat": get_value(
                row,
                settings["habitat_column"],
            ),
            "country": get_value(
                row,
                settings["country_column"],
            ),
            "countryCode": get_value(
                row,
                settings["country_code_column"],
            ),
            "stateProvince": get_value(
                row,
                settings["state_province_column"],
            ),
            "county": get_value(
                row,
                settings["county_column"],
            ),
            "municipality": get_value(
                row,
                settings["municipality_column"],
            ),
            "locality": locality,
            "decimalLatitude": format_decimal_number(
                get_raw_value(
                    row,
                    settings["latitude_column"],
                )
            ),
            "decimalLongitude": format_decimal_number(
                get_raw_value(
                    row,
                    settings["longitude_column"],
                )
            ),
            "geodeticDatum": settings["geodetic_datum"].strip(),
            "coordinateUncertaintyInMeters": extract_numeric_value(
                get_raw_value(
                    row,
                    settings["coordinate_uncertainty_column"],
                )
            ),
            "minimumElevationInMeters": extract_numeric_value(
                raw_elevation
            ),
            "maximumElevationInMeters": extract_numeric_value(
                raw_elevation
            ),
            "verbatimElevation": clean_value(raw_elevation),
            "scientificName": scientific_name,
            "taxonRank": get_value(
                row,
                settings["taxon_rank_column"],
            ),
            "identificationQualifier": get_value(
                row,
                settings["identification_qualifier_column"],
            ),
            "identifiedBy": format_dwc_people(
                get_raw_value(
                    row,
                    settings["identified_by_column"],
                )
            ),
            "dateIdentified": build_dwc_identification_date(
                row=row,
                settings=settings,
            ),
            "occurrenceRemarks": get_value(
                row,
                settings["occurrence_remarks_column"],
            ),
        }

        records.append(record)

    result = pd.DataFrame(records)

    if not settings["include_empty_columns"]:
        non_empty_columns = [
            column
            for column in result.columns
            if result[column].astype(str).str.strip().ne("").any()
        ]
        result = result[non_empty_columns]

    return result.fillna("")


def darwin_core_validation_messages(
    data: pd.DataFrame,
) -> list[tuple[str, str]]:
    """Return lightweight quality messages for the Darwin Core export."""

    messages: list[tuple[str, str]] = []

    if "catalogNumber" in data.columns:
        catalog_numbers = data["catalogNumber"].astype(str).str.strip()
        non_empty_catalog_numbers = catalog_numbers[catalog_numbers.ne("")]
        missing_count = int(catalog_numbers.eq("").sum())
        duplicate_count = int(non_empty_catalog_numbers.duplicated().sum())

        if missing_count:
            messages.append(
                (
                    "warning",
                    f"{missing_count} record(s) have no catalogNumber.",
                )
            )

        if duplicate_count:
            messages.append(
                (
                    "warning",
                    f"{duplicate_count} duplicate catalogNumber value(s) were found.",
                )
            )

    if "occurrenceID" in data.columns:
        occurrence_ids = data["occurrenceID"].astype(str).str.strip()
        non_empty_occurrence_ids = occurrence_ids[occurrence_ids.ne("")]
        duplicate_occurrence_ids = int(
            non_empty_occurrence_ids.duplicated().sum()
        )

        if occurrence_ids.eq("").all():
            messages.append(
                (
                    "info",
                    "occurrenceID is empty. That is allowed for a draft export, "
                    "but a stable globally unique identifier is recommended "
                    "before publication.",
                )
            )

        if duplicate_occurrence_ids:
            messages.append(
                (
                    "warning",
                    f"{duplicate_occurrence_ids} duplicate occurrenceID value(s) "
                    "were found.",
                )
            )

    for column, label in (
        ("eventDate", "eventDate"),
        ("scientificName", "scientificName"),
        ("locality", "locality"),
    ):
        if column in data.columns:
            missing_count = int(
                data[column].astype(str).str.strip().eq("").sum()
            )

            if missing_count:
                messages.append(
                    (
                        "info",
                        f"{missing_count} record(s) have no {label}.",
                    )
                )

    return messages


# =========================================================
# PDF TEXT HELPERS
# =========================================================

def get_font_name(style: str) -> str:
    """Return the correct PDF font for a text style."""

    if style == "italic":
        return PDF_FONT_ITALIC

    if style == "bold":
        return PDF_FONT_BOLD

    return PDF_FONT_REGULAR


def split_long_word(
    word: str,
    font_name: str,
    font_size: float,
    maximum_width: float,
) -> list[str]:
    """Split a word that is wider than the label."""

    pieces = []
    current_piece = ""

    for character in word:
        test_piece = current_piece + character

        width = pdfmetrics.stringWidth(
            test_piece,
            font_name,
            font_size,
        )

        if width <= maximum_width:
            current_piece = test_piece
        else:
            if current_piece:
                pieces.append(current_piece)

            current_piece = character

    if current_piece:
        pieces.append(current_piece)

    return pieces


def wrap_styled_line(
    text: str,
    style: str,
    font_size: float,
    maximum_width: float,
) -> list[dict[str, str]]:
    """Wrap one styled line according to its printed width."""

    if not text:
        return []

    font_name = get_font_name(style)
    words = text.split()

    wrapped_lines: list[dict[str, str]] = []
    current_line = ""

    for word in words:
        word_width = pdfmetrics.stringWidth(
            word,
            font_name,
            font_size,
        )

        if word_width > maximum_width:
            if current_line:
                wrapped_lines.append(
                    {
                        "text": current_line,
                        "style": style,
                    }
                )
                current_line = ""

            pieces = split_long_word(
                word,
                font_name,
                font_size,
                maximum_width,
            )

            for piece in pieces:
                wrapped_lines.append(
                    {
                        "text": piece,
                        "style": style,
                    }
                )

            continue

        test_line = (
            word
            if not current_line
            else f"{current_line} {word}"
        )

        test_width = pdfmetrics.stringWidth(
            test_line,
            font_name,
            font_size,
        )

        if test_width <= maximum_width:
            current_line = test_line
        else:
            if current_line:
                wrapped_lines.append(
                    {
                        "text": current_line,
                        "style": style,
                    }
                )

            current_line = word

    if current_line:
        wrapped_lines.append(
            {
                "text": current_line,
                "style": style,
            }
        )

    return wrapped_lines


def prepare_styled_lines(
    raw_lines: list[dict[str, str]],
    font_size: float,
    maximum_width: float,
) -> list[dict[str, str]]:
    """Wrap all logical lines of a label."""

    prepared_lines: list[dict[str, str]] = []

    for line in raw_lines:
        prepared_lines.extend(
            wrap_styled_line(
                text=line["text"],
                style=line["style"],
                font_size=font_size,
                maximum_width=maximum_width,
            )
        )

    return prepared_lines


def fit_styled_text_to_label(
    raw_lines: list[dict[str, str]],
    preferred_font_size: float,
    minimum_font_size: float,
    maximum_width: float,
    maximum_height: float,
    line_spacing: float,
) -> tuple[list[dict[str, str]], float, bool]:
    """Reduce the font size until the label fits."""

    font_size = preferred_font_size

    while font_size >= minimum_font_size:
        lines = prepare_styled_lines(
            raw_lines=raw_lines,
            font_size=font_size,
            maximum_width=maximum_width,
        )

        line_height = font_size * line_spacing
        required_height = len(lines) * line_height

        if required_height <= maximum_height:
            return lines, font_size, True

        font_size -= 0.25

    final_lines = prepare_styled_lines(
        raw_lines=raw_lines,
        font_size=minimum_font_size,
        maximum_width=maximum_width,
    )

    return final_lines, minimum_font_size, False


# =========================================================
# LIVE HTML PREVIEW
# =========================================================

def lines_to_html(
    lines: list[dict[str, str]],
) -> str:
    """Convert label lines to safe HTML."""

    html_lines = []

    for line in lines:
        safe_text = escape(line["text"])
        style = line["style"]

        if style == "italic":
            html_lines.append(f"<em>{safe_text}</em>")
        elif style == "bold":
            html_lines.append(f"<strong>{safe_text}</strong>")
        else:
            html_lines.append(safe_text)

    return "<br>".join(html_lines)


def render_live_label(
    title: str,
    lines: list[dict[str, str]],
    width_mm: float,
    height_mm: float,
    font_size_pt: float,
) -> None:
    """Draw an approximate label preview in Streamlit."""

    preview_scale = 4.3

    width_px = max(int(width_mm * preview_scale), 90)
    height_px = max(int(height_mm * preview_scale), 45)

    html_content = lines_to_html(lines)

    st.markdown(f"**{escape(title)}**")

    st.markdown(
        f"""
        <div style="
            width: {width_px}px;
            min-height: {height_px}px;
            border: 1px solid #555;
            background: white;
            color: black;
            padding: 6px;
            overflow: hidden;
            font-family: Arial, Helvetica, sans-serif;
            font-size: {max(font_size_pt * 1.45, 8)}px;
            line-height: 1.08;
            box-sizing: border-box;
        ">
            {html_content}
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================================================
# FILE UPLOAD
# =========================================================

uploaded_file = st.file_uploader(
    "Upload Excel or CSV file",
    type=["xlsx", "xls", "csv"],
)

if uploaded_file is None:
    st.info("Upload an Excel or CSV file to begin.")
    st.stop()


try:
    # Read the sheet without assuming where the column names are.
    # This lets the user choose the Excel/CSV row that contains headers.
    if uploaded_file.name.lower().endswith(".csv"):
        raw_dataframe = pd.read_csv(uploaded_file, header=None)
    else:
        raw_dataframe = pd.read_excel(uploaded_file, header=None)
except Exception as error:
    st.error(f"Could not read the file: {error}")
    st.stop()


if raw_dataframe.empty:
    st.warning("The uploaded file is empty.")
    st.stop()


# =========================================================
# HEADER ROW
# =========================================================

st.subheader("1. Table header")

header_row_number = st.number_input(
    "Excel row containing the column names",
    min_value=1,
    max_value=len(raw_dataframe),
    value=1,
    step=1,
    help=(
        "Usually this is row 1. If your file begins with a title or notes, "
        "choose the row where the actual table headings begin."
    ),
)

header_index = int(header_row_number) - 1
raw_headers = raw_dataframe.iloc[header_index].tolist()

# Make blank and duplicate headings safe for Streamlit/Pandas.
column_names = []
name_counts = {}

for position, value in enumerate(raw_headers, start=1):
    base_name = clean_value(value) or f"Column {position}"
    occurrence = name_counts.get(base_name, 0) + 1
    name_counts[base_name] = occurrence

    if occurrence == 1:
        column_names.append(base_name)
    else:
        column_names.append(f"{base_name} ({occurrence})")

dataframe = raw_dataframe.iloc[header_index + 1:].copy()
dataframe.columns = column_names
dataframe = dataframe.reset_index(drop=True)

if dataframe.empty:
    st.warning("There are no data rows below the selected header row.")
    st.stop()


# =========================================================
# EXCEL PREVIEW
# =========================================================

st.subheader("2. Excel preview")

preview_dataframe = dataframe.head(10).copy()

for column in preview_dataframe.columns:
    preview_dataframe[column] = preview_dataframe[column].map(
        clean_value
    )

st.dataframe(
    preview_dataframe.astype(str),
    width="stretch",
)


# =========================================================
# ROW SELECTION
# =========================================================

st.subheader("3. Select Excel rows")

first_data_excel_row = int(header_row_number) + 1
last_excel_row = len(raw_dataframe)

default_row_ranges = f"{first_data_excel_row}-{last_excel_row}"

row_ranges_text = st.text_input(
    "Excel rows to print",
    value=default_row_ranges,
    key=(
        f"row_ranges_{uploaded_file.name}_"
        f"{int(header_row_number)}_{len(raw_dataframe)}"
    ),
    placeholder="10-20, 34, 41-56, 72",
    help=(
        "Enter individual Excel rows and/or inclusive ranges, separated "
        "by commas. Example: 10-20, 34, 41-56, 72."
    ),
)

selected_excel_rows, row_selection_errors = parse_excel_row_ranges(
    range_text=row_ranges_text,
    minimum_row=first_data_excel_row,
    maximum_row=last_excel_row,
)

if row_selection_errors:
    for error_message in row_selection_errors:
        st.error(error_message)
    st.stop()

selected_positions = [
    excel_row - first_data_excel_row
    for excel_row in selected_excel_rows
]

dataframe = dataframe.iloc[selected_positions].reset_index(drop=True)

if len(selected_excel_rows) <= 20:
    selected_rows_summary = ", ".join(
        str(row_number)
        for row_number in selected_excel_rows
    )
else:
    selected_rows_summary = (
        f"{selected_excel_rows[0]} ... {selected_excel_rows[-1]}"
    )

st.caption(
    f"Selected {len(dataframe)} row(s). "
    f"Excel rows: {selected_rows_summary}. "
    "Overlapping ranges are printed only once."
)


# =========================================================
# COLUMN MAPPING
# =========================================================

all_columns = dataframe.columns.tolist()
optional_columns = [NOT_USED] + all_columns

st.subheader("4. Match Excel columns")

mapping_left, mapping_middle, mapping_right = st.columns(3)


with mapping_left:
    specimen_id_column = st.selectbox(
        "Specimen ID — optional",
        optional_columns,
        help=(
            "The ID will be printed as the first line "
            "of the collection label."
        ),
    )

    locality_columns = st.multiselect(
        "Location — select one or several columns",
        all_columns,
        help="Example: Country + Region + Locality.",
    )

    locality_separator_option = st.selectbox(
        "Location separator",
        options=[
            ", ",
            " · ",
            " ",
            " / ",
        ],
        index=0,
    )

    latitude_column = st.selectbox(
        "Latitude — optional",
        optional_columns,
    )

    longitude_column = st.selectbox(
        "Longitude — optional",
        optional_columns,
    )

    altitude_column = st.selectbox(
        "Altitude — optional",
        optional_columns,
    )


with mapping_middle:
    date_column = st.selectbox(
        "Collection date",
        optional_columns,
    )

    date_format = st.selectbox(
        "Date format",
        options=[
            "15.VII.2026",
            "15/07/2026",
            "15-07-2026",
            "2026-07-15",
        ],
        index=0,
    )

    collector_column = st.selectbox(
        "Collectors",
        optional_columns,
    )

    shorten_collector_names = st.checkbox(
        "Shorten collector first names",
        value=True,
        help=(
            "Saniya Sagutdinova becomes "
            "S. Sagutdinova."
        ),
    )

    print_coordinates = st.checkbox(
        "Print coordinates / altitude",
        value=True,
        disabled=(
            latitude_column == NOT_USED
            and longitude_column == NOT_USED
            and altitude_column == NOT_USED
        ),
    )

    coordinate_decimal_places = st.number_input(
        "Coordinate decimal places",
        min_value=1,
        max_value=8,
        value=4,
        step=1,
        disabled=not print_coordinates,
    )

    coordinate_separator = st.selectbox(
        "Coordinate separator",
        options=[
            ", ",
            " ",
            " · ",
        ],
        index=0,
        disabled=not print_coordinates,
    )


with mapping_right:
    scientific_name_columns = st.multiselect(
        "Scientific name — select several columns",
        all_columns,
        help="Example: Genus + Qualifier + Species.",
    )

    identifier_column = st.selectbox(
        "Identifier / determiner — optional",
        optional_columns,
    )

    shorten_identifier_names = st.checkbox(
        "Shorten identifier first names",
        value=True,
    )


st.markdown("#### Optional collection details")

additional_left, additional_middle, additional_right = st.columns(3)

with additional_left:
    collecting_method_column = st.selectbox(
        "Collecting method — optional",
        optional_columns,
        help="Examples: sweep net, light trap, Malaise trap, hand collected.",
    )

    habitat_column = st.selectbox(
        "Habitat — optional",
        optional_columns,
        help="Example: dry calcareous grassland.",
    )

with additional_middle:
    host_column = st.selectbox(
        "Host — optional",
        optional_columns,
        help="Host plant, animal, fungus, or other associated organism.",
    )

    sex_column = st.selectbox(
        "Sex — optional",
        optional_columns,
        help="Male/female values are converted to ♂/♀ when recognised.",
    )

with additional_right:
    life_stage_column = st.selectbox(
        "Life stage — optional",
        optional_columns,
        help="Examples: adult, larva, nymph, pupa, egg.",
    )

    additional_details_layout = st.radio(
        "Additional details layout",
        options=[
            "Compact — combine fields",
            "Separate line for each field",
        ],
        horizontal=False,
        help=(
            "Compact layout saves space. Separate lines are easier to read "
            "but may require a taller label."
        ),
    )

    additional_details_separator = st.selectbox(
        "Additional details separator",
        options=[
            " · ",
            "; ",
            ", ",
            " / ",
        ],
        index=0,
        disabled=(
            additional_details_layout
            == "Separate line for each field"
        ),
    )


# =========================================================
# IDENTIFICATION SETTINGS
# =========================================================

st.subheader("5. Determination settings")

create_determination_label = st.checkbox(
    "Create a separate determination label",
    value=True,
)

print_id_on_determination_label = st.checkbox(
    "Print specimen ID on determination label",
    value=False,
    disabled=(
        not create_determination_label
        or specimen_id_column == NOT_USED
    ),
    help=(
        "Uses the same specimen ID column selected "
        "for the collection label."
    ),
)

identification_year_mode = st.radio(
    "Identification year",
    options=[
        "Do not print year",
        "Column from Excel",
        "One year for all labels",
    ],
    horizontal=True,
)

identification_year_column = NOT_USED
fixed_identification_year = ""

if identification_year_mode == "Column from Excel":
    identification_year_column = st.selectbox(
        "Identification year column",
        optional_columns,
    )

elif identification_year_mode == "One year for all labels":
    fixed_identification_year = st.text_input(
        "Identification year",
        value=str(datetime.now().year),
    )


# =========================================================
# LABEL SIZE SETTINGS
# =========================================================

st.subheader("6. Label size and typography")

st.markdown("#### Collection label")

collection_settings = st.columns(4)

with collection_settings[0]:
    collection_width_mm = st.number_input(
        "Collection width, mm",
        min_value=10.0,
        max_value=60.0,
        value=20.0,
        step=1.0,
    )

with collection_settings[1]:
    collection_height_mm = st.number_input(
        "Collection height, mm",
        min_value=6.0,
        max_value=40.0,
        value=10.0,
        step=1.0,
    )

with collection_settings[2]:
    collection_font_size = st.number_input(
        "Collection font, pt",
        min_value=3.0,
        max_value=10.0,
        value=5.0,
        step=0.25,
    )

with collection_settings[3]:
    collection_line_spacing = st.number_input(
        "Collection line spacing",
        min_value=0.8,
        max_value=1.5,
        value=1.05,
        step=0.05,
    )


if create_determination_label:
    st.markdown("#### Determination label")

    determination_settings = st.columns(4)

    with determination_settings[0]:
        determination_width_mm = st.number_input(
            "Determination width, mm",
            min_value=10.0,
            max_value=60.0,
            value=20.0,
            step=1.0,
        )

    with determination_settings[1]:
        determination_height_mm = st.number_input(
            "Determination height, mm",
            min_value=5.0,
            max_value=30.0,
            value=7.0,
            step=1.0,
        )

    with determination_settings[2]:
        determination_font_size = st.number_input(
            "Determination font, pt",
            min_value=3.0,
            max_value=10.0,
            value=5.0,
            step=0.25,
        )

    with determination_settings[3]:
        determination_line_spacing = st.number_input(
            "Determination line spacing",
            min_value=0.8,
            max_value=1.5,
            value=1.05,
            step=0.05,
        )

else:
    determination_width_mm = collection_width_mm
    determination_height_mm = collection_height_mm
    determination_font_size = collection_font_size
    determination_line_spacing = collection_line_spacing


draw_borders = st.checkbox(
    "Draw cutting borders",
    value=True,
)


# =========================================================
# BUILD PREVIEW LABELS
# =========================================================

st.subheader("7. Live preview")

preview_excel_row = st.selectbox(
    "Preview Excel row",
    options=selected_excel_rows,
    index=0,
    help="Only rows included in the selection are shown here.",
)

preview_position = selected_excel_rows.index(int(preview_excel_row))
preview_row = dataframe.iloc[preview_position]

collection_preview_lines = build_collection_lines(
    row=preview_row,
    specimen_id_column=specimen_id_column,
    locality_columns=locality_columns,
    locality_separator=locality_separator_option,
    latitude_column=latitude_column,
    longitude_column=longitude_column,
    altitude_column=altitude_column,
    print_coordinates=print_coordinates,
    coordinate_decimal_places=int(coordinate_decimal_places),
    coordinate_separator=coordinate_separator,
    date_column=date_column,
    date_format=date_format,
    collector_column=collector_column,
    shorten_collector_names=shorten_collector_names,
    collecting_method_column=collecting_method_column,
    habitat_column=habitat_column,
    host_column=host_column,
    sex_column=sex_column,
    life_stage_column=life_stage_column,
    additional_details_layout=additional_details_layout,
    additional_details_separator=additional_details_separator,
)

determination_preview_lines = build_determination_lines(
    row=preview_row,
    specimen_id_column=specimen_id_column,
    print_specimen_id=print_id_on_determination_label,
    scientific_name_columns=scientific_name_columns,
    identifier_column=identifier_column,
    shorten_identifier_names=shorten_identifier_names,
    identification_year_mode=identification_year_mode,
    identification_year_column=identification_year_column,
    fixed_identification_year=fixed_identification_year,
)

preview_columns = st.columns(2)

with preview_columns[0]:
    render_live_label(
        title="Collection label",
        lines=collection_preview_lines,
        width_mm=collection_width_mm,
        height_mm=collection_height_mm,
        font_size_pt=collection_font_size,
    )

with preview_columns[1]:
    if create_determination_label:
        render_live_label(
            title="Determination label",
            lines=determination_preview_lines,
            width_mm=determination_width_mm,
            height_mm=determination_height_mm,
            font_size_pt=determination_font_size,
        )
    else:
        st.info(
            "Separate determination labels are disabled."
        )


# =========================================================
# PDF LABEL DRAWING
# =========================================================

def draw_label(
    pdf: canvas.Canvas,
    x: float,
    y: float,
    label_width: float,
    label_height: float,
    raw_lines: list[dict[str, str]],
    preferred_font_size: float,
    line_spacing: float,
    draw_label_border: bool,
) -> bool:
    """Draw one label and return True if all text fits."""

    inner_padding = 0.75 * mm

    available_width = label_width - (2 * inner_padding)
    available_height = label_height - (2 * inner_padding)

    prepared_lines, actual_font_size, fits = (
        fit_styled_text_to_label(
            raw_lines=raw_lines,
            preferred_font_size=preferred_font_size,
            minimum_font_size=3.0,
            maximum_width=available_width,
            maximum_height=available_height,
            line_spacing=line_spacing,
        )
    )

    if draw_label_border:
        pdf.setLineWidth(0.2)
        pdf.rect(
            x,
            y,
            label_width,
            label_height,
        )

    line_height = actual_font_size * line_spacing

    text_y = (
        y
        + label_height
        - inner_padding
        - actual_font_size
    )

    for line in prepared_lines:
        if text_y < y + inner_padding:
            break

        font_name = get_font_name(line["style"])

        pdf.setFont(
            font_name,
            actual_font_size,
        )

        pdf.drawString(
            x + inner_padding,
            text_y,
            line["text"],
        )

        text_y -= line_height

    return fits


# =========================================================
# PDF GENERATION
# =========================================================

def create_pdf(
    data: pd.DataFrame,
) -> tuple[bytes, int]:
    """Create an A4 PDF with collection and determination labels."""

    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)

    page_width, page_height = A4

    page_margin = 7 * mm
    horizontal_gap = 1.5 * mm
    vertical_gap = 1.5 * mm

    x = page_margin
    y_top = page_height - page_margin
    current_row_height = 0.0

    overflow_count = 0

    label_jobs: list[dict[str, Any]] = []

    for _, row in data.iterrows():
        collection_lines = build_collection_lines(
            row=row,
            specimen_id_column=specimen_id_column,
            locality_columns=locality_columns,
            locality_separator=locality_separator_option,
            latitude_column=latitude_column,
            longitude_column=longitude_column,
            altitude_column=altitude_column,
            print_coordinates=print_coordinates,
            coordinate_decimal_places=int(coordinate_decimal_places),
            coordinate_separator=coordinate_separator,
            date_column=date_column,
            date_format=date_format,
            collector_column=collector_column,
            shorten_collector_names=shorten_collector_names,
            collecting_method_column=collecting_method_column,
            habitat_column=habitat_column,
            host_column=host_column,
            sex_column=sex_column,
            life_stage_column=life_stage_column,
            additional_details_layout=additional_details_layout,
            additional_details_separator=additional_details_separator,
        )

        label_jobs.append(
            {
                "lines": collection_lines,
                "width": collection_width_mm * mm,
                "height": collection_height_mm * mm,
                "font_size": collection_font_size,
                "line_spacing": collection_line_spacing,
            }
        )

        if create_determination_label:
            determination_lines = build_determination_lines(
                row=row,
                specimen_id_column=specimen_id_column,
                print_specimen_id=print_id_on_determination_label,
                scientific_name_columns=scientific_name_columns,
                identifier_column=identifier_column,
                shorten_identifier_names=shorten_identifier_names,
                identification_year_mode=identification_year_mode,
                identification_year_column=identification_year_column,
                fixed_identification_year=fixed_identification_year,
            )

            label_jobs.append(
                {
                    "lines": determination_lines,
                    "width": determination_width_mm * mm,
                    "height": determination_height_mm * mm,
                    "font_size": determination_font_size,
                    "line_spacing": determination_line_spacing,
                }
            )

    for job in label_jobs:
        label_width = job["width"]
        label_height = job["height"]

        if x + label_width > page_width - page_margin:
            x = page_margin
            y_top -= current_row_height + vertical_gap
            current_row_height = 0.0

        if y_top - label_height < page_margin:
            pdf.showPage()
            x = page_margin
            y_top = page_height - page_margin
            current_row_height = 0.0

        y = y_top - label_height

        fits = draw_label(
            pdf=pdf,
            x=x,
            y=y,
            label_width=label_width,
            label_height=label_height,
            raw_lines=job["lines"],
            preferred_font_size=job["font_size"],
            line_spacing=job["line_spacing"],
            draw_label_border=draw_borders,
        )

        if not fits:
            overflow_count += 1

        x += label_width + horizontal_gap
        current_row_height = max(
            current_row_height,
            label_height,
        )

    pdf.save()
    buffer.seek(0)

    return buffer.getvalue(), overflow_count


# =========================================================
# EXPORT
# =========================================================

st.subheader("8. Export")

configuration_is_valid = True

if not locality_columns:
    st.warning(
        "Select at least one location column."
    )
    configuration_is_valid = False

if date_column == NOT_USED:
    st.warning(
        "No collection-date column is selected."
    )

if collector_column == NOT_USED:
    st.warning(
        "No collector column is selected."
    )

if (
    create_determination_label
    and not scientific_name_columns
):
    st.warning(
        "Select at least one scientific-name column "
        "for the determination label."
    )
    configuration_is_valid = False


if configuration_is_valid:
    pdf_bytes, overflow_count = create_pdf(dataframe)

    if overflow_count:
        st.warning(
            f"{overflow_count} labels contain too much text "
            "to fit even at 3 pt. Increase their height or "
            "shorten the text."
        )
    else:
        st.success(
            "All labels fit inside the selected dimensions."
        )

    st.download_button(
        label="📄 Create A4 PDF",
        data=pdf_bytes,
        file_name="entomology_labels.pdf",
        mime="application/pdf",
    )


with st.expander(
    "🌿 Darwin Core CSV export — optional",
    expanded=False,
):
    st.caption(
        "Advanced export for museums, collection databases, and "
        "biodiversity-data workflows. Nothing here changes the PDF labels. "
        "The download is a flat Simple Darwin Core CSV, not a full "
        "Darwin Core Archive."
    )

    (
        dwc_dataset_tab,
        dwc_event_tab,
        dwc_specimen_tab,
        dwc_identification_tab,
    ) = st.tabs(
        [
            "Dataset",
            "Event & location",
            "Specimen",
            "Identification",
        ]
    )

    with dwc_dataset_tab:
        dataset_left, dataset_right = st.columns(2)

        with dataset_left:
            dwc_basis_of_record = st.selectbox(
                "basisOfRecord",
                options=[
                    "PreservedSpecimen",
                    "FossilSpecimen",
                    "LivingSpecimen",
                    "MaterialSample",
                    "HumanObservation",
                    "MachineObservation",
                    "MaterialCitation",
                ],
                index=0,
                key="dwc_basis_of_record",
                help=(
                    "For pinned, ethanol-preserved, slide-mounted, or other "
                    "preserved insects, PreservedSpecimen is usually appropriate."
                ),
            )

            dwc_occurrence_status = st.selectbox(
                "occurrenceStatus",
                options=[
                    "detected",
                    "notDetected",
                    "",
                ],
                index=0,
                key="dwc_occurrence_status",
                format_func=lambda value: value or "Leave blank",
            )

            dwc_dataset_name = st.text_input(
                "datasetName — optional",
                key="dwc_dataset_name",
                placeholder="Croatia field course 2026",
            )

            dwc_institution_code = st.text_input(
                "institutionCode — optional",
                key="dwc_institution_code",
                placeholder="Institution acronym",
            )

            dwc_collection_code = st.text_input(
                "collectionCode — optional",
                key="dwc_collection_code",
                placeholder="Diptera",
            )

        with dataset_right:
            dwc_license_choice = st.selectbox(
                "license — optional",
                options=[
                    "Leave blank",
                    "CC0 1.0",
                    "CC BY 4.0",
                    "Custom",
                ],
                key="dwc_license_choice",
                help="Choose only a licence you are authorised to apply.",
            )

            if dwc_license_choice == "CC0 1.0":
                dwc_license = (
                    "https://creativecommons.org/publicdomain/zero/1.0/"
                )
            elif dwc_license_choice == "CC BY 4.0":
                dwc_license = (
                    "https://creativecommons.org/licenses/by/4.0/"
                )
            elif dwc_license_choice == "Custom":
                dwc_license = st.text_input(
                    "Custom licence URL or text",
                    key="dwc_custom_license",
                )
            else:
                dwc_license = ""

            dwc_rights_holder = st.text_input(
                "rightsHolder — optional",
                key="dwc_rights_holder",
            )

            dwc_include_empty_columns = st.checkbox(
                "Keep completely empty Darwin Core columns",
                value=False,
                key="dwc_include_empty_columns",
                help=(
                    "Normally EntoLabel removes fields that are empty for "
                    "every selected record."
                ),
            )

    with dwc_event_tab:
        event_left, event_middle, event_right = st.columns(3)

        with event_left:
            dwc_event_date_column = st.selectbox(
                "eventDate column",
                optional_columns,
                index=option_index(optional_columns, date_column),
                key="dwc_event_date_column",
                help="Dates are converted to ISO format when possible.",
            )

            dwc_recorded_by_column = st.selectbox(
                "recordedBy column",
                optional_columns,
                index=option_index(optional_columns, collector_column),
                key="dwc_recorded_by_column",
            )

            dwc_sampling_protocol_column = st.selectbox(
                "samplingProtocol column",
                optional_columns,
                index=option_index(
                    optional_columns,
                    collecting_method_column,
                ),
                key="dwc_sampling_protocol_column",
            )

            dwc_habitat_column = st.selectbox(
                "habitat column",
                optional_columns,
                index=option_index(optional_columns, habitat_column),
                key="dwc_habitat_column",
            )

        with event_middle:
            dwc_country_column = st.selectbox(
                "country column — optional",
                optional_columns,
                key="dwc_country_column",
            )

            dwc_country_code_column = st.selectbox(
                "countryCode column — optional",
                optional_columns,
                key="dwc_country_code_column",
                help="Prefer a two-letter ISO country code when available.",
            )

            dwc_state_province_column = st.selectbox(
                "stateProvince column — optional",
                optional_columns,
                key="dwc_state_province_column",
            )

            dwc_county_column = st.selectbox(
                "county column — optional",
                optional_columns,
                key="dwc_county_column",
            )

            dwc_municipality_column = st.selectbox(
                "municipality column — optional",
                optional_columns,
                key="dwc_municipality_column",
            )

        with event_right:
            dwc_locality_columns = st.multiselect(
                "locality — select one or several columns",
                all_columns,
                default=[
                    column
                    for column in locality_columns
                    if column in all_columns
                ],
                key="dwc_locality_columns",
            )

            dwc_locality_separator = st.selectbox(
                "Locality separator",
                options=[
                    ", ",
                    " | ",
                    " · ",
                    " / ",
                ],
                index=0,
                key="dwc_locality_separator",
            )

            dwc_latitude_column = st.selectbox(
                "decimalLatitude column",
                optional_columns,
                index=option_index(optional_columns, latitude_column),
                key="dwc_latitude_column",
            )

            dwc_longitude_column = st.selectbox(
                "decimalLongitude column",
                optional_columns,
                index=option_index(optional_columns, longitude_column),
                key="dwc_longitude_column",
            )

            dwc_elevation_column = st.selectbox(
                "Elevation column",
                optional_columns,
                index=option_index(optional_columns, altitude_column),
                key="dwc_elevation_column",
                help=(
                    "The same numeric value is exported as minimum and "
                    "maximum elevation, while the original value is kept "
                    "as verbatimElevation."
                ),
            )

            dwc_geodetic_datum = st.text_input(
                "geodeticDatum — optional",
                key="dwc_geodetic_datum",
                placeholder="WGS84",
                help="Fill this only when the coordinate datum is known.",
            )

            dwc_coordinate_uncertainty_column = st.selectbox(
                "coordinateUncertaintyInMeters column — optional",
                optional_columns,
                key="dwc_coordinate_uncertainty_column",
            )

    with dwc_specimen_tab:
        specimen_left, specimen_right = st.columns(2)

        with specimen_left:
            dwc_catalog_number_column = st.selectbox(
                "catalogNumber column",
                optional_columns,
                index=option_index(optional_columns, specimen_id_column),
                key="dwc_catalog_number_column",
                help=(
                    "This normally uses the same specimen ID selected for "
                    "the label."
                ),
            )

            dwc_occurrence_id_mode = st.selectbox(
                "occurrenceID",
                options=[
                    "Leave blank",
                    "Use an Excel column",
                    "Prefix + catalogNumber",
                    "Use catalogNumber directly",
                ],
                key="dwc_occurrence_id_mode",
                help=(
                    "A stable globally unique occurrenceID is recommended "
                    "for publication. A local catalogNumber alone may not be "
                    "globally unique."
                ),
            )

            dwc_occurrence_id_column = NOT_USED
            dwc_occurrence_id_prefix = ""

            if dwc_occurrence_id_mode == "Use an Excel column":
                dwc_occurrence_id_column = st.selectbox(
                    "occurrenceID column",
                    optional_columns,
                    key="dwc_occurrence_id_column",
                )

            elif dwc_occurrence_id_mode == "Prefix + catalogNumber":
                dwc_occurrence_id_prefix = st.text_input(
                    "Stable prefix",
                    key="dwc_occurrence_id_prefix",
                    placeholder="https://example.org/specimens/",
                )

            dwc_record_number_column = st.selectbox(
                "recordNumber column — optional",
                optional_columns,
                key="dwc_record_number_column",
                help="A collector's field number, when different from catalogNumber.",
            )

            dwc_individual_count_column = st.selectbox(
                "individualCount column — optional",
                optional_columns,
                key="dwc_individual_count_column",
            )

        with specimen_right:
            dwc_sex_column = st.selectbox(
                "sex column",
                optional_columns,
                index=option_index(optional_columns, sex_column),
                key="dwc_sex_column",
            )

            dwc_life_stage_column = st.selectbox(
                "lifeStage column",
                optional_columns,
                index=option_index(optional_columns, life_stage_column),
                key="dwc_life_stage_column",
            )

            dwc_host_column = st.selectbox(
                "Host column → associatedTaxa",
                optional_columns,
                index=option_index(optional_columns, host_column),
                key="dwc_host_column",
                help=(
                    "A host such as Quercus robur is exported as "
                    '"host":"Quercus robur".'
                ),
            )

            dwc_preparations_column = st.selectbox(
                "preparations column — optional",
                optional_columns,
                key="dwc_preparations_column",
                help="Examples: pinned, ethanol 96%, slide-mounted.",
            )

            dwc_occurrence_remarks_column = st.selectbox(
                "occurrenceRemarks column — optional",
                optional_columns,
                key="dwc_occurrence_remarks_column",
            )

    with dwc_identification_tab:
        identification_left, identification_right = st.columns(2)

        with identification_left:
            dwc_scientific_name_columns = st.multiselect(
                "scientificName — select one or several columns",
                all_columns,
                default=[
                    column
                    for column in scientific_name_columns
                    if column in all_columns
                ],
                key="dwc_scientific_name_columns",
            )

            dwc_identified_by_column = st.selectbox(
                "identifiedBy column",
                optional_columns,
                index=option_index(optional_columns, identifier_column),
                key="dwc_identified_by_column",
            )

            dwc_identification_date_mode = st.selectbox(
                "dateIdentified",
                options=[
                    "Use current determination settings",
                    "Use an Excel column",
                    "Leave blank",
                ],
                key="dwc_identification_date_mode",
            )

            dwc_identification_date_column = NOT_USED

            if dwc_identification_date_mode == "Use an Excel column":
                dwc_identification_date_column = st.selectbox(
                    "dateIdentified column",
                    optional_columns,
                    key="dwc_identification_date_column",
                )

        with identification_right:
            dwc_identification_qualifier_column = st.selectbox(
                "identificationQualifier column — optional",
                optional_columns,
                key="dwc_identification_qualifier_column",
                help="Examples: cf., aff., ?, sensu lato.",
            )

            dwc_taxon_rank_column = st.selectbox(
                "taxonRank column — optional",
                optional_columns,
                key="dwc_taxon_rank_column",
                help="Examples: species, genus, family.",
            )

    darwin_core_settings = {
        "basis_of_record": dwc_basis_of_record,
        "occurrence_status": dwc_occurrence_status,
        "dataset_name": dwc_dataset_name,
        "institution_code": dwc_institution_code,
        "collection_code": dwc_collection_code,
        "license": dwc_license,
        "rights_holder": dwc_rights_holder,
        "include_empty_columns": dwc_include_empty_columns,
        "event_date_column": dwc_event_date_column,
        "recorded_by_column": dwc_recorded_by_column,
        "sampling_protocol_column": dwc_sampling_protocol_column,
        "habitat_column": dwc_habitat_column,
        "country_column": dwc_country_column,
        "country_code_column": dwc_country_code_column,
        "state_province_column": dwc_state_province_column,
        "county_column": dwc_county_column,
        "municipality_column": dwc_municipality_column,
        "locality_columns": dwc_locality_columns,
        "locality_separator": dwc_locality_separator,
        "latitude_column": dwc_latitude_column,
        "longitude_column": dwc_longitude_column,
        "elevation_column": dwc_elevation_column,
        "geodetic_datum": dwc_geodetic_datum,
        "coordinate_uncertainty_column": (
            dwc_coordinate_uncertainty_column
        ),
        "catalog_number_column": dwc_catalog_number_column,
        "occurrence_id_mode": dwc_occurrence_id_mode,
        "occurrence_id_column": dwc_occurrence_id_column,
        "occurrence_id_prefix": dwc_occurrence_id_prefix,
        "record_number_column": dwc_record_number_column,
        "individual_count_column": dwc_individual_count_column,
        "sex_column": dwc_sex_column,
        "life_stage_column": dwc_life_stage_column,
        "host_column": dwc_host_column,
        "preparations_column": dwc_preparations_column,
        "occurrence_remarks_column": dwc_occurrence_remarks_column,
        "scientific_name_columns": dwc_scientific_name_columns,
        "identified_by_column": dwc_identified_by_column,
        "identification_date_mode": dwc_identification_date_mode,
        "identification_date_column": dwc_identification_date_column,
        "identification_qualifier_column": (
            dwc_identification_qualifier_column
        ),
        "taxon_rank_column": dwc_taxon_rank_column,
        "current_identification_year_mode": identification_year_mode,
        "current_identification_year_column": identification_year_column,
        "current_fixed_identification_year": fixed_identification_year,
    }

    darwin_core_dataframe = create_darwin_core_dataframe(
        data=dataframe,
        settings=darwin_core_settings,
    )

    st.markdown("#### Export check")
    st.caption(
        f"{len(darwin_core_dataframe)} record(s) · "
        f"{len(darwin_core_dataframe.columns)} Darwin Core field(s)"
    )

    for message_type, message_text in darwin_core_validation_messages(
        darwin_core_dataframe
    ):
        if message_type == "warning":
            st.warning(message_text)
        else:
            st.info(message_text)

    show_dwc_preview = st.checkbox(
        "Show Darwin Core preview",
        value=False,
        key="show_dwc_preview",
    )

    if show_dwc_preview:
        st.dataframe(
            darwin_core_dataframe.head(20).astype(str),
            width="stretch",
        )

    darwin_core_csv = darwin_core_dataframe.to_csv(
        index=False,
        lineterminator="\n",
    ).encode("utf-8-sig")

    st.download_button(
        label="Download Darwin Core CSV",
        data=darwin_core_csv,
        file_name="entolabel_darwin_core.csv",
        mime="text/csv",
        key="download_darwin_core_csv",
    )
