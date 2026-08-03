from __future__ import annotations

from datetime import date, datetime, time
from html import escape
from io import BytesIO
from pathlib import Path
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

def append_inline_value(
    base_text: str,
    extra_text: str,
    separator: str = " · ",
) -> str:
    """Join two label values without wasting a separate line."""

    if base_text and extra_text:
        return f"{base_text}{separator}{extra_text}"

    return base_text or extra_text


def build_collection_lines(
    row: pd.Series,
    specimen_id_column: str,
    specimen_id_placement: str,
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
) -> list[dict[str, str]]:
    """Build a compact collection label with configurable ID placement."""

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

    collector_line = f"leg. {collectors}" if collectors else ""
    separate_first_id = ""
    separate_last_id = ""

    if specimen_id:
        if specimen_id_placement == "Compact — before first content":
            # Keep the catalogue number at the very beginning of the label
            # without spending a separate line. Prefer the locality line,
            # then fall back to the next available collection-data line.
            if locality:
                locality = append_inline_value(
                    specimen_id,
                    locality,
                )
            elif coordinates_line:
                coordinates_line = append_inline_value(
                    specimen_id,
                    coordinates_line,
                )
            elif formatted_date:
                formatted_date = append_inline_value(
                    specimen_id,
                    formatted_date,
                )
            elif collector_line:
                collector_line = append_inline_value(
                    specimen_id,
                    collector_line,
                )
            else:
                separate_first_id = specimen_id
        elif specimen_id_placement == "Separate first line":
            separate_first_id = specimen_id
        elif specimen_id_placement == "Separate last line":
            separate_last_id = specimen_id
        elif specimen_id_placement == "Compact — after collector":
            if collector_line:
                collector_line = append_inline_value(
                    collector_line,
                    specimen_id,
                )
            elif formatted_date:
                formatted_date = append_inline_value(
                    formatted_date,
                    specimen_id,
                )
            elif coordinates_line:
                coordinates_line = append_inline_value(
                    coordinates_line,
                    specimen_id,
                )
            elif locality:
                locality = append_inline_value(locality, specimen_id)
            else:
                separate_last_id = specimen_id
        elif specimen_id_placement == "Compact — after date":
            if formatted_date:
                formatted_date = append_inline_value(
                    formatted_date,
                    specimen_id,
                )
            elif collector_line:
                collector_line = append_inline_value(
                    collector_line,
                    specimen_id,
                )
            elif coordinates_line:
                coordinates_line = append_inline_value(
                    coordinates_line,
                    specimen_id,
                )
            elif locality:
                locality = append_inline_value(locality, specimen_id)
            else:
                separate_last_id = specimen_id
        # "Do not print" intentionally leaves the ID out.

    lines: list[dict[str, str]] = []

    if separate_first_id:
        lines.append(
            {
                "text": separate_first_id,
                "style": "bold",
            }
        )

    for text in (
        locality,
        coordinates_line,
        formatted_date,
        collector_line,
    ):
        if text:
            lines.append(
                {
                    "text": text,
                    "style": "regular",
                }
            )

    if separate_last_id:
        lines.append(
            {
                "text": separate_last_id,
                "style": "bold",
            }
        )

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
    automatically_enlarge: bool = False,
    maximum_font_size: float | None = None,
) -> tuple[list[dict[str, str]], float, bool]:
    """Choose the largest allowed font size that still fits the label."""

    start_font_size = preferred_font_size

    if automatically_enlarge and maximum_font_size is not None:
        start_font_size = max(
            preferred_font_size,
            maximum_font_size,
        )

    font_size = start_font_size

    while font_size >= minimum_font_size:
        lines = prepare_styled_lines(
            raw_lines=raw_lines,
            font_size=font_size,
            maximum_width=maximum_width,
        )

        line_height = font_size * line_spacing
        required_height = (
            0.0
            if not lines
            else font_size + (len(lines) - 1) * line_height
        )

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
    line_spacing: float,
    automatically_enlarge: bool,
    maximum_font_size_pt: float,
    vertical_alignment: str,
) -> None:
    """Draw a preview using the same wrapping and fitting logic as the PDF."""

    preview_scale = 4.3
    preview_padding_pt = 0.45 * mm

    width_px = max(int(width_mm * preview_scale), 90)
    height_px = max(int(height_mm * preview_scale), 45)

    available_width = max((width_mm * mm) - 2 * preview_padding_pt, 1)
    available_height = max((height_mm * mm) - 2 * preview_padding_pt, 1)

    prepared_lines, actual_font_size, _ = fit_styled_text_to_label(
        raw_lines=lines,
        preferred_font_size=font_size_pt,
        minimum_font_size=3.0,
        maximum_width=available_width,
        maximum_height=available_height,
        line_spacing=line_spacing,
        automatically_enlarge=automatically_enlarge,
        maximum_font_size=maximum_font_size_pt,
    )

    html_content = lines_to_html(prepared_lines)
    justify_content = (
        "center"
        if vertical_alignment == "Balanced"
        else "flex-start"
    )

    st.markdown(f"**{escape(title)}**")

    st.markdown(
        f"""
        <div style="
            width: {width_px}px;
            min-height: {height_px}px;
            border: 1px solid #555;
            background: white;
            color: black;
            padding: 3px;
            overflow: hidden;
            font-family: Arial, Helvetica, sans-serif;
            font-size: {max(actual_font_size * 1.45, 8)}px;
            line-height: {line_spacing};
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
            justify-content: {justify_content};
        ">
            <div>{html_content}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.caption(f"Preview font: {actual_font_size:.2f} pt")


# =========================================================
# ONBOARDING AND FILE UPLOAD
# =========================================================

def normalize_heading(value: Any) -> str:
    """Normalise a column heading for safe automatic matching."""

    text = clean_value(value).casefold()
    return "".join(character for character in text if character.isalnum())


def suggest_column(
    columns: list[str],
    aliases: list[str],
) -> str:
    """Return the first column matching one of the expected aliases."""

    normalized_columns = {
        normalize_heading(column): column
        for column in columns
    }

    for alias in aliases:
        exact_match = normalized_columns.get(normalize_heading(alias))
        if exact_match:
            return exact_match

    # A conservative partial match helps with headings such as
    # "Elevation (m)" without confusing unrelated short headings.
    for alias in aliases:
        normalized_alias = normalize_heading(alias)
        if len(normalized_alias) < 4:
            continue

        for normalized_column, original_column in normalized_columns.items():
            if (
                normalized_alias in normalized_column
                or normalized_column in normalized_alias
            ):
                return original_column

    return NOT_USED


def detect_header_row(raw_data: pd.DataFrame) -> int:
    """Suggest a likely 1-based header row from the first 20 rows."""

    known_headings = {
        normalize_heading(value)
        for value in [
            "Specimen ID",
            "Catalogue number",
            "Country",
            "Region",
            "Locality",
            "Latitude",
            "Longitude",
            "Elevation",
            "Altitude",
            "Collection date",
            "Date",
            "Collector",
            "Scientific name",
            "Taxon",
            "Genus",
            "Species",
            "Determined by",
            "Identifier",
        ]
    }

    best_index = 0
    best_score = -1.0

    for index in range(min(len(raw_data), 20)):
        values = [
            clean_value(value)
            for value in raw_data.iloc[index].tolist()
        ]
        nonempty_values = [value for value in values if value]

        if not nonempty_values:
            continue

        normalized_values = {
            normalize_heading(value)
            for value in nonempty_values
        }
        known_matches = len(normalized_values & known_headings)
        text_values = sum(
            any(character.isalpha() for character in value)
            for value in nonempty_values
        )
        uniqueness = len(set(nonempty_values)) / len(nonempty_values)

        score = known_matches * 10 + text_values + uniqueness

        if score > best_score:
            best_score = score
            best_index = index

    return best_index + 1


st.markdown("### Start with your own collection data")
st.caption(
    "EntoLabel now suggests common column mappings automatically. "
    "You can review every suggestion before creating the PDF."
)

onboarding_columns = st.columns(3)
with onboarding_columns[0]:
    st.markdown("**1. Upload**  \nExcel or CSV; your headings do not need to use exactly the same names.")
with onboarding_columns[1]:
    st.markdown("**2. Review mapping**  \nEntoLabel matches locality, date, collector, taxon and ID fields.")
with onboarding_columns[2]:
    st.markdown("**3. Preview and export**  \nCheck the real label layout before downloading the A4 PDF.")

template_columns = [
    "Specimen ID",
    "Taxon",
    "Country",
    "Region",
    "Locality",
    "Latitude",
    "Longitude",
    "Elevation (m)",
    "Collection date",
    "Collector",
    "Determined by",
]

empty_template = pd.DataFrame(columns=template_columns)
example_template = pd.DataFrame(
    [
        {
            "Specimen ID": "AUR-0001",
            "Taxon": "Carabus auratus",
            "Country": "Aurelia",
            "Region": "Mossreach",
            "Locality": "Moonfern Hollow",
            "Latitude": 47.2015,
            "Longitude": 8.5142,
            "Elevation (m)": 612,
            "Collection date": "2026-04-03",
            "Collector": "Mira Solberg",
            "Determined by": "Elian Voss",
        },
        {
            "Specimen ID": "AUR-0002",
            "Taxon": "Lucilia sericata",
            "Country": "Aurelia",
            "Region": "Silverfen",
            "Locality": "Glasswing Meadow",
            "Latitude": 47.1842,
            "Longitude": 8.5371,
            "Elevation (m)": 548,
            "Collection date": "2026-05-19",
            "Collector": "Nora Vale",
            "Determined by": "Elian Voss",
        },
    ]
)

template_downloads = st.columns(2)
with template_downloads[0]:
    st.download_button(
        "Download empty CSV template",
        data=empty_template.to_csv(index=False).encode("utf-8-sig"),
        file_name="EntoLabel_empty_template.csv",
        mime="text/csv",
        use_container_width=True,
    )
with template_downloads[1]:
    st.download_button(
        "Download fictional example",
        data=example_template.to_csv(index=False).encode("utf-8-sig"),
        file_name="EntoLabel_fictional_example.csv",
        mime="text/csv",
        use_container_width=True,
    )

uploaded_file = st.file_uploader(
    "Upload your Excel or CSV file",
    type=["xlsx", "xls", "csv"],
    help=(
        "The first row normally contains column names. If your file begins "
        "with notes or a title, EntoLabel will suggest another header row."
    ),
)

if uploaded_file is None:
    st.info(
        "Upload your own file, or download the template above to see the "
        "recommended structure."
    )
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

suggested_header_row = detect_header_row(raw_dataframe)

header_row_number = st.number_input(
    "Excel row containing the column names",
    min_value=1,
    max_value=len(raw_dataframe),
    value=int(suggested_header_row),
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

row_selection_columns = st.columns(2)

with row_selection_columns[0]:
    start_excel_row = st.number_input(
        "From Excel row",
        min_value=first_data_excel_row,
        max_value=last_excel_row,
        value=first_data_excel_row,
        step=1,
    )

with row_selection_columns[1]:
    end_excel_row = st.number_input(
        "To Excel row",
        min_value=int(start_excel_row),
        max_value=last_excel_row,
        value=last_excel_row,
        step=1,
    )

start_position = int(start_excel_row) - first_data_excel_row
end_position = int(end_excel_row) - first_data_excel_row

dataframe = dataframe.iloc[start_position:end_position + 1].reset_index(
    drop=True
)

st.caption(
    f"Selected {len(dataframe)} row(s): "
    f"Excel rows {int(start_excel_row)}–{int(end_excel_row)}."
)


# =========================================================
# COLUMN MAPPING
# =========================================================

all_columns = dataframe.columns.tolist()
optional_columns = [NOT_USED] + all_columns

suggested_specimen_id = suggest_column(
    all_columns,
    [
        "Specimen ID",
        "Catalogue number",
        "Catalog number",
        "Accession number",
        "Specimen number",
    ],
)
suggested_latitude = suggest_column(
    all_columns,
    ["Latitude", "Decimal latitude", "Lat"],
)
suggested_longitude = suggest_column(
    all_columns,
    ["Longitude", "Decimal longitude", "Lon", "Lng"],
)
suggested_altitude = suggest_column(
    all_columns,
    ["Elevation (m)", "Elevation", "Altitude (m)", "Altitude"],
)
suggested_date = suggest_column(
    all_columns,
    ["Collection date", "Date collected", "Event date", "Date"],
)
suggested_collector = suggest_column(
    all_columns,
    ["Collector", "Collectors", "Collected by", "Recorded by"],
)
suggested_identifier = suggest_column(
    all_columns,
    ["Determined by", "Identified by", "Determiner", "Identifier"],
)

suggested_locality_columns = []
for aliases in [
    ["Country"],
    ["Region", "State", "Province", "Canton", "County", "District"],
    ["Locality", "Site", "Location", "Verbatim locality"],
]:
    suggestion = suggest_column(all_columns, aliases)
    if (
        suggestion != NOT_USED
        and suggestion not in suggested_locality_columns
    ):
        suggested_locality_columns.append(suggestion)

suggested_scientific_name_columns = []
combined_taxon_column = suggest_column(
    all_columns,
    ["Scientific name", "Taxon", "Taxon name"],
)

if combined_taxon_column != NOT_USED:
    suggested_scientific_name_columns = [combined_taxon_column]
else:
    for aliases in [
        ["Genus"],
        ["Qualifier", "Identification qualifier"],
        ["Species", "Specific epithet"],
    ]:
        suggestion = suggest_column(all_columns, aliases)
        if (
            suggestion != NOT_USED
            and suggestion not in suggested_scientific_name_columns
        ):
            suggested_scientific_name_columns.append(suggestion)

st.subheader("4. Match Excel columns")
st.caption(
    "Suggested mappings are selected automatically from common biodiversity "
    "headings. Please check them before exporting."
)

mapping_left, mapping_middle, mapping_right = st.columns(3)


with mapping_left:
    specimen_id_column = st.selectbox(
        "Specimen ID — optional",
        optional_columns,
        index=optional_columns.index(suggested_specimen_id),
        help=(
            "Choose the catalogue or specimen number column. "
            "Its position on the label can be configured below."
        ),
    )

    specimen_id_placement = st.selectbox(
        "Specimen ID placement",
        options=[
            "Compact — before first content",
            "Compact — after date",
            "Compact — after collector",
            "Separate first line",
            "Separate last line",
            "Do not print",
        ],
        index=0,
        disabled=specimen_id_column == NOT_USED,
        help=(
            "The default places the ID at the very beginning of the first "
            "content line, for example: ENT-0001 · Switzerland, Zurich. "
            "Compact placements save one line and fall back automatically "
            "when the preferred content field is missing."
        ),
    )

    locality_columns = st.multiselect(
        "Location — select one or several columns",
        all_columns,
        default=suggested_locality_columns,
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
        index=optional_columns.index(suggested_latitude),
    )

    longitude_column = st.selectbox(
        "Longitude — optional",
        optional_columns,
        index=optional_columns.index(suggested_longitude),
    )

    altitude_column = st.selectbox(
        "Altitude — optional",
        optional_columns,
        index=optional_columns.index(suggested_altitude),
    )


with mapping_middle:
    date_column = st.selectbox(
        "Collection date",
        optional_columns,
        index=optional_columns.index(suggested_date),
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
        index=optional_columns.index(suggested_collector),
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
        default=suggested_scientific_name_columns,
        help="Example: one Taxon column, or Genus + Qualifier + Species.",
    )

    identifier_column = st.selectbox(
        "Identifier / determiner — optional",
        optional_columns,
        index=optional_columns.index(suggested_identifier),
    )

    shorten_identifier_names = st.checkbox(
        "Shorten identifier first names",
        value=True,
    )


# =========================================================
# IDENTIFICATION SETTINGS
# =========================================================

st.subheader("5. Determination settings")

labels_to_create = st.radio(
    "Labels to create",
    options=[
        "Collection + determination labels",
        "Collection labels only",
        "Determination labels only",
    ],
    horizontal=True,
)

create_collection_label = labels_to_create in {
    "Collection + determination labels",
    "Collection labels only",
}
create_determination_label = labels_to_create in {
    "Collection + determination labels",
    "Determination labels only",
}

print_id_on_determination_label = st.checkbox(
    "Print specimen ID on determination label",
    value=True,
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
        value=1.00,
        step=0.05,
    )

collection_layout_settings = st.columns(3)

with collection_layout_settings[0]:
    collection_auto_enlarge = st.checkbox(
        "Use free space automatically",
        value=True,
        help=(
            "EntoLabel enlarges short labels up to the selected maximum, "
            "then shrinks only when necessary."
        ),
    )

with collection_layout_settings[1]:
    collection_max_font_size = st.number_input(
        "Maximum collection font, pt",
        min_value=3.0,
        max_value=12.0,
        value=6.5,
        step=0.25,
        disabled=not collection_auto_enlarge,
    )

with collection_layout_settings[2]:
    collection_vertical_alignment = st.selectbox(
        "Collection text position",
        options=["Balanced", "Top"],
        index=0,
        help=(
            "Balanced centres a short text block vertically instead of "
            "leaving it crowded into the top-left corner."
        ),
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
            value=1.00,
            step=0.05,
        )

    determination_layout_settings = st.columns(3)

    with determination_layout_settings[0]:
        determination_auto_enlarge = st.checkbox(
            "Use free space on determination labels",
            value=True,
        )

    with determination_layout_settings[1]:
        determination_max_font_size = st.number_input(
            "Maximum determination font, pt",
            min_value=3.0,
            max_value=12.0,
            value=6.5,
            step=0.25,
            disabled=not determination_auto_enlarge,
        )

    with determination_layout_settings[2]:
        determination_vertical_alignment = st.selectbox(
            "Determination text position",
            options=["Balanced", "Top"],
            index=0,
        )

else:
    determination_width_mm = collection_width_mm
    determination_height_mm = collection_height_mm
    determination_font_size = collection_font_size
    determination_line_spacing = collection_line_spacing
    determination_auto_enlarge = collection_auto_enlarge
    determination_max_font_size = collection_max_font_size
    determination_vertical_alignment = collection_vertical_alignment


draw_borders = st.checkbox(
    "Draw cutting borders",
    value=True,
)


# =========================================================
# BUILD PREVIEW LABELS
# =========================================================

st.subheader("7. Live preview")

preview_excel_row = st.number_input(
    "Preview Excel row",
    min_value=int(start_excel_row),
    max_value=int(end_excel_row),
    value=int(start_excel_row),
    step=1,
)

preview_row = dataframe.iloc[int(preview_excel_row) - int(start_excel_row)]

collection_preview_lines = build_collection_lines(
    row=preview_row,
    specimen_id_column=specimen_id_column,
    specimen_id_placement=specimen_id_placement,
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

if create_collection_label and create_determination_label:
    preview_columns = st.columns(2)

    with preview_columns[0]:
        render_live_label(
            title="Collection label",
            lines=collection_preview_lines,
            width_mm=collection_width_mm,
            height_mm=collection_height_mm,
            font_size_pt=collection_font_size,
            line_spacing=collection_line_spacing,
            automatically_enlarge=collection_auto_enlarge,
            maximum_font_size_pt=collection_max_font_size,
            vertical_alignment=collection_vertical_alignment,
        )

    with preview_columns[1]:
        render_live_label(
            title="Determination label",
            lines=determination_preview_lines,
            width_mm=determination_width_mm,
            height_mm=determination_height_mm,
            font_size_pt=determination_font_size,
            line_spacing=determination_line_spacing,
            automatically_enlarge=determination_auto_enlarge,
            maximum_font_size_pt=determination_max_font_size,
            vertical_alignment=determination_vertical_alignment,
        )

elif create_collection_label:
    render_live_label(
        title="Collection label",
        lines=collection_preview_lines,
        width_mm=collection_width_mm,
        height_mm=collection_height_mm,
        font_size_pt=collection_font_size,
        line_spacing=collection_line_spacing,
        automatically_enlarge=collection_auto_enlarge,
        maximum_font_size_pt=collection_max_font_size,
        vertical_alignment=collection_vertical_alignment,
    )

else:
    render_live_label(
        title="Determination label",
        lines=determination_preview_lines,
        width_mm=determination_width_mm,
        height_mm=determination_height_mm,
        font_size_pt=determination_font_size,
        line_spacing=determination_line_spacing,
        automatically_enlarge=determination_auto_enlarge,
        maximum_font_size_pt=determination_max_font_size,
        vertical_alignment=determination_vertical_alignment,
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
    automatically_enlarge: bool,
    maximum_font_size: float,
    vertical_alignment: str,
) -> bool:
    """Draw one compact label and return True if all text fits."""

    # A smaller padding is still safe for cutting, but makes much better use
    # of tiny entomological labels.
    inner_padding = 0.45 * mm

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
            automatically_enlarge=automatically_enlarge,
            maximum_font_size=maximum_font_size,
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
    text_block_height = (
        actual_font_size
        + max(len(prepared_lines) - 1, 0) * line_height
    )

    if vertical_alignment == "Balanced":
        text_y = (
            y
            + inner_padding
            + ((available_height + text_block_height) / 2)
            - actual_font_size
        )
    else:
        text_y = (
            y
            + label_height
            - inner_padding
            - actual_font_size
        )

    for line in prepared_lines:
        if text_y < y + inner_padding - 0.01:
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
        if create_collection_label:
            collection_lines = build_collection_lines(
                row=row,
                specimen_id_column=specimen_id_column,
                specimen_id_placement=specimen_id_placement,
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
            )

            label_jobs.append(
                {
                    "lines": collection_lines,
                    "width": collection_width_mm * mm,
                    "height": collection_height_mm * mm,
                    "font_size": collection_font_size,
                    "line_spacing": collection_line_spacing,
                    "auto_enlarge": collection_auto_enlarge,
                    "max_font_size": collection_max_font_size,
                    "vertical_alignment": collection_vertical_alignment,
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
                    "auto_enlarge": determination_auto_enlarge,
                    "max_font_size": determination_max_font_size,
                    "vertical_alignment": determination_vertical_alignment,
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
            automatically_enlarge=job["auto_enlarge"],
            maximum_font_size=job["max_font_size"],
            vertical_alignment=job["vertical_alignment"],
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

if create_collection_label and not locality_columns:
    st.warning(
        "Select at least one location column."
    )
    configuration_is_valid = False

if create_collection_label and date_column == NOT_USED:
    st.warning(
        "No collection-date column is selected."
    )

if create_collection_label and collector_column == NOT_USED:
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

    if labels_to_create == "Determination labels only":
        pdf_filename = "determination_labels.pdf"
    elif labels_to_create == "Collection labels only":
        pdf_filename = "collection_labels.pdf"
    else:
        pdf_filename = "entomology_labels.pdf"

    st.download_button(
        label="📄 Create A4 PDF",
        data=pdf_bytes,
        file_name=pdf_filename,
        mime="application/pdf",
    )
