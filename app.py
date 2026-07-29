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
    if uploaded_file.name.lower().endswith(".csv"):
        dataframe = pd.read_csv(uploaded_file)
    else:
        dataframe = pd.read_excel(uploaded_file)
except Exception as error:
    st.error(f"Could not read the file: {error}")
    st.stop()


if dataframe.empty:
    st.warning("The uploaded file is empty.")
    st.stop()


# =========================================================
# EXCEL PREVIEW
# =========================================================

st.subheader("1. Excel preview")

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
# COLUMN MAPPING
# =========================================================

all_columns = dataframe.columns.tolist()
optional_columns = [NOT_USED] + all_columns

st.subheader("2. Match Excel columns")

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


# =========================================================
# IDENTIFICATION SETTINGS
# =========================================================

st.subheader("3. Determination settings")

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

st.subheader("4. Label size and typography")

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

st.subheader("5. Live preview")

preview_row_number = st.number_input(
    "Preview Excel row",
    min_value=1,
    max_value=len(dataframe),
    value=1,
    step=1,
)

preview_row = dataframe.iloc[int(preview_row_number) - 1]

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

st.subheader("6. Export")

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
