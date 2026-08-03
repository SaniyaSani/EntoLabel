# EntoLabel - compact label layout update

This build improves both onboarding and PDF label use of space.

## Main changes

- Specimen ID placement is configurable:
  - Compact - before first content (new default; e.g. `ENT-0001 · Switzerland, Zurich`)
  - Compact - after date
  - Compact - after collector
  - Separate first line
  - Separate last line
  - Do not print
- Short labels can automatically use a larger font up to a chosen maximum.
- Text can be vertically balanced instead of staying in the top-left corner.
- Inner label padding is reduced from 0.75 mm to 0.45 mm.
- Default line spacing is slightly tighter.
- The live preview uses the same wrapping and font-fitting logic as the PDF.
- Common spreadsheet columns are suggested automatically.
- A likely header row is detected automatically.
- Empty and fictional CSV templates can be downloaded from the start screen.

## Run locally

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

## Deploy

For Streamlit Community Cloud, place `app.py` and `requirements.txt` in the repository root and select `app.py` as the entry point.

## Demo data

`EntoLabel_fictional_clean_demo.xlsx` contains fictional people, localities, coordinates and specimen records. Scientific names are real only to make the labels look realistic.
