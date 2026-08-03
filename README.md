# EntoLabel compact layout update

This version improves tiny collection-label layouts:

- **Specimen ID first, inline** is the default: `A1 · Locality`.
- **Latitude and longitude stay together on one printed line**.
- Altitude, collection date and collector are packed into a flowing metadata line.
- Wrapped continuation lines do not begin with an orphaned middle dot.
- Short labels may enlarge automatically; long labels shrink only as needed.
- **Print specimen ID on determination label** is enabled by default.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

The included XLSX file is a fully fictional demonstration dataset.
