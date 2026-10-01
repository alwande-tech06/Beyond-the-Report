"""
Text analytics for the sustainability / ESG reports (APFA802 - Beyond the Report).

Counts how often each ESG theme and each type of language (commitment, performance, risk) appears in a report,
and keeps one example sentence with its page number so every signal can be traced back to the document.

Run:  python nlp.py            scans the report PDFs one folder up and writes data/report_themes.csv
"""
from __future__ import annotations

import io
import re
from pathlib import Path

import pandas as pd

OUT = Path(__file__).parent / "data" / "report_themes.csv"
REPORT_DIR = Path(__file__).parent.parent

# file name -> (company, report title)
REPORTS = {
    "sustainability-report-2025_ Valterra Platinum.pdf": ("Valterra Platinum", "Valterra Platinum Sustainability report 2025"),
    "Impala-report-2025.pdf": ("Impala Platinum", "Implats ESG Report 2025"),
    "SSW-SR25.pdf": ("Sibanye-Stillwater", "Sibanye-Stillwater Sustainability Report 2025"),
    "sustainability-report-2025-interactive_GOLDFIELD.pdf": ("Gold Fields", "Gold Fields Sustainability Report 2025"),
    "exx-2025-esg-report.pdf": ("Exxaro Resources", "Exxaro ESG report 2025"),
}

# theme -> (group, regex). Word stems are matched case-insensitively on word boundaries.
THEMES = {
    "Climate and emissions": ("Topic", r"climate|decarboni[sz]\w*|emissions?|ghg|greenhouse|carbon|net[- ]zero|scope [123]"),
    "Energy": ("Topic", r"energy|renewable\w*|electricity|solar|wind power|diesel"),
    "Water": ("Topic", r"water|effluent|recycl\w+ water|withdrawal"),
    "Tailings and waste": ("Topic", r"tailings|tsf\b|gistm|waste"),
    "Biodiversity and closure": ("Topic", r"biodiversity|rehabilitat\w+|mine closure|closure plan\w*|land management"),
    "Safety and health": ("Topic", r"safety|fatalit\w+|injur\w+|ltifr|trifr|occupational health|silicosis"),
    "Workforce": ("Topic", r"employees?|workforce|training|skills development|diversity|women|gender"),
    "Community": ("Topic", r"communit\w+|socio[- ]economic|social and labour|slp\b|csi\b|host communit\w+"),
    "Human rights": ("Topic", r"human rights|grievance\w*|modern slavery|indigenous"),
    "Governance and ethics": ("Topic", r"governance|board|ethic\w+|anti[- ]corruption|whistle[- ]?blow\w+|compliance"),
    "Assurance": ("Topic", r"assurance|assured|isae ?3000|aa1000|independent(?:ly)? (?:assur|verif)\w+"),
    "Financial connectivity": ("Topic", r"capex|capital expenditure|(?:financial|closure|rehabilitation|environmental) provisions?|provisions? for|impairment\w*|cost of capital|cash flows?|carbon tax"),
    "Commitment language": ("Language", r"will|aims?|aiming|commit\w*|targets?|plans? to|intend\w*|ambition\w*|by 20[3-5]0"),
    "Performance language": ("Language", r"achieved|reduced|increased|decreased|improved|delivered|recorded|completed|performance"),
    "Risk language": ("Language", r"risks?|uncertaint\w+|exposure|threat\w*|vulnerab\w+|challeng\w+"),
}
_RX = {t: re.compile(rf"\b(?:{rx})\b", re.I) for t, (_, rx) in THEMES.items()}
_SENT = re.compile(r"(?<=[.!?])\s+")


def extract_pages(source) -> list[str]:
    """Text of each page. `source` is a path or the bytes of a PDF. Image-only pages come back empty."""
    from pypdf import PdfReader  # imported here so the dashboard still starts if pypdf is missing

    reader = PdfReader(io.BytesIO(source) if isinstance(source, (bytes, bytearray)) else str(source))
    pages = []
    for p in reader.pages:
        try:
            pages.append(re.sub(r"\s+", " ", p.extract_text() or "").strip())
        except Exception:  # noqa: BLE001 - one unreadable page should not lose the report
            pages.append("")
    return pages


def _snippet(text: str, m: re.Match) -> str:
    """The sentence around a match, trimmed for display."""
    dot = text.rfind(". ", 0, m.start())
    start = dot + 2 if dot != -1 else 0
    if m.start() - start > 160:  # long run-on text: start at a word boundary instead of mid-word
        start = text.find(" ", m.start() - 160) + 1
    end = text.find(". ", m.end())
    end = end + 1 if end != -1 else len(text)
    if end - m.end() > 200:
        end = text.rfind(" ", m.end(), m.end() + 200)
    return text[start:end].strip()


def theme_counts(pages: list[str]) -> pd.DataFrame:
    """One row per theme: mentions, pages mentioning it, mentions per 100 pages and an example with its PDF page."""
    n_pages = max(len(pages), 1)
    words = sum(len(p.split()) for p in pages)
    rows = []
    for theme, (group, _) in THEMES.items():
        rx, mentions, on_pages, best = _RX[theme], 0, 0, (0, None, "")
        for i, text in enumerate(pages, start=1):
            found = list(rx.finditer(text))
            if not found:
                continue
            mentions += len(found)
            on_pages += 1
            if len(found) > best[0]:  # example comes from the page that talks about the theme most
                digits = [m for m in found if re.search(r"\d", _snippet(text, m))]
                best = (len(found), i, _snippet(text, (digits or found)[0]))
        rows.append(dict(theme=theme, group=group, mentions=mentions, pages_mentioning=on_pages,
                         per_100_pages=round(mentions / n_pages * 100, 1),
                         per_10k_words=round(mentions / max(words, 1) * 10000, 1),
                         example_page=best[1], example=best[2]))
    df = pd.DataFrame(rows)
    df["report_pages"] = len(pages)
    df["report_words"] = words
    return df


def scan_folder(folder: Path = REPORT_DIR) -> pd.DataFrame:
    parts = []
    for name, (company, title) in REPORTS.items():
        path = folder / name
        if not path.exists():
            print(f"  missing: {name}")
            continue
        print(f"  reading {name} ...", flush=True)
        df = theme_counts(extract_pages(path))
        df.insert(0, "company", company)
        df.insert(1, "source_document", title)
        parts.append(df)
    return pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()


def load(path: Path = OUT) -> pd.DataFrame:
    return pd.read_csv(path) if path.exists() else pd.DataFrame()


if __name__ == "__main__":
    out = scan_folder()
    if out.empty:
        print("No report PDFs found in", REPORT_DIR)
    else:
        out.to_csv(OUT, index=False, encoding="utf-8")
        print(out.groupby("company")[["report_pages", "report_words"]].first().to_string())
        print("Wrote", OUT)
