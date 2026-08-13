from dataclasses import dataclass

@dataclass
class ThemeColors:
    background: str
    surface: str
    surface_secondary: str

    text_primary: str
    text_secondary: str
    text_muted: str

    accent: str

    positive: str
    warning: str
    negative: str
    info: str

    input_background: str
    input_border: str

    border: str

class Theme:
    name: str
    colors: ThemeColors
    font_family: str = "Arial"

    # Common table styles available in XlsxWriter
    table_style: str
