from dataclasses import dataclass
from typing import Optional

@dataclass
class GenerationRequest:
    template_id: str
    year: int
    locale: str = "pt_BR"
    currency: str = "BRL"
    theme: str = "light"
    with_sample_data: bool = False

    # Domain specific config
    profile: str = "couple"
    reserve_months: int = 6
    projection_horizon: int = 12
