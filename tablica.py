from dataclasses import dataclass

import pandas as pd


@dataclass
class Tablica:
    imya: str
    df: pd.DataFrame
