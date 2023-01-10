from dataclasses import dataclass

@dataclass
class Erzeuger:
    name: str
    material: str
    plz: str
    ort: str
    avv: str
    menge: str
    