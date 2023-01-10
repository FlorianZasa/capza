from dataclasses import dataclass

@dataclass
class LA:
    datum: str
    material_bez: str
    material_kenn: str
    wassergehalt: int
    einwaage_fs: int
    intauswaage_fs: int
    ts_der_probe: int
    result_ts: int
    result_wasserfaktor: float
    result_wasserfaktor_getrocknet: float
    result_wasserfaktor_getrocknet: float

    def calculate_ts(self, feuchte):
        return 100-float(feuchte)

    def calculate_lipos_ts(self, auswaage, tara, einwaage_sox_frisch):
        return (float(auswaage) - (float(tara)))/(float(einwaage_sox_frisch) / 100)

    def calculate_gv(self, gv_auswaage, gv_tara, gv_einwaage):
        return 100 - (float(gv_auswaage)) - float(gv_tara) / float(gv_einwaage) * 100

    def calculate_tds(self, tds_auswaage, tds_tara, tds_einwaage):
        return 100 - (float(tds_auswaage)) - float(tds_tara) / float(tds_einwaage) * 100
