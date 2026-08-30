import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "PAIEMENT CLIENT" / "BSA" / "BSA_excel_to_modele.py"

# Les scripts de conversion ont des dépendances optionnelles pour les tests
# métier existants. Les tests de relevés sont donc ignorés si les bibliothèques
# Excel/PDF ne sont pas installées dans l'environnement courant.
try:
    import openpyxl  # type: ignore
    import pdfplumber  # type: ignore

    HAS_CONVERSION_DEPS = bool(getattr(openpyxl, "__file__", None))
except ImportError:
    HAS_CONVERSION_DEPS = False


def load_converter():
    spec = importlib.util.spec_from_file_location("bsa_excel_to_modele_test", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


@unittest.skipUnless(
    HAS_CONVERSION_DEPS,
    "openpyxl et pdfplumber sont nécessaires pour tester les relevés Excel",
)
class BsaExcelParserTest(unittest.TestCase):
    def test_releve_afg_est_converti_en_quatre_lignes(self):
        converter = load_converter()
        source = ROOT / "PAIEMENT CLIENT" / "BSA" / "Excel" / (
            "14-08-2026 AFG ASSURANCES 1050265 SALFATU LOT94_082026.xlsx"
        )
        meta, lignes = converter.parse_source(source)
        self.assertEqual(meta["ref"], "1050265")
        self.assertEqual(meta["date_reglement"], "14/08/2026")
        self.assertEqual(len(lignes), 4)
        self.assertEqual(sum(l["Montant_Paye_Regle"] for l in lignes), 38_240)

    def test_releve_orange_est_converti_en_84_lignes(self):
        converter = load_converter()
        source = ROOT / "PAIEMENT CLIENT" / "BSA" / "Excel" / (
            "20-08-2026 ORANGE 1130623 SALFATU LOT103_082026.xlsx"
        )
        meta, lignes = converter.parse_source(source)
        self.assertEqual(meta["ref"], "1130623")
        self.assertEqual(meta["date_reglement"], "20/08/2026")
        self.assertEqual(len(lignes), 84)
        self.assertEqual(sum(l["Montant_Paye_Regle"] for l in lignes), 1_226_080)


if __name__ == "__main__":
    unittest.main()
