import importlib.util
import sys
import types
import unittest
from pathlib import Path

# The rules are pure Python; stub optional conversion dependencies so the unit
# tests can run on a machine where PDF/Excel packages are not installed.
if "pdfplumber" not in sys.modules:
    sys.modules["pdfplumber"] = types.ModuleType("pdfplumber")
if "openpyxl" not in sys.modules:
    openpyxl = types.ModuleType("openpyxl")
    openpyxl.Workbook = object
    openpyxl.load_workbook = lambda *args, **kwargs: None
    openpyxl.styles = types.ModuleType("openpyxl.styles")
    sys.modules["openpyxl"] = openpyxl
    sys.modules["openpyxl.styles"] = openpyxl.styles

ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = [
    ROOT / "PAIEMENT CLIENT" / "pdf_paiement_to_excel.py",
    ROOT / "PAIEMENT CLIENT" / "BSA" / "BSA_paiement_to_excel.py",
]


def load(path):
    spec = importlib.util.spec_from_file_location(path.stem + str(len(str(path))), path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class BsaRulesTest(unittest.TestCase):
    def test_business_rules_in_both_entry_points(self):
        cases = [
            ((0, 10_000, 10_000, 500, 500), (10_000, 0, 10_000, 0)),
            ((95, 10_000, 10_000, 0, 0), (10_000, 0, 10_000, 0)),
            ((95, 10_000, 9_500, 500, 0), (10_000, 500, 9_500, 0)),
            # Rule 4 must win over rule 3.
            ((95, 10_000, 0, 10_000, 0), (10_000, 0, 0, 10_000)),
        ]
        keys = ("Montant_Reclame_Brut", "Ticket_Moderateur",
                "Montant_Paye_Regle", "Montant_Exclu_Rejet")
        for script in SCRIPTS:
            fn = load(script).calcul_montants_bsa
            for inputs, expected in cases:
                with self.subTest(script=script.name, inputs=inputs):
                    result = fn(*inputs)
                    self.assertEqual(tuple(result[k] for k in keys), expected)

    def test_fallback_preserves_bsa_amounts(self):
        for script in SCRIPTS:
            result = load(script).calcul_montants_bsa(0, 10_000, 8_000, 2_000, 1_000)
            self.assertEqual(result["Ticket_Moderateur"], 1_000)
            self.assertEqual(result["Montant_Exclu_Rejet"], 2_000)


if __name__ == "__main__":
    unittest.main()
