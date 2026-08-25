import unittest
from .utils import balance_chemical_equation

class TestChemicalBalancer(unittest.TestCase):
    def test_coefficient_parsing(self):
        # 1. Combination / Synthesis
        self.assertEqual(balance_chemical_equation("H2 + O2 = H2O"), "2H2 + O2 = 2H2O")
        self.assertEqual(balance_chemical_equation("2H2 + O2 = 2H2O"), "2H2 + O2 = 2H2O")
        
        # 2. Decomposition
        self.assertEqual(balance_chemical_equation("H2O = H2 + O2"), "2H2O = 2H2 + O2")
        
        # 3. Single Displacement
        self.assertEqual(balance_chemical_equation("Zn + HCl = ZnCl2 + H2"), "Zn + 2HCl = ZnCl2 + H2")
        
        # 4. Double Displacement
        self.assertEqual(balance_chemical_equation("AgNO3 + NaCl = AgCl + NaNO3"), "AgNO3 + NaCl = AgCl + NaNO3")
        
        # 5. Combustion
        self.assertEqual(balance_chemical_equation("CH4 + O2 = CO2 + H2O"), "CH4 + 2O2 = CO2 + 2H2O")
        
        # 6. Acid + Base / Neutralization
        self.assertEqual(balance_chemical_equation("H2SO4 + NaOH = Na2SO4 + H2O"), "H2SO4 + 2NaOH = Na2SO4 + 2H2O")
        
        # 7. Redox
        self.assertEqual(balance_chemical_equation("Fe + O2 = Fe2O3"), "4Fe + 3O2 = 2Fe2O3")
        self.assertEqual(balance_chemical_equation("4Fe + 3O2 = 2Fe2O3"), "4Fe + 3O2 = 2Fe2O3")
        
        # 8. Metal + Acid
        self.assertEqual(balance_chemical_equation("Mg + HCl = MgCl2 + H2"), "Mg + 2HCl = MgCl2 + H2")
        
        # 9. Metal + Water
        self.assertEqual(balance_chemical_equation("Na + H2O = NaOH + H2"), "2Na + 2H2O = 2NaOH + H2")
        
        # 10. Precipitation
        self.assertEqual(balance_chemical_equation("BaCl2 + Na2SO4 = BaSO4 + NaCl"), "BaCl2 + Na2SO4 = BaSO4 + 2NaCl")

    def test_parsing_messy_coefficients(self):
        # Even if they put 100 before H2O, it should just ignore it and balance correctly
        self.assertEqual(balance_chemical_equation("100H2 + 5O2 = 12H2O"), "2H2 + O2 = 2H2O")
        
    def test_invalid_input(self):
        with self.assertRaisesRegex(ValueError, "Invalid chemical equation"):
            balance_chemical_equation("H2 + O2 = invalid")
            
        with self.assertRaisesRegex(ValueError, "Invalid chemical equation"):
            balance_chemical_equation("Not an equation")

if __name__ == '__main__':
    unittest.main()
