"""Geometric law contract: endpoints, monotonicity and drive convention."""
import importlib.util,unittest
from pathlib import Path
path=Path(__file__).resolve().parents[1]/'skills/optomech-design/scripts/motion_law.py'
spec=importlib.util.spec_from_file_location('focus_law',path);law=importlib.util.module_from_spec(spec);spec.loader.exec_module(law)

class FocusLawTests(unittest.TestCase):
    def test_reversible_geometric_travel_and_external_gearing(self):
        rows=list(law.samples(-40,40,-2,2,21,'smoothstep',-5))
        self.assertEqual(rows[0],[-40,-2,200]);self.assertEqual(rows[-1],[40,2,-200]);self.assertEqual(rows[10],[0,0,0])
        self.assertTrue(all(a[1]<b[1] for a,b in zip(rows,rows[1:])))
        for a,b in zip(rows,rows[::-1]):self.assertAlmostEqual(a[1],-b[1])
        self.assertLess(rows[1][1]-rows[0][1],rows[10][1]-rows[9][1])
    def test_invalid_motion_specification_is_rejected(self):
        for args in [(0,0,0,1,21,'linear'),(0,40,0,1,1,'linear'),(0,float('inf'),0,1,21,'linear')]:
            with self.assertRaises(ValueError):list(law.samples(*args))

if __name__=='__main__':unittest.main()
