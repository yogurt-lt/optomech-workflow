import copy,json,re,unittest
from pathlib import Path
from optomech.generator import generate_macro,blank_references
from optomech.geometry import validate_plan,self_intersects
from optomech.prescription import parse_zmx,step_metadata,optical_groups

ROOT=Path(__file__).resolve().parents[1]

class WorkflowTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.fixed=json.loads((ROOT/'examples/fixed-projection/plan.json').read_text())
        cls.film=json.loads((ROOT/'examples/film-projection/plan.json').read_text())
    def test_demo_geometry(self):
        for p in (self.fixed,self.film):self.assertEqual(validate_plan(p)['errors'],[])
    def test_glued_groups_and_known_stop(self):
        s=parse_zmx((ROOT/'examples/fixed-projection/source.zmx').read_bytes())
        meta=step_metadata((ROOT/'examples/fixed-projection/source.STP').read_text())
        groups=optical_groups(s,meta['surface_pairs'])
        self.assertEqual(len(groups),8);self.assertEqual(groups[4],[10,11,12]);self.assertTrue(s[13]['stop'])
    def test_generator_preserves_save_and_loading_fixes(self):
        macro=generate_macro(self.film)
        self.assertEqual(macro.count('0, SAVE_SILENT)'),2)
        self.assertIn('Dim loaded(6)',macro)
        self.assertIn('component.Transform2 = move',macro)
        self.assertIn('StrComp(model.GetPathName, target, vbTextCompare)',macro)
        self.assertIn('-1017.',macro)
    def test_crossed_profile_is_blocked(self):
        p=copy.deepcopy(self.film);p['parts'][1]['profile_zr_mm']=[[0,1],[1,2],[0,2],[1,1]]
        self.assertTrue(validate_plan(p)['errors'])
        with self.assertRaises(ValueError):generate_macro(p)
    def test_mechanical_collision_is_blocked(self):
        p=copy.deepcopy(self.film);p['parts'][1]['profile_zr_mm']=copy.deepcopy(p['parts'][0]['profile_zr_mm'])
        self.assertTrue(any('Mechanical overlap' in s for s in validate_plan(p,full=False)['errors']))
    def test_blank_references_are_explicit(self):
        self.assertEqual(blank_references(self.film),[])
        refs=blank_references(self.fixed);self.assertEqual(len(refs),10)
        self.assertTrue(all(not self_intersects(p['profile_zr_mm']) for p in refs))
        macro=generate_macro(self.fixed,include_blanks=True);self.assertIn('Dim loaded(15)',macro)
        self.assertIn('Lens3_RevisedBlank_Reference',macro)
        self.assertIn('AddAdhesivePorts model, plane, axis',macro)
        self.assertIn('FeatureCut3(True, False, reverseCut',macro)
    def test_no_machine_paths_in_sources(self):
        for p in (ROOT/'examples').rglob('source.*'):
            text=p.read_text(encoding='utf-8')
            self.assertNotIn('C:\\Users',text);self.assertNotIn('C:/Users',text)

    def test_port_variable_does_not_use_vba_reserved_circle(self):
        macro=generate_macro(self.fixed,include_blanks=True)
        self.assertIn('Dim sketch As SldWorks.Sketch, portCircle As SldWorks.SketchSegment',macro)
        self.assertIn('Set portCircle = model.SketchManager.CreateCircleByRadius',macro)
        self.assertIsNone(re.search(r'\b(?:Dim|Set|,)\s+circle\b',macro,re.I))

    def test_ports_use_individual_closed_nonconstruction_sketches(self):
        macro=generate_macro(self.fixed)
        ports=macro[macro.index('Private Sub AddAdhesivePorts'):macro.index('Private Sub ApplyConceptColor')]
        row_loop=ports.index('For row = 0 To 2')
        circle=ports.index('Set portCircle =')
        cut=ports.index('Set cut =')
        end_loop=ports.index('Next row')
        self.assertTrue(row_loop < circle < cut < end_loop)
        self.assertIn('portCircle.ConstructionGeometry = False',ports)
        self.assertIn('FeatureCut3(True, False, reverseCut, 1, 0,',ports)
        self.assertIn('False, False, False, False, False, False, True, False, False, False, 0, 0#, False)',ports)
        self.assertIn('Abs(p(2))',ports)

    def test_port_math_point_uses_double_safearray_and_readback(self):
        macro=generate_macro(self.fixed)
        ports=macro[macro.index('Private Sub AddAdhesivePorts'):macro.index('Private Sub ApplyConceptColor')]
        self.assertIn('Dim position(2) As Double',ports)
        self.assertNotIn('position = Array(',ports)
        self.assertIn('position(2) = CDbl(zs(row))',ports)
        self.assertIn('inputCheck = mp.ArrayData',ports)
        self.assertIn('inputCheck(2) - position(2)',ports)
        self.assertIn('Port transformed coordinate radius mismatch.',ports)

    def test_unmodified_lens_references_do_not_require_blank_changes(self):
        plan=copy.deepcopy(self.fixed)
        plan['blank_revisions']=[]
        refs=blank_references(plan)
        self.assertEqual(len(refs),10)
        self.assertTrue(all('Optical_Reference' in r['name'] for r in refs))
        macro=generate_macro(plan,include_blanks=True)
        self.assertIn('Dim loaded(15)',macro)

if __name__=='__main__':unittest.main()
