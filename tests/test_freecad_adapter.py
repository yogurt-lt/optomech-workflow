"""Pure-Python adapter checks. These do not execute the FreeCAD kernel."""
import ast,copy,json,unittest
from optomech.freecad_generator import generate_freecad_macro

def sleeve():
    return dict(title='Concept "quotes" / 中文',source_origin_z_mm=-12.5,parts=[dict(name='01_Sleeve',profile_zr_mm=[[0,10],[20,10],[20,7],[0,7]],color_rgb=[.18,.31,.55])],lenses=[])

def embedded(text):
    node=next(n for n in ast.parse(text).body if isinstance(n,ast.Assign) and any(isinstance(t,ast.Name) and t.id=='PLAN' for t in n.targets))
    return json.loads(ast.literal_eval(node.value.args[0]))

class FreeCADAdapterTests(unittest.TestCase):
    def test_serialization_preserves_plan_and_source_offset(self):
        p=sleeve();before=copy.deepcopy(p);text,report=generate_freecad_macro(p);data=embedded(text)
        self.assertEqual(data['title'],p['title']);self.assertEqual(data['source_origin_z_mm'],-12.5)
        self.assertEqual(data['parts'][0]['profile_zr_mm'],p['parts'][0]['profile_zr_mm'])
        self.assertEqual(p,before);self.assertEqual(report['runtime_validation'],'pending in FreeCAD')
    def test_dedicated_cut_is_not_silently_lost(self):
        p=sleeve();p['radial_features']=[dict(part='01_Sleeve',zs_mm=[10],diameter_mm=2)]
        with self.assertRaisesRegex(ValueError,'dedicated cuts'):generate_freecad_macro(p)
        text,report=generate_freecad_macro(p,allow_base_only=True)
        self.assertTrue(report['extra_features_not_implemented']);self.assertTrue(embedded(text)['capabilities']['baseline_only'])
    def test_invalid_data_cannot_be_emitted_as_code(self):
        for value in [float('nan'),float('inf'),True,'meters']:
            p=sleeve();p['source_origin_z_mm']=value
            with self.assertRaises(ValueError):generate_freecad_macro(p)
        p=sleeve();p['parts'][0]['color_rgb']=['0;import os',0,0]
        with self.assertRaises(ValueError):generate_freecad_macro(p)
    def test_glass_request_does_not_claim_unavailable_bodies(self):
        _,report=generate_freecad_macro(sleeve(),include_blanks=True)
        self.assertTrue(report['requested_glass']);self.assertFalse(report['included_glass']);self.assertEqual(report['included_glass_count'],0)

    def test_explicit_radial_holes_and_slots_are_serialized(self):
        p=sleeve()
        p['radial_features']=[dict(part='01_Sleeve',zs_mm=[10],diameter_mm=2.5,angles_degrees=[0,120,240])]
        p['parts'][0].update(cap_slot='front',wrench_slots=dict(width_mm=2.,radial_depth_mm=2.,axial_depth_mm=1.5))
        text,report=generate_freecad_macro(p)
        cuts=embedded(text)['parts'][0]['cuts']
        self.assertEqual([c['kind'] for c in cuts],['radial_holes','wrench_slots'])
        self.assertEqual(cuts[0]['angles_degrees'],[0,120,240])
        self.assertEqual(cuts[1]['axial_depth_mm'],1.5)
        self.assertFalse(report['extra_features_not_implemented'])
        self.assertIn('PartDesign::Pocket',text)

    def test_invalid_native_cut_dimensions_rejected(self):
        for value in [-1,0,True,float('nan')]:
            p=sleeve()
            p['parts'][0].update(cap_slot='front',wrench_slots=dict(width_mm=value,radial_depth_mm=2.,axial_depth_mm=1.5))
            with self.assertRaisesRegex(ValueError,'wrench slot'):generate_freecad_macro(p)

    def test_radial_windows_preserve_axial_bounds_and_angles(self):
        p=sleeve();spec=dict(width_mm=2.,z_start_mm=-.05,z_end_mm=1.5,angles_degrees=[30,150,270])
        p['parts'][0]['radial_windows']=[spec]
        text,report=generate_freecad_macro(p)
        cut=embedded(text)['parts'][0]['cuts'][0]
        self.assertEqual(cut,dict(kind='radial_windows',**spec))
        self.assertFalse(report['extra_features_not_implemented'])
        for bad in [dict(spec,width_mm=0),dict(spec,z_end_mm=-1),dict(spec,angles_degrees=[float('nan')])]:
            p['parts'][0]['radial_windows']=[bad]
            with self.assertRaisesRegex(ValueError,'radial window'):generate_freecad_macro(p)

if __name__=='__main__':unittest.main()
